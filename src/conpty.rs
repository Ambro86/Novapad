use crate::{accessibility::to_wide, log_debug};
use std::ffi::c_void;
use std::mem::size_of;
use windows::Win32::Foundation::{CloseHandle, HANDLE};
use windows::Win32::Storage::FileSystem::WriteFile;
use windows::Win32::System::Console::{
    COORD, CTRL_C_EVENT, ClosePseudoConsole, CreatePseudoConsole, GenerateConsoleCtrlEvent, HPCON,
    ResizePseudoConsole,
};
use windows::Win32::System::Pipes::CreatePipe;
use windows::Win32::System::Threading::{
    CREATE_NEW_PROCESS_GROUP, CreateProcessW, DeleteProcThreadAttributeList,
    EXTENDED_STARTUPINFO_PRESENT, InitializeProcThreadAttributeList, LPPROC_THREAD_ATTRIBUTE_LIST,
    PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE, PROCESS_INFORMATION, STARTUPINFOEXW,
    UpdateProcThreadAttribute,
};
use windows::core::{PCWSTR, PWSTR};

pub struct ConPtySession {
    hpc: HPCON,
    input_write: HANDLE,
    process_handle: HANDLE,
    process_id: u32,
}

pub struct ConPtySpawn {
    pub session: ConPtySession,
    pub output_read: HANDLE,
}

impl ConPtySession {
    pub fn spawn(command: &str, cols: i16, rows: i16) -> windows::core::Result<ConPtySpawn> {
        let mut input_read = HANDLE(0);
        let mut input_write = HANDLE(0);
        unsafe {
            CreatePipe(&mut input_read, &mut input_write, None, 0)?;
        }

        log_debug(&format!(
            "ConPty: spawn() called with command: {}, cols: {}, rows: {}",
            command, cols, rows
        ));

        let mut output_read = HANDLE(0);
        let mut output_write = HANDLE(0);
        let output_pipe = unsafe { CreatePipe(&mut output_read, &mut output_write, None, 0) };
        if let Err(err) = output_pipe {
            unsafe {
                let _ = CloseHandle(input_read);
                let _ = CloseHandle(input_write);
            }
            return Err(err);
        }

        let hpc = unsafe {
            CreatePseudoConsole(COORD { X: cols, Y: rows }, input_read, output_write, 0)?
        };

        unsafe {
            let _ = CloseHandle(input_read);
            let _ = CloseHandle(output_write);
        }

        let mut attr_list_size: usize = 0;
        unsafe {
            let _ = InitializeProcThreadAttributeList(
                LPPROC_THREAD_ATTRIBUTE_LIST(std::ptr::null_mut()),
                1,
                0,
                &mut attr_list_size,
            );
        }
        let mut attr_list_buf = vec![0u8; attr_list_size];
        let attr_list = LPPROC_THREAD_ATTRIBUTE_LIST(attr_list_buf.as_mut_ptr() as *mut _);
        if let Err(err) =
            unsafe { InitializeProcThreadAttributeList(attr_list, 1, 0, &mut attr_list_size) }
        {
            unsafe {
                ClosePseudoConsole(hpc);
                let _ = CloseHandle(input_write);
                let _ = CloseHandle(output_read);
            }
            return Err(err);
        }

        if let Err(err) = unsafe {
            UpdateProcThreadAttribute(
                attr_list,
                0,
                PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE as usize,
                Some(hpc.0 as *const c_void),
                size_of::<HPCON>(),
                None,
                None,
            )
        } {
            unsafe {
                DeleteProcThreadAttributeList(attr_list);
                ClosePseudoConsole(hpc);
                let _ = CloseHandle(input_write);
                let _ = CloseHandle(output_read);
            }
            return Err(err);
        }

        let mut si_ex = STARTUPINFOEXW::default();
        si_ex.StartupInfo.cb = size_of::<STARTUPINFOEXW>() as u32;
        si_ex.lpAttributeList = attr_list;

        let mut cmdline = to_wide(command);
        let mut proc_info = PROCESS_INFORMATION::default();
        let flags = EXTENDED_STARTUPINFO_PRESENT | CREATE_NEW_PROCESS_GROUP;
        if let Err(err) = unsafe {
            CreateProcessW(
                PCWSTR::null(),
                PWSTR(cmdline.as_mut_ptr()),
                None,
                None,
                true,
                flags,
                None,
                None,
                &si_ex.StartupInfo,
                &mut proc_info,
            )
        } {
            unsafe {
                DeleteProcThreadAttributeList(attr_list);
                ClosePseudoConsole(hpc);
                let _ = CloseHandle(input_write);
                let _ = CloseHandle(output_read);
            }
            return Err(err);
        }

        unsafe {
            DeleteProcThreadAttributeList(attr_list);
        }

        unsafe {
            let _ = CloseHandle(proc_info.hThread);
        }

        Ok(ConPtySpawn {
            session: ConPtySession {
                hpc,
                input_write,
                process_handle: proc_info.hProcess,
                process_id: proc_info.dwProcessId,
            },
            output_read,
        })
    }

    pub fn write_input(&self, text: &str) -> bool {
        if self.input_write.0 == 0 {
            return false;
        }
        let bytes = text.as_bytes();
        let mut written = 0u32;
        unsafe { WriteFile(self.input_write, Some(bytes), Some(&mut written), None).is_ok() }
    }

    pub fn resize(&self, cols: i16, rows: i16) -> bool {
        unsafe { ResizePseudoConsole(self.hpc, COORD { X: cols, Y: rows }).is_ok() }
    }

    pub fn send_ctrl_c(&self) -> bool {
        unsafe { GenerateConsoleCtrlEvent(CTRL_C_EVENT, self.process_id).is_ok() }
    }

    pub fn close(&mut self) {
        unsafe {
            if self.input_write.0 != 0 {
                let _ = CloseHandle(self.input_write);
                self.input_write = HANDLE(0);
            }
            if self.process_handle.0 != 0 {
                let _ = CloseHandle(self.process_handle);
                self.process_handle = HANDLE(0);
            }
            if self.hpc.0 != 0 {
                ClosePseudoConsole(self.hpc);
                self.hpc = HPCON::default();
            }
        }
    }
}
