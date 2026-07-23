#!/usr/bin/env python3
"""Re-apply phase-1 unsafe reductions after a clean src checkout."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    # rss_window transmute
    p = ROOT / "src/app_windows/rss_window.rs"
    t = p.read_text(encoding="utf-8")
    t = t.replace(
        "s.tree_proc = mem::transmute::<isize, WNDPROC>(old)",
        "s.tree_proc = crate::isize_to_wndproc_safe(old)",
    )
    t = t.replace(
        "s.preview_proc = mem::transmute::<isize, WNDPROC>(old)",
        "s.preview_proc = crate::isize_to_wndproc_safe(old)",
    )
    p.write_text(t, encoding="utf-8")
    print("rss_window: transmute fixed")

    # editor_manager
    p = ROOT / "src/editor_manager.rs"
    t = p.read_text(encoding="utf-8")
    old = """CallWindowProcW(
                    Some(std::mem::transmute::<
                        isize,
                        unsafe extern \"system\" fn(HWND, u32, WPARAM, LPARAM) -> LRESULT,
                    >(prev)),
                    hwnd,
                    msg,
                    wparam,
                    lparam,
                )"""
    new = """crate::call_window_proc_w_safe(
                    crate::isize_to_wndproc_safe(prev),
                    hwnd,
                    msg,
                    wparam,
                    lparam,
                )"""
    if old not in t:
        raise SystemExit("editor_manager pattern not found")
    t = t.replace(old, new)
    t = t.replace("CallWindowProcW, DefWindowProcW", "DefWindowProcW")
    p.write_text(t, encoding="utf-8")
    print("editor_manager: transmute fixed")

    # spellcheck
    p = ROOT / "src/spellcheck/windows_spellcheck.rs"
    t = p.read_text(encoding="utf-8")
    if "ComGuard" not in t:
        t = t.replace(
            "use crate::log_debug;",
            "use crate::com_guard::ComGuard;\nuse crate::log_debug;",
        )
        t = t.replace(
            "use windows::Win32::System::Com::{\n"
            "    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoInitializeEx, CoTaskMemFree, IEnumString,\n"
            "};",
            "use windows::Win32::System::Com::{CLSCTX_ALL, CoTaskMemFree, IEnumString};",
        )
        t = t.replace(
            "pub struct WindowsSpellChecker {\n"
            "    factory: Option<ISpellCheckerFactory>,\n"
            "    checker: Option<ISpellChecker>,\n"
            "    language: Option<String>,\n"
            "}",
            "pub struct WindowsSpellChecker {\n"
            "    factory: Option<ISpellCheckerFactory>,\n"
            "    checker: Option<ISpellChecker>,\n"
            "    language: Option<String>,\n"
            "    _com: Option<ComGuard>,\n"
            "}",
        )
        old = """    fn ensure_com(&self) -> bool {
        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            if hr.is_ok() || hr.0 == 0x80010106_u32 as i32 {
                true
            } else {
                log_debug(&format!("Spellcheck: CoInitializeEx failed: {hr:?}"));
                false
            }
        }
    }"""
        new = """    fn ensure_com(&mut self) -> bool {
        if self._com.is_some() {
            return true;
        }
        match ComGuard::new_sta() {
            Ok(guard) => {
                self._com = Some(guard);
                true
            }
            Err(err) => {
                log_debug(&format!("Spellcheck: CoInitializeEx failed: {err:?}"));
                false
            }
        }
    }"""
        if old not in t:
            raise SystemExit("spellcheck ensure_com not found")
        t = t.replace(old, new)
        p.write_text(t, encoding="utf-8")
        print("spellcheck: ComGuard")
    else:
        print("spellcheck already migrated")

    # accessibility JAWS -> ComGuard (simplified: only if still manual)
    p = ROOT / "src/accessibility.rs"
    t = p.read_text(encoding="utf-8")
    if "CoUninitialize" in t and "jaws_invoke_saystring" in t:
        t = t.replace(
            "use windows::Win32::System::Com::{\n"
            "    CLSCTX_ALL, CLSIDFromProgID, COINIT_APARTMENTTHREADED, CoUninitialize, DISPATCH_METHOD,\n"
            "    DISPPARAMS, EXCEPINFO, IDispatch,\n"
            "};",
            "use windows::Win32::System::Com::{\n"
            "    CLSCTX_ALL, CLSIDFromProgID, DISPATCH_METHOD, DISPPARAMS, EXCEPINFO, IDispatch,\n"
            "};",
        )
        # replace init/uninit pattern with ComGuard for both jaws functions
        import re

        def wrap_jaws(fn_src: str) -> str:
            # Replace manual init block with ComGuard at start of body
            fn_src = re.sub(
                r"let init_res = crate::com_guard::co_initialize_ex_safe\(None, COINIT_APARTMENTTHREADED\);\s*"
                r"if let Err\(e\) = init_res\.ok\(\) \{\s*"
                r"if log_failures \{\s*"
                r"crate::log_debug\(&format!\(\"JAWS CoInitializeEx failed: \{e\}\"\)\);\s*"
                r"\}\s*"
                r"return Err\(e\);\s*"
                r"\}\s*"
                r"let should_uninit = init_res\.is_ok\(\);\s*",
                "let _com = match crate::com_guard::ComGuard::new_sta() {\n"
                "        Ok(guard) => guard,\n"
                "        Err(e) => {\n"
                "            if log_failures {\n"
                '                crate::log_debug(&format!("JAWS CoInitializeEx failed: {e}"));\n'
                "            }\n"
                "            return Err(e);\n"
                "        }\n"
                "    };\n\n    ",
                fn_src,
                count=1,
            )
            fn_src = re.sub(
                r"\s*if should_uninit \{\s*unsafe \{ CoUninitialize\(\) \};\s*\}\s*",
                "\n",
                fn_src,
            )
            # unwrap (|| { ... })() if present after our earlier style - keep as is if still there
            return fn_src

        # crude: only strip CoUninitialize pairs if pattern matches
        if "COINIT_APARTMENTTHREADED" in t:
            t = wrap_jaws(t)
            # remove unused import if any remain
            t = t.replace("COINIT_APARTMENTTHREADED, ", "")
            t = t.replace(", CoUninitialize", "")
            t = t.replace("CoUninitialize, ", "")
            p.write_text(t, encoding="utf-8")
            print("accessibility: attempted ComGuard")
        else:
            print("accessibility: unexpected state")
    else:
        print("accessibility already without CoUninitialize or missing jaws")

    # audio_utils metadata
    p = ROOT / "src/audio_utils.rs"
    t = p.read_text(encoding="utf-8")
    if "set_file_metadata" in t and "CoUninitialize" in t:
        t = t.replace(
            "use windows::Win32::System::Com::{COINIT_MULTITHREADED, CoUninitialize};\n",
            "",
        )
        old = """pub fn set_file_metadata(
    path: &Path,
    title: Option<&str>,
    author: Option<&str>,
    comment: Option<&str>,
) -> Result<(), String> {
    let init_res = crate::com_guard::co_initialize_ex_safe(None, COINIT_MULTITHREADED);
    if let Err(e) = init_res.ok() {
        crate::log_debug(&format!(
            "CoInitializeEx failed while setting metadata: {e}"
        ));
    }
    let should_uninit = init_res.is_ok();

    let result = (|| {
        let path_wide = to_wide(path.to_str().ok_or("Invalid path")?);

        let store: IPropertyStore = unsafe {
            SHGetPropertyStoreFromParsingName(PCWSTR(path_wide.as_ptr()), None, GPS_READWRITE)
        }
        .map_err(|e| format!("SHGetPropertyStoreFromParsingName failed: {}", e))?;

        if let Some(t) = title {
            set_prop(&store, &PKEY_TITLE, t)?;
        }
        if let Some(a) = author {
            set_prop(&store, &PKEY_AUTHOR, a)?;
        }
        if let Some(c) = comment {
            set_prop(&store, &PKEY_COMMENT, c)?;
        }

        unsafe { store.Commit() }.map_err(|e| format!("IPropertyStore::Commit failed: {}", e))
    })();

    if should_uninit {
        unsafe { CoUninitialize() };
    }

    result
}"""
        new = """pub fn set_file_metadata(
    path: &Path,
    title: Option<&str>,
    author: Option<&str>,
    comment: Option<&str>,
) -> Result<(), String> {
    let _com = match crate::com_guard::ComGuard::new_mta() {
        Ok(guard) => Some(guard),
        Err(e) => {
            crate::log_debug(&format!(
                "CoInitializeEx failed while setting metadata: {e}"
            ));
            None
        }
    };

    let path_wide = to_wide(path.to_str().ok_or("Invalid path")?);

    let store: IPropertyStore = unsafe {
        SHGetPropertyStoreFromParsingName(PCWSTR(path_wide.as_ptr()), None, GPS_READWRITE)
    }
    .map_err(|e| format!("SHGetPropertyStoreFromParsingName failed: {}", e))?;

    if let Some(t) = title {
        set_prop(&store, &PKEY_TITLE, t)?;
    }
    if let Some(a) = author {
        set_prop(&store, &PKEY_AUTHOR, a)?;
    }
    if let Some(c) = comment {
        set_prop(&store, &PKEY_COMMENT, c)?;
    }

    unsafe { store.Commit() }.map_err(|e| format!("IPropertyStore::Commit failed: {}", e))
}"""
        if old not in t:
            raise SystemExit("audio_utils set_file_metadata not found")
        t = t.replace(old, new)
        p.write_text(t, encoding="utf-8")
        print("audio_utils: ComGuard")
    else:
        print("audio_utils already migrated or different")


if __name__ == "__main__":
    main()
