use std::cell::RefCell;
use std::fs;
use std::io::{Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};
use std::rc::Rc;
use std::sync::mpsc;
use std::sync::{
    Arc,
    atomic::{AtomicBool, Ordering},
};
use std::time::{Duration, Instant};

use windows::Win32::Foundation::{HINSTANCE, HWND};
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance};
use windows::Win32::UI::Shell::Common::COMDLG_FILTERSPEC;
use windows::Win32::UI::Shell::{
    FileSaveDialog, IFileSaveDialog, IShellItem, SHCreateItemFromParsingName, SIGDN_FILESYSPATH,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, DispatchMessageW, MSG, PM_REMOVE, PeekMessageW, SW_HIDE,
    ShowWindow, TranslateMessage, WINDOW_EX_STYLE, WS_OVERLAPPED,
};
use windows::core::PCWSTR;

use crate::accessibility::to_wide;
use crate::{editor_manager, i18n, show_error, with_state};

const ROUTE_MAP_WIDTH: i32 = 1200;
const ROUTE_MAP_HEIGHT: i32 = 800;
const ROUTE_MAP_TIMEOUT: Duration = Duration::from_secs(25);

enum RenderEvent {
    Captured,
    Failed(String),
}

pub fn export_current_route_map_image(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let Some(route_map) = editor_manager::current_route_map(parent) else {
        show_error(parent, language, &i18n::tr(language, "route.map.no_map"));
        return;
    };

    let Some(output_path) = save_route_map_dialog(parent, language, &route_map.suggested_filename)
    else {
        return;
    };

    match build_leaflet_html(&route_map, language)
        .and_then(|html| capture_route_map_with_webview2(parent, &html, &output_path))
    {
        Ok(()) => {
            crate::screen_reader_speak(&i18n::tr(language, "route.map.saved"));
        }
        Err(error) => show_error(parent, language, &error),
    }
}

fn save_route_map_dialog(
    parent: HWND,
    language: crate::settings::Language,
    suggested_filename: &str,
) -> Option<PathBuf> {
    let _com = match crate::com_guard::ComGuard::new_sta() {
        Ok(com) => com,
        Err(error) => {
            crate::log_debug(&format!("Route map save dialog COM init failed: {error}"));
            return None;
        }
    };

    unsafe {
        // SAFETY: The current thread has COM initialized above; FileSaveDialog is a COM coclass
        // designed to be created on an STA UI thread.
        let dialog: IFileSaveDialog = CoCreateInstance(&FileSaveDialog, None, CLSCTX_ALL).ok()?;

        let filter_name = to_wide("PNG (*.png)");
        let filter_spec = to_wide("*.png");
        let filters = [COMDLG_FILTERSPEC {
            pszName: PCWSTR(filter_name.as_ptr()),
            pszSpec: PCWSTR(filter_spec.as_ptr()),
        }];
        dialog.SetFileTypes(&filters).ok()?;
        dialog.SetFileTypeIndex(1).ok()?;

        let default_ext = to_wide("png");
        dialog
            .SetDefaultExtension(PCWSTR(default_ext.as_ptr()))
            .ok()?;

        let default_filename = if suggested_filename.trim().is_empty() {
            i18n::tr(language, "route.map.default_filename")
        } else {
            format!("{suggested_filename}.png")
        };
        let default_name = to_wide(&default_filename);
        dialog.SetFileName(PCWSTR(default_name.as_ptr())).ok()?;

        let initial_dir = crate::settings::default_images_save_folder();
        crate::log_if_err!(std::fs::create_dir_all(&initial_dir));
        let initial_dir_wide = to_wide(&initial_dir);
        if let Ok(shell_folder) =
            SHCreateItemFromParsingName::<_, _, IShellItem>(PCWSTR(initial_dir_wide.as_ptr()), None)
        {
            if let Err(error) = dialog.SetDefaultFolder(&shell_folder) {
                crate::log_debug(&format!("Route map SetDefaultFolder failed: {error}"));
            }
            if let Err(error) = dialog.SetFolder(&shell_folder) {
                crate::log_debug(&format!("Route map SetFolder failed: {error}"));
            }
        }

        dialog.Show(parent).ok()?;
        let result = dialog.GetResult().ok()?;
        let path = result.GetDisplayName(SIGDN_FILESYSPATH).ok()?;
        Some(PathBuf::from(path.to_string().ok()?))
    }
}

fn build_leaflet_html(
    route_map: &crate::app_windows::route_service::RouteMapData,
    language: crate::settings::Language,
) -> Result<String, String> {
    let geometry = serde_json::to_string(&route_map.geometry)
        .map_err(|error| format!("{} {error}", i18n::tr(language, "route.map.geometry_error")))?;
    let from_label = serde_json::to_string(&route_map.from_label)
        .map_err(|error| format!("{} {error}", i18n::tr(language, "route.map.from_error")))?;
    let to_label = serde_json::to_string(&route_map.to_label)
        .map_err(|error| format!("{} {error}", i18n::tr(language, "route.map.to_error")))?;

    Ok(format!(
        r#"<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width={ROUTE_MAP_WIDTH}, initial-scale=1">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<style>
html, body, #map {{
    width: {ROUTE_MAP_WIDTH}px;
    height: {ROUTE_MAP_HEIGHT}px;
    margin: 0;
    padding: 0;
    overflow: hidden;
}}
</style>
</head>
<body>
<div id="map"></div>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
const geometry = {geometry};
const fromLabel = {from_label};
const toLabel = {to_label};
const latLngs = geometry.map(([lon, lat]) => [lat, lon]);
let postedReady = false;

function postReady() {{
    if (postedReady) {{
        return;
    }}
    postedReady = true;
    window.chrome.webview.postMessage('map_ready');
}}

const map = L.map('map', {{ zoomControl: false, attributionControl: true, fadeAnimation: false }});
const tiles = L.tileLayer('https://tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
    maxZoom: 19,
    crossOrigin: true,
    attribution: '&copy; OpenStreetMap contributors'
}}).addTo(map);
const routeShadow = L.polyline(latLngs, {{
    color: '#ffffff',
    weight: 13,
    opacity: 1,
    lineCap: 'round',
    lineJoin: 'round'
}}).addTo(map);
const line = L.polyline(latLngs, {{
    color: '#0b57d0',
    weight: 7,
    opacity: 0.98,
    lineCap: 'round',
    lineJoin: 'round'
}}).addTo(map);
const startMarker = L.marker(latLngs[0]).addTo(map).bindPopup(fromLabel);
const endMarker = L.marker(latLngs[latLngs.length - 1]).addTo(map).bindPopup(toLabel);
const bounds = line.getBounds()
    .extend(startMarker.getLatLng())
    .extend(endMarker.getLatLng());
map.fitBounds(bounds.pad(0.25), {{ padding: [120, 120], maxZoom: 16, animate: false }});

map.whenReady(() => {{
    tiles.once('load', () => window.setTimeout(postReady, 500));
    window.setTimeout(postReady, 4500);
}});
</script>
</body>
</html>"#
    ))
}

fn capture_route_map_with_webview2(
    parent: HWND,
    html: &str,
    output_path: &Path,
) -> Result<(), String> {
    let _com = crate::com_guard::ComGuard::new_sta()
        .map_err(|error| format!("Impossibile inizializzare COM per WebView2: {error}"))?;
    let host = create_capture_host_window(parent)?;
    let result = capture_route_map_with_webview2_inner(host, html, output_path);
    unsafe {
        // SAFETY: host was created by create_capture_host_window in this function and is no
        // longer used after the WebView2 capture loop returns.
        if let Err(error) = DestroyWindow(host) {
            crate::log_debug(&format!("Route map DestroyWindow failed: {error}"));
        }
    }
    result
}

fn capture_route_map_with_webview2_inner(
    host: HWND,
    html: &str,
    output_path: &Path,
) -> Result<(), String> {
    if output_path.exists() {
        fs::remove_file(output_path).map_err(|error| {
            format!(
                "Impossibile sostituire il file immagine esistente {}: {error}",
                output_path.display()
            )
        })?;
    }

    let user_data_dir = std::env::temp_dir()
        .join("sonarpad_route_map")
        .join(format!("webview2_profile_{}", std::process::id()));
    fs::create_dir_all(&user_data_dir)
        .map_err(|error| format!("Impossibile creare il profilo temporaneo WebView2: {error}"))?;

    let (tx, rx) = mpsc::channel::<RenderEvent>();
    let controller_holder: Rc<RefCell<Option<webview2::Controller>>> = Rc::new(RefCell::new(None));
    let controller_holder_for_create = controller_holder.clone();
    let html = html.to_string();
    let output_path = output_path.to_path_buf();
    let nav_completed = Arc::new(AtomicBool::new(false));
    let capture_started = Arc::new(AtomicBool::new(false));
    let hwnd_webview2 = host.0 as winapi::shared::windef::HWND;

    webview2::Environment::builder()
        .with_user_data_folder(&user_data_dir)
        .with_additional_browser_arguments("--no-first-run --no-default-browser-check")
        .build(move |environment| {
            let environment = environment.map_err(|error| {
                crate::log_debug(&format!("Route map WebView2 environment failed: {error}"));
                error
            })?;
            let tx_create = tx.clone();
            environment.create_controller(hwnd_webview2, move |controller| {
                let controller = controller.map_err(|error| {
                    crate::log_debug(&format!("Route map WebView2 controller failed: {error}"));
                    error
                })?;
                controller.put_bounds(winapi::shared::windef::RECT {
                    left: 0,
                    top: 0,
                    right: ROUTE_MAP_WIDTH,
                    bottom: ROUTE_MAP_HEIGHT,
                })?;
                controller.put_is_visible(true)?;

                let webview = controller.get_webview()?;
                let tx_nav = tx_create.clone();
                let nav_completed_for_nav = nav_completed.clone();
                webview.add_navigation_completed(move |_, args| {
                    if args.get_is_success()? {
                        nav_completed_for_nav.store(true, Ordering::SeqCst);
                    } else if tx_nav
                        .send(RenderEvent::Failed(
                            "La pagina mappa WebView2 non e' stata caricata.".to_string(),
                        ))
                        .is_err()
                    {
                        crate::log_debug("Route map navigation failure receiver dropped");
                    }
                    Ok(())
                })?;

                let tx_msg = tx_create.clone();
                let output_for_capture = output_path.clone();
                let nav_completed_for_msg = nav_completed.clone();
                let capture_started_for_msg = capture_started.clone();
                webview.add_web_message_received(move |webview, message| {
                    let message = message.try_get_web_message_as_string()?;
                    if message != "map_ready" {
                        return Ok(());
                    }
                    if !nav_completed_for_msg.load(Ordering::SeqCst) {
                        if tx_msg
                            .send(RenderEvent::Failed(
                                "La mappa ha inviato map_ready prima del completamento navigazione."
                                    .to_string(),
                            ))
                            .is_err()
                        {
                            crate::log_debug("Route map premature ready receiver dropped");
                        }
                        return Ok(());
                    }
                    if capture_started_for_msg.swap(true, Ordering::SeqCst) {
                        return Ok(());
                    }

                    let stream = webview2::Stream::from_bytes(&[]);
                    let mut stream_for_callback = stream.clone();
                    let tx_capture = tx_msg.clone();
                    let output_for_callback = output_for_capture.clone();
                    webview.capture_preview(
                        webview2::CapturePreviewImageFormat::PNG,
                        stream,
                        move |result| {
                            if let Err(error) = result {
                                if tx_capture
                                    .send(RenderEvent::Failed(format!(
                                        "Cattura PNG WebView2 fallita: {error}"
                                    )))
                                    .is_err()
                                {
                                    crate::log_debug("Route map capture failure receiver dropped");
                                }
                                return Ok(());
                            }
                            stream_for_callback.seek(SeekFrom::Start(0))?;
                            let mut png = Vec::new();
                            stream_for_callback.read_to_end(&mut png)?;
                            fs::write(&output_for_callback, png)?;
                            if tx_capture.send(RenderEvent::Captured).is_err() {
                                crate::log_debug("Route map capture receiver dropped");
                            }
                            Ok(())
                        },
                    )?;
                    Ok(())
                })?;

                webview.navigate_to_string(&html)?;
                controller_holder_for_create.replace(Some(controller));
                Ok(())
            })
        })
        .map_err(|error| format!("Impossibile avviare WebView2 per la mappa: {error}"))?;

    let result = wait_for_render_event(rx);
    controller_holder.borrow_mut().take();
    crate::log_if_err!(fs::remove_dir_all(&user_data_dir));
    result
}

fn wait_for_render_event(rx: mpsc::Receiver<RenderEvent>) -> Result<(), String> {
    let deadline = Instant::now() + ROUTE_MAP_TIMEOUT;
    loop {
        match rx.try_recv() {
            Ok(RenderEvent::Captured) => return Ok(()),
            Ok(RenderEvent::Failed(error)) => return Err(error),
            Err(mpsc::TryRecvError::Disconnected) => {
                return Err("Renderer WebView2 mappa chiuso prima della cattura.".to_string());
            }
            Err(mpsc::TryRecvError::Empty) => {}
        }

        if Instant::now() >= deadline {
            return Err("La mappa non e' diventata pronta entro il tempo massimo.".to_string());
        }

        pump_one_message();
    }
}

fn pump_one_message() {
    unsafe {
        // SAFETY: Standard UI-thread message pumping while waiting for WebView2 async callbacks;
        // MSG is initialized by PeekMessageW before being dispatched.
        let mut message = MSG::default();
        if PeekMessageW(&mut message, HWND(0), 0, 0, PM_REMOVE).as_bool() {
            let _translated = TranslateMessage(&message);
            DispatchMessageW(&message);
        } else {
            std::thread::sleep(Duration::from_millis(10));
        }
    }
}

fn create_capture_host_window(parent: HWND) -> Result<HWND, String> {
    let class_name = to_wide("STATIC");
    let title = to_wide("Sonarpad route map capture");
    let hwnd = unsafe {
        // SAFETY: Uses a built-in STATIC window class to host WebView2 offscreen; all pointers are
        // valid NUL-terminated UTF-16 buffers for the duration of the call.
        CreateWindowExW(
            WINDOW_EX_STYLE(0),
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_OVERLAPPED,
            -32000,
            -32000,
            ROUTE_MAP_WIDTH,
            ROUTE_MAP_HEIGHT,
            parent,
            None,
            HINSTANCE(0),
            None,
        )
    };
    if hwnd.0 == 0 {
        return Err("Impossibile creare la finestra temporanea per WebView2.".to_string());
    }

    unsafe {
        // SAFETY: hwnd is a valid window handle created above. It is kept hidden offscreen; WebView2
        // still receives a real parent HWND and bounds for rendering.
        let _shown = ShowWindow(hwnd, SW_HIDE);
    }
    Ok(hwnd)
}
