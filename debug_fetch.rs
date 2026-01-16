mod curl_client;
fn main() {
    let url = "https://www.adnkronos.com/economia/vicenzaoro-2026-inaugurato-oggi-il-salone-internazionale-del-gioiello-di-ieg_7HBF2sDgfGE1ZAQSK3FjxG";
    match curl_client::fetch_url_impersonated(url) {
        Ok(bytes) => {
            let html = String::from_utf8_lossy(&bytes);
            std::fs::write("debug_adnkronos.html", html.as_ref()).unwrap();
            println!("HTML salvato in debug_adnkronos.html ({} bytes)", bytes.len());
        }
        Err(e) => println!("Errore: {}", e),
    }
}
