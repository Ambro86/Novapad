use sonarpad::tts_engine::*;

fn main() {
    let rt = tokio::runtime::Runtime::new().unwrap();
    rt.block_on(async {
        println!("Testing Edge TTS...");
        // Test normal text
        let res = download_edge_chunk_ws_with_retry("Hello world", "en-US-AriaNeural", 0, 0, 0).await;
        match res {
            Ok(bytes) => println!("Normal Text - Success! Bytes: {}", bytes.len()),
            Err(e) => println!("Normal Text - Error: {}", e),
        }
        
        // Test mstts:silence inside prosody
        let res = download_edge_chunk_ws_with_retry("Hello <mstts:silence type=\"Sentenceboundary\" value=\"500ms\"/> world", "en-US-AriaNeural", 0, 0, 0).await;
        match res {
            Ok(bytes) => println!("mstts:silence inside prosody - Success! Bytes: {}", bytes.len()),
            Err(e) => println!("mstts:silence inside prosody - Error: {}", e),
        }
        
        // Test break inside prosody
        let res = download_edge_chunk_ws_with_retry("Hello <break time=\"500ms\"/> world", "en-US-AriaNeural", 0, 0, 0).await;
        match res {
            Ok(bytes) => println!("break inside prosody - Success! Bytes: {}", bytes.len()),
            Err(e) => println!("break inside prosody - Error: {}", e),
        }
        
        // Test break OUTSIDE prosody
        let res = download_edge_chunk_ws_with_retry("Hello </prosody><break time=\"500ms\"/><prosody pitch='+0Hz' rate='+0%' volume='+0%'> world", "en-US-AriaNeural", 0, 0, 0).await;
        match res {
            Ok(bytes) => println!("break outside prosody - Success! Bytes: {}", bytes.len()),
            Err(e) => println!("break outside prosody - Error: {}", e),
        }
    });
}
