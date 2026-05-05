import asyncio
import websockets
import uuid
import json
import ssl

async def test_edge():
    ws_url = "wss://speech.platform.bing.com/consumer/speech/synthesize/readaloud/edge/v1?TrustedClientToken=6A5AA1D4EAFF4E9FB37E23D68491D6F4"
    ssl_context = ssl.create_default_context()
    
    headers = {
        "Origin": "chrome-extension://jdiccldimpdaibmpdkjnbnkndfdndkgc",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36 Edg/113.0.1774.35"
    }

    async with websockets.connect(ws_url, ssl=ssl_context, additional_headers=headers) as ws:
        request_id = uuid.uuid4().hex
        
        # 1. Send speech.config
        config = {"context": {"synthesis": {"audio": {"metadataoptions": {"sentenceBoundaryEnabled": "false", "wordBoundaryEnabled": "true"}, "outputFormat": "audio-24khz-48kbitrate-mono-mp3"}}}}
        config_msg = f"X-Timestamp: 1680000000000\r\nContent-Type: application/json; charset=utf-8\r\nPath: speech.config\r\n\r\n{json.dumps(config)}"
        await ws.send(config_msg)
        
        # 2. Send ssml
        # Test 1: inside prosody
        ssml = "<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='en-US'><voice name='en-US-AriaNeural'><prosody pitch='+0Hz' rate='+0%' volume='+0%'>Hello <mstts:silence type=\"Sentenceboundary\" value=\"500ms\"/> world</prosody></voice></speak>"
        ssml_msg = f"X-RequestId: {request_id}\r\nContent-Type: application/ssml+xml\r\nX-Timestamp: 1680000000000Z\r\nPath: ssml\r\n\r\n{ssml}"
        
        print("Sending SSML:", ssml)
        await ws.send(ssml_msg)
        
        # 3. Receive
        while True:
            try:
                response = await ws.recv()
                if isinstance(response, str):
                    print("Received TEXT:", response.split('\r\n\r\n')[0])
                    if "Path: turn.end" in response:
                        break
                else:
                    print("Received BINARY data of length:", len(response))
            except Exception as e:
                print("Error:", e)
                break

asyncio.run(test_edge())
