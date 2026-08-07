# Test Python del worker audiodescrizione

Questa cartella contiene 73 test derivati dalla suite di Omni Describer.
Sono stati adattati al worker headless di Sonarpad:

1. Nessuna dipendenza da wxPython o dal player di Omni.
2. Nessun avvio di `ffmpeg.exe` o `ffprobe.exe`.
3. Gemini e ONNX Runtime sono simulati; non servono rete o chiavi API.
4. La preparazione WAV e dei chunk video viene verificata come responsabilità del backend Rust di Sonarpad.
5. I test TTS, ducking, esportazione e progetto finale sono nei moduli Rust con prefisso `omni_port_`.
