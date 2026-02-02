import asyncio
import logging
import os
import time
import edge_tts

INPUT_PATH = r"C:\Users\ambro\iCloudDrive\Lorna Byrne\2 Lorna Byrne-Una scala per il cielo.txt"
VOICE = "it-IT-ElsaNeural"
OUTPUT_AUDIO = os.path.join(os.path.dirname(__file__), "edge_tts_output.mp3")
LOG_PATH = os.path.join(os.path.dirname(__file__), "edge_tts_read_file.log")
MAX_LINES = 10

logging.basicConfig(
    filename=LOG_PATH,
    filemode="a",
    format="%(asctime)s %(levelname)s %(message)s",
    level=logging.INFO,
)

async def main() -> None:
    if not os.path.exists(INPUT_PATH):
        raise FileNotFoundError(INPUT_PATH)

    with open(INPUT_PATH, "r", encoding="utf-8") as f:
        lines = []
        for _ in range(MAX_LINES):
            line = f.readline()
            if line == "":
                break
            lines.append(line)
        text = "".join(lines)

    logging.info("Edge TTS run start: chars=%d voice=%s", len(text), VOICE)
    t_start = time.perf_counter()
    audio_bytes = 0
    first_audio_at = None

    try:
        communicate = edge_tts.Communicate(text, VOICE)
        with open(OUTPUT_AUDIO, "wb") as out:
            async for chunk in communicate.stream():
                if chunk["type"] == "audio":
                    if first_audio_at is None:
                        first_audio_at = time.perf_counter()
                        logging.info("First audio after %.2fs", first_audio_at - t_start)
                    data = chunk["data"]
                    out.write(data)
                    audio_bytes += len(data)
    except Exception:
        logging.exception("Edge TTS run error")
        raise
    finally:
        t_end = time.perf_counter()
        logging.info(
            "Edge TTS run end: elapsed=%.2fs audio_bytes=%d output=%s",
            t_end - t_start,
            audio_bytes,
            OUTPUT_AUDIO,
        )

if __name__ == "__main__":
    asyncio.run(main())
