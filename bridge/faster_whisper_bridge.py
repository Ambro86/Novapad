#!/usr/bin/env python3
import argparse
import json
import os
import sys
import wave


def print_json(payload):
    print(json.dumps(payload, ensure_ascii=False), flush=True)


def audio_duration_seconds(path):
    try:
        with wave.open(path, "rb") as wav_file:
            rate = wav_file.getframerate()
            frames = wav_file.getnframes()
            if rate <= 0:
                return 0.0
            return float(frames) / float(rate)
    except Exception:
        return 0.0


def choose_device():
    try:
        import torch  # type: ignore

        if torch.cuda.is_available():
            return "cuda", "int8_float16"
    except Exception:
        pass
    return "cpu", "int8"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Path to WAV file")
    parser.add_argument("--model", required=True, help="small | medium | large-v3")
    parser.add_argument("--language", default="it", help="Language code, e.g. it")
    parser.add_argument("--download-root", default="", help="Model cache directory")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print_json({"ok": False, "error": f"input file not found: {args.input}"})
        return 1

    try:
        from faster_whisper import WhisperModel  # type: ignore
    except Exception as exc:
        print_json({"ok": False, "error": f"faster-whisper import failed: {exc}"})
        return 1

    try:
        device, compute_type = choose_device()
        model_kwargs = {
            "device": device,
            "compute_type": compute_type,
        }
        if args.download_root:
            model_kwargs["download_root"] = args.download_root

        model = WhisperModel(args.model, **model_kwargs)
        total_duration = audio_duration_seconds(args.input)
        last_progress = 0

        segments, info = model.transcribe(
            args.input,
            language=args.language or None,
            vad_filter=False,
            beam_size=5,
            condition_on_previous_text=False,
        )

        parts = []
        for segment in segments:
            text = (segment.text or "").strip()
            if text:
                parts.append(text)
            if total_duration > 0 and segment.end is not None:
                pct = int((float(segment.end) / total_duration) * 100.0)
                if pct > last_progress:
                    last_progress = max(0, min(99, pct))
                    print(f"PROGRESS:{last_progress}", flush=True)

        transcript = " ".join(parts).strip()
        language = ""
        try:
            language = getattr(info, "language", "") or ""
        except Exception:
            language = ""

        print_json({"ok": True, "text": transcript, "language": language})
        return 0
    except Exception as exc:
        print_json({"ok": False, "error": str(exc)})
        return 1


if __name__ == "__main__":
    sys.exit(main())
