#!/usr/bin/env python3
import argparse
import json
import os
import sys
import wave


def print_json(payload):
    print(json.dumps(payload, ensure_ascii=False), flush=True)


def configure_stdio_utf8():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def audio_duration_seconds(path):
    try:
        import av  # type: ignore

        with av.open(path) as container:
            if container.duration is not None and container.duration > 0:
                return float(container.duration) / float(av.time_base)
            for stream in container.streams.audio:
                if (
                    stream.duration is not None
                    and stream.duration > 0
                    and stream.time_base is not None
                ):
                    return float(stream.duration * stream.time_base)
                if (
                    stream.frames is not None
                    and stream.frames > 0
                    and stream.rate is not None
                    and float(stream.rate) > 0.0
                ):
                    return float(stream.frames) / float(stream.rate)
    except Exception:
        pass

    try:
        with wave.open(path, "rb") as wav_file:
            rate = wav_file.getframerate()
            frames = wav_file.getnframes()
            if rate <= 0:
                return 0.0
            return float(frames) / float(rate)
    except Exception:
        return 0.0


def candidate_backends():
    force_cuda = os.environ.get("SONARPAD_FORCE_CUDA", "").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )

    if force_cuda:
        return [("cuda", "int8_float16"), ("cpu", "int8")]

    try:
        import torch  # type: ignore

        if torch.cuda.is_available():
            return [("cuda", "int8_float16"), ("cpu", "int8")]
    except Exception:
        pass

    return [("cpu", "int8")]


def format_timestamp(seconds):
    total = max(0, int(seconds))
    hours = total // 3600
    minutes = (total % 3600) // 60
    secs = total % 60
    if hours > 0:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"


def main():
    configure_stdio_utf8()
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Path to audio file")
    parser.add_argument("--model", required=True, help="small | medium | large-v3")
    parser.add_argument("--language", default="", help="Language code, e.g. it")
    parser.add_argument("--download-root", default="", help="Model cache directory")
    parser.add_argument("--timestamps", action="store_true", help="Include segment timestamps")
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
        model = None
        selected_device = "cpu"
        selected_compute_type = "int8"
        last_init_error = ""
        for device, compute_type in candidate_backends():
            model_kwargs = {
                "device": device,
                "compute_type": compute_type,
            }
            if args.download_root:
                model_kwargs["download_root"] = args.download_root
            try:
                model = WhisperModel(args.model, **model_kwargs)
                selected_device = device
                selected_compute_type = compute_type
                break
            except Exception as exc:
                last_init_error = f"{device}/{compute_type}: {exc}"

        if model is None:
            print_json({"ok": False, "error": f"model init failed: {last_init_error}"})
            return 1
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
                if args.timestamps:
                    start_ts = format_timestamp(getattr(segment, "start", 0.0) or 0.0)
                    parts.append(f"[{start_ts}] {text}")
                else:
                    parts.append(text)
            if total_duration > 0 and segment.end is not None:
                pct = int((float(segment.end) / total_duration) * 100.0)
                if pct > last_progress:
                    last_progress = max(0, min(99, pct))
                    print(f"PROGRESS:{last_progress}", flush=True)

        transcript = ("\n".join(parts) if args.timestamps else " ".join(parts)).strip()
        language = ""
        try:
            language = getattr(info, "language", "") or ""
        except Exception:
            language = ""

        print_json(
            {
                "ok": True,
                "text": transcript,
                "language": language,
                "backend": selected_device,
                "compute_type": selected_compute_type,
            }
        )
        return 0
    except Exception as exc:
        print_json({"ok": False, "error": str(exc)})
        return 1


if __name__ == "__main__":
    sys.exit(main())
