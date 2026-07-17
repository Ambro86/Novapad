#!/usr/bin/env python3
import argparse
import json
import os
import sys
import wave


BRIDGE_PROTOCOL_VERSION = 2
LONG_AUDIO_THRESHOLD_SECONDS = 30 * 60
LONG_AUDIO_CHUNK_SECONDS = 15 * 60
WHISPER_SAMPLE_RATE = 16000
QUIET_SEARCH_RADIUS_SECONDS = 10
QUIET_WINDOW_SECONDS = 0.2
QUIET_RMS_THRESHOLD = 0.02
FALLBACK_OVERLAP_SECONDS = 1


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


def choose_quiet_split(samples, target_samples):
    """Return a nearby quiet split and whether it is a real pause."""
    import numpy as np  # type: ignore

    radius = min(
        int(QUIET_SEARCH_RADIUS_SECONDS * WHISPER_SAMPLE_RATE),
        max(1, target_samples // 4),
    )
    window = max(1, int(QUIET_WINDOW_SECONDS * WHISPER_SAMPLE_RATE))
    search_start = max(0, target_samples - radius - window // 2)
    search_end = min(samples.size, target_samples + radius + window // 2)
    region = samples[search_start:search_end]
    if region.size < window:
        return min(target_samples, samples.size), False

    squared = region.astype(np.float64) ** 2
    cumulative = np.concatenate(([0.0], np.cumsum(squared)))
    step = max(1, window // 4)
    best_start = 0
    best_mean_square = float("inf")
    for start in range(0, region.size - window + 1, step):
        mean_square = (cumulative[start + window] - cumulative[start]) / window
        if mean_square < best_mean_square:
            best_mean_square = mean_square
            best_start = start

    split = search_start + best_start + window // 2
    rms_normalized = (best_mean_square**0.5) / 32768.0
    return max(1, min(split, samples.size)), rms_normalized <= QUIET_RMS_THRESHOLD


def iter_resampled_audio_chunks(path, chunk_seconds=LONG_AUDIO_CHUNK_SECONDS):
    """Decode bounded mono/16 kHz blocks, preferring pauses as boundaries."""
    import av  # type: ignore
    import numpy as np  # type: ignore

    chunk_samples = max(1, int(chunk_seconds * WHISPER_SAMPLE_RATE))
    search_samples = min(
        int(QUIET_SEARCH_RADIUS_SECONDS * WHISPER_SAMPLE_RATE),
        max(1, chunk_samples // 4),
    )
    overlap_samples = int(FALLBACK_OVERLAP_SECONDS * WHISPER_SAMPLE_RATE)
    ready_samples = chunk_samples + search_samples
    pending = []
    pending_samples = 0
    buffer_start_samples = 0

    with av.open(path, metadata_errors="ignore") as container:
        if not container.streams.audio:
            raise RuntimeError("input file has no audio stream")
        resampler = av.AudioResampler(
            format="s16",
            layout="mono",
            rate=WHISPER_SAMPLE_RATE,
        )

        def append_frame(frame):
            nonlocal pending_samples
            samples = frame.to_ndarray().reshape(-1)
            if samples.size:
                samples = np.ascontiguousarray(samples, dtype=np.int16)
                pending.append(samples)
                pending_samples += int(samples.size)

        for frame in container.decode(audio=0):
            for resampled in resampler.resample(frame):
                append_frame(resampled)
            if pending_samples >= ready_samples:
                combined = np.concatenate(pending)
                while combined.size >= ready_samples:
                    split, found_quiet = choose_quiet_split(combined, chunk_samples)
                    chunk = combined[:split]
                    yield (
                        chunk.astype(np.float32) / 32768.0,
                        float(buffer_start_samples) / WHISPER_SAMPLE_RATE,
                    )
                    consumed = split
                    if not found_quiet:
                        consumed = max(1, split - min(overlap_samples, split // 4))
                    combined = combined[consumed:]
                    buffer_start_samples += consumed
                pending = [combined.copy()] if combined.size else []
                pending_samples = int(combined.size)

        for resampled in resampler.resample(None):
            append_frame(resampled)

    if pending_samples:
        combined = np.concatenate(pending)
        while combined.size > chunk_samples:
            split, found_quiet = choose_quiet_split(combined, chunk_samples)
            chunk = combined[:split]
            yield (
                chunk.astype(np.float32) / 32768.0,
                float(buffer_start_samples) / WHISPER_SAMPLE_RATE,
            )
            consumed = split
            if not found_quiet:
                consumed = max(1, split - min(overlap_samples, split // 4))
            combined = combined[consumed:]
            buffer_start_samples += consumed
        if combined.size:
            yield (
                combined.astype(np.float32) / 32768.0,
                float(buffer_start_samples) / WHISPER_SAMPLE_RATE,
            )


def append_transcribed_segments(
    segments,
    parts,
    timestamps,
    offset_seconds,
    total_duration,
    last_progress,
):
    for segment in segments:
        text = (segment.text or "").strip()
        if text:
            if timestamps:
                start = offset_seconds + (getattr(segment, "start", 0.0) or 0.0)
                parts.append(f"[{format_timestamp(start)}] {text}")
            else:
                parts.append(text)
        if total_duration > 0 and segment.end is not None:
            end = offset_seconds + float(segment.end)
            pct = int((end / total_duration) * 100.0)
            next_progress = max(0, min(99, pct))
            if next_progress > last_progress:
                last_progress = next_progress
                print(f"PROGRESS:{last_progress}", flush=True)
    return last_progress


def transcribe_long_input(model, input_path, language, timestamps, total_duration):
    parts = []
    last_progress = 0
    selected_language = language or None
    detected_language = ""

    for audio_chunk, offset_seconds in iter_resampled_audio_chunks(input_path):
        chunk_duration = float(audio_chunk.shape[0]) / WHISPER_SAMPLE_RATE
        segments, info = model.transcribe(
            audio_chunk,
            language=selected_language,
            vad_filter=False,
            beam_size=5,
            condition_on_previous_text=False,
        )
        if not detected_language:
            detected_language = getattr(info, "language", "") or ""
        if selected_language is None and detected_language:
            selected_language = detected_language
        last_progress = append_transcribed_segments(
            segments,
            parts,
            timestamps,
            offset_seconds,
            total_duration,
            last_progress,
        )
        if total_duration > 0:
            chunk_end = offset_seconds + chunk_duration
            pct = int((chunk_end / total_duration) * 100.0)
            next_progress = max(0, min(99, pct))
            if next_progress > last_progress:
                last_progress = next_progress
                print(f"PROGRESS:{last_progress}", flush=True)

    transcript = ("\n".join(parts) if timestamps else " ".join(parts)).strip()
    return {"ok": True, "text": transcript, "language": detected_language}


def transcribe_input(model, input_path, language, timestamps):
    if not os.path.isfile(input_path):
        return {"ok": False, "error": f"input file not found: {input_path}"}

    total_duration = audio_duration_seconds(input_path)
    if total_duration >= LONG_AUDIO_THRESHOLD_SECONDS:
        return transcribe_long_input(
            model,
            input_path,
            language,
            timestamps,
            total_duration,
        )

    last_progress = 0
    segments, info = model.transcribe(
        input_path,
        language=language or None,
        vad_filter=False,
        beam_size=5,
        condition_on_previous_text=False,
    )

    parts = []
    append_transcribed_segments(
        segments,
        parts,
        timestamps,
        0.0,
        total_duration,
        last_progress,
    )

    transcript = ("\n".join(parts) if timestamps else " ".join(parts)).strip()
    detected_language = ""
    try:
        detected_language = getattr(info, "language", "") or ""
    except Exception:
        detected_language = ""
    return {"ok": True, "text": transcript, "language": detected_language}


def worker_loop(model, selected_device, selected_compute_type):
    print_json(
        {
            "ready": True,
            "bridge_version": BRIDGE_PROTOCOL_VERSION,
            "backend": selected_device,
            "compute_type": selected_compute_type,
        }
    )

    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue
        try:
            request = json.loads(line)
        except Exception as exc:
            print_json({"ok": False, "error": f"invalid worker request: {exc}"})
            continue

        command = str(request.get("command", "") or "").strip().lower()
        if command == "shutdown":
            return 0
        if command != "transcribe":
            print_json({"ok": False, "error": f"unsupported worker command: {command}"})
            continue

        try:
            result = transcribe_input(
                model,
                str(request.get("input", "") or ""),
                str(request.get("language", "") or ""),
                bool(request.get("timestamps", False)),
            )
            result["backend"] = selected_device
            result["compute_type"] = selected_compute_type
            print_json(result)
        except Exception as exc:
            print_json({"ok": False, "error": str(exc)})
    return 0


def main():
    configure_stdio_utf8()
    parser = argparse.ArgumentParser()
    parser.add_argument("--bridge-version", action="store_true")
    parser.add_argument("--input", default="", help="Path to audio file")
    parser.add_argument("--model", default="", help="small | medium | large-v3")
    parser.add_argument("--language", default="", help="Language code, e.g. it")
    parser.add_argument("--download-root", default="", help="Model cache directory")
    parser.add_argument("--timestamps", action="store_true", help="Include segment timestamps")
    parser.add_argument("--worker", action="store_true", help="Keep model loaded for multiple requests")
    args = parser.parse_args()

    if args.bridge_version:
        print_json({"bridge_version": BRIDGE_PROTOCOL_VERSION})
        return 0
    if not args.model:
        print_json({"ok": False, "error": "model is required"})
        return 1

    if not args.worker and not os.path.isfile(args.input):
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
        if args.worker:
            return worker_loop(model, selected_device, selected_compute_type)

        result = transcribe_input(model, args.input, args.language, args.timestamps)
        result["backend"] = selected_device
        result["compute_type"] = selected_compute_type
        print_json(result)
        return 0
    except Exception as exc:
        print_json({"ok": False, "error": str(exc)})
        return 1


if __name__ == "__main__":
    sys.exit(main())
