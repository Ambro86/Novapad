"""Isolated native-pitch probe for copied Google Natural voice packages."""

from __future__ import annotations

import dataclasses
import hashlib
import importlib
import json
import os
from pathlib import Path
import sys
import time
import types
import wave


def main() -> None:
    if len(sys.argv) not in (3, 4):
        raise SystemExit(
            "usage: google_natural_pitch_probe.py ADDON_DIR VARIANT_DIR [ENGINE_DIR]"
        )
    addon_dir = Path(sys.argv[1]).resolve()
    variant_dir = Path(sys.argv[2]).resolve()

    package = types.ModuleType("googleTtsForNvda")
    package.__path__ = [str(addon_dir)]
    sys.modules["googleTtsForNvda"] = package
    catalog_module = importlib.import_module("googleTtsForNvda.catalog")
    voice_store = importlib.import_module("googleTtsForNvda.voice_store")
    voice_store.data_root = lambda: variant_dir
    voice_store.voice_dir = lambda: variant_dir / "voices"
    bridge_module = importlib.import_module("googleTtsForNvda.bridge")
    if len(sys.argv) == 4:
        bridge_module.ENGINE_DIR = Path(sys.argv[3]).resolve()

    packages = []
    for candidate in catalog_module.VoiceCatalog.load().packages:
        archive = variant_dir / "voices" / candidate.fileName
        if archive.is_file():
            packages.append(
                dataclasses.replace(
                    candidate,
                    sha256Checksum=hashlib.sha256(archive.read_bytes()).hexdigest(),
                    compressedSize=archive.stat().st_size,
                )
            )
    catalog = catalog_module.VoiceCatalog(packages)
    original_runtime_json = catalog.to_runtime_json

    def runtime_json() -> str:
        data = json.loads(original_runtime_json())
        for voice_package in data:
            if voice_package["id"] == "it-it-x-multi-seanet":
                voice_package["dependentVoiceId"] = "it-it-x-multi"
        return json.dumps(data)

    catalog.to_runtime_json = runtime_json
    speaker_id = os.environ.get(
        "PITCH_PROBE_SPEAKER", "it-it-x-multi-seanet:kda"
    )
    speaker = next(item for item in catalog.speakers if item.id == speaker_id)
    bridge = bridge_module.ChromeTtsBridge(catalog)
    probe_text = os.environ.get(
        "PITCH_PROBE_TEXT",
        "Questa frase verifica il controllo nativo del tono.",
    )
    output_suffix = os.environ.get("PITCH_PROBE_OUTPUT_SUFFIX", "")
    try:
        bridge.ensure_connection()
        time.sleep(1)
        pitch_specs = (("low", 0.4), ("normal", 1.0), ("high", 1.6))
        if os.environ.get("PITCH_PROBE_NORMAL_ONLY") == "1":
            pitch_specs = (("normal", 1.0),)
        for label, pitch in pitch_specs:
            audio_parts: list[bytes] = []
            options = {
                "voiceId": speaker.id,
                "voiceName": speaker.name,
                "lang": speaker.language,
                "rate": 1.175,
                "pitch": pitch,
                "volume": 1.0,
                "outputGain": 1.0,
            }
            probe_segments = os.environ.get("PITCH_PROBE_SEGMENTS")
            if probe_segments:
                segments = [part.strip() for part in probe_segments.split("|") if part.strip()]
                for index, segment in enumerate(segments):
                    bridge.speak(segment, options, audio_parts.append)
                    if index + 1 < len(segments):
                        audio_parts.append(b"\0\0" * 2_400)  # 100 ms at 24 kHz
            else:
                bridge.speak(
                    probe_text,
                    options,
                    audio_parts.append,
                )
            pcm = b"".join(audio_parts)
            with wave.open(
                str(variant_dir / f"{label}{output_suffix}.wav"), "wb"
            ) as output:
                output.setnchannels(1)
                output.setsampwidth(2)
                output.setframerate(24_000)
                output.writeframes(pcm)
            print(label, pitch, len(pcm), hashlib.sha256(pcm).hexdigest())
    finally:
        bridge.terminate()


if __name__ == "__main__":
    main()
