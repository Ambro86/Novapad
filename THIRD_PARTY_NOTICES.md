# Third-Party Notices

This project may distribute and/or download at runtime third-party binaries and data files.
Each component remains under its own license and terms.

The project license (`LICENSE`) does **not** replace third-party licenses.

## Included in `dll/`

- `bass.dll`
- `bass_aac.dll`
- `bass_alac.dll`
- `bass_fx.dll`
- `bassflac.dll`
- `bassopus.dll`
- `basswma.dll`
  - Vendor: Un4seen Developments
  - Product: BASS Audio Library and add-ons
  - License/terms: see vendor website and license files provided by vendor
  - URL: https://www.un4seen.com/

- `libcurl.dll`
  - Project: curl
  - License: curl license
  - URL: https://curl.se/

- `zlib.dll`
  - Project: zlib
  - License: zlib license
  - URL: https://zlib.net/

- `pdfium.dll`
  - Project: PDFium
  - License: BSD-style (as distributed by upstream)
  - URL: https://pdfium.googlesource.com/pdfium/

- `nvdaControllerClient64.dll`
  - Project: NVDA
  - License: GPL (NVDA project terms)
  - URL: https://www.nvaccess.org/

- `sapi4_bridge_32.exe`
- `faster_whisper_bridge.exe`
  - Built for this project; they may embed/use third-party Python/Rust dependencies at build/runtime.

- `whisper-cuda-runtime-win64-cu12.zip`
  - NVIDIA CUDA/cuDNN runtime redistribution package (for optional GPU transcription)
  - License/terms: NVIDIA redistribution terms
  - URL: https://developer.nvidia.com/

- `cacert.pem`
  - CA certificate bundle from cURL/CA Extract source
  - URL: https://curl.se/docs/caextract.html

## Runtime-downloaded components

The application may download additional assets, for example:

- Whisper/faster-whisper model files from Hugging Face
- Optional CUDA runtime package from this repository releases/files
- Bridge helper binaries from this repository

These files keep their original licenses and terms from their upstream providers.

## Compliance note

If you redistribute binaries, you are responsible for:

- verifying each third-party license and redistribution permission,
- shipping required notices/license texts,
- complying with attribution and commercial/non-commercial restrictions where applicable.
