Sonarpad Google TTS integration
===============================

This integration was developed for Sonarpad using the Google TTS For NVDA
add-on supplied as a technical reference.

Original add-on:
  Google TTS For NVDA 0.2
  Developers: Dao Duc Trung and Nguyen Anh Duc
  Project: https://github.com/nguyenanhduc09/Google-TTS-For-NVDA

The integration uses a managed headless Google Chrome process and the Chrome
WASM TTS engine assets. Voice packages are not bundled with Sonarpad: users
download them on demand from the URLs in the supplied voice catalog. Downloaded
packages are verified by size and SHA-256 checksum.

The existing Faster Whisper Python bridge is independent and is not modified by
this integration.

Distribution note
-----------------
The supplied NVDA add-on archive does not contain an overall license file for
the add-on or for all bundled Google WASM engine assets. Before distributing a
public Sonarpad build that embeds those assets, confirm that redistribution is
permitted by the respective rights holders and applicable service terms.
