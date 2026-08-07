# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import (
    collect_data_files,
    collect_dynamic_libs,
    collect_submodules,
    copy_metadata,
)

root = Path(SPEC).resolve().parent
runtime = root / "audio_description_runtime"


def is_runtime_genai_module(name):
    return not (
        name == "google.genai.tests"
        or name.startswith("google.genai.tests.")
        or name == "google.genai._test_api_client"
    )


genai_datas = collect_data_files("google.genai")
genai_datas += copy_metadata("google-genai")
genai_binaries = collect_dynamic_libs("google.genai")
genai_hiddenimports = collect_submodules(
    "google.genai",
    filter=is_runtime_genai_module,
)

added_datas = [
    *genai_datas,
    (str(runtime / "audio_describer" / "assets" / "pyannote-segmentation"), "assets/pyannote-segmentation"),
    (str(runtime / "THIRD_PARTY_LICENSES"), "THIRD_PARTY_LICENSES"),
]
a = Analysis(
    [str(root / "audio_description_bridge.py")],
    pathex=[str(runtime)],
    binaries=genai_binaries,
    datas=added_datas,
    hiddenimports=[
        *genai_hiddenimports,
        *collect_submodules("google.api_core"),
        *collect_submodules("google.auth"),
        "onnxruntime",
        "numpy",
    ],
    excludes=[
        "google.genai.tests",
        "google.genai._test_api_client",
        "pytest",
        "_pytest",
        "tensorflow",
        "torch",
        "torchaudio",
        "torchvision",
        "torchcodec",
        "mypy",
        "IPython",
        "ipywidgets",
        "trio",
        "outcome",
        "mcp",
        "sentencepiece",
        "transformers",
        "brotli",
        "brotlicffi",
        "tornado",
        # Optional submodules imported conditionally by third-party packages.
        # Listing the dotted names prevents PyInstaller from treating them as
        # unresolved dependencies while keeping platform modules untouched.
        "IPython.core",
        "IPython.core.magic_arguments",
        "IPython.core.magic",
        "IPython.core.formatters",
        "IPython.display",
        "mypy.version",
        "mypy.util",
        "mypy.typevars",
        "mypy.types",
        "mypy.server",
        "mypy.server.trigger",
        "mypy.semanal",
        "mypy.plugins",
        "mypy.plugins.common",
        "mypy.plugin",
        "mypy.options",
        "mypy.nodes",
        "mypy.errorcodes",
        "mypy.typeops",
        "mypy.type_visitor",
        "mypy.state",
        "mypy.expandtype",
        "trio.testing",
        "trio.to_thread",
        "trio.socket",
        "trio.lowlevel",
        "trio.from_thread",
        "_pytest.outcomes",
        "mcp.types",
        "mcp.shared",
        "mcp.shared._httpx_utils",
        "mcp.client",
        "mcp.client.streamable_http",
        "tornado.concurrent",
        "pyannote",
        "scipy",
        "pandas",
        "sklearn",
        "matplotlib",
        "av",
        "yt_dlp",
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="audio_description_bridge",
    debug=False,
    strip=False,
    upx=True,
    console=True,
)
