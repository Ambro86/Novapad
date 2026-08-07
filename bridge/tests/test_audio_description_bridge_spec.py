import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SPEC = (ROOT / "bridge" / "audio_description_bridge.spec").read_text(encoding="utf-8")


class AudioDescriptionBridgeSpecTests(unittest.TestCase):
    def test_google_genai_test_modules_are_filtered_before_collection(self):
        self.assertIn("collect_data_files", SPEC)
        self.assertIn("collect_dynamic_libs", SPEC)
        self.assertIn("filter=is_runtime_genai_module", SPEC)
        self.assertIn('name == "google.genai.tests"', SPEC)
        self.assertIn('name.startswith("google.genai.tests.")', SPEC)
        self.assertIn('name == "google.genai._test_api_client"', SPEC)
        self.assertNotIn('collect_all("google.genai")', SPEC)

    def test_pytest_and_google_test_modules_are_excluded_from_analysis(self):
        self.assertIn('"google.genai.tests"', SPEC)
        self.assertIn('"google.genai._test_api_client"', SPEC)
        self.assertIn('"pytest"', SPEC)
        self.assertIn('"_pytest"', SPEC)

    def test_unused_optional_runtime_families_are_excluded(self):
        optional_modules = (
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
        )
        for module in optional_modules:
            with self.subTest(module=module):
                self.assertIn(f'"{module}"', SPEC)

    def test_optional_warning_submodules_are_excluded(self):
        optional_submodules = (
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
        )
        for module in optional_submodules:
            with self.subTest(module=module):
                self.assertIn(f'"{module}"', SPEC)

    def test_cross_platform_standard_library_modules_are_not_excluded(self):
        platform_modules = ("pwd", "grp", "posix", "fcntl", "termios", "resource")
        for module in platform_modules:
            with self.subTest(module=module):
                self.assertNotIn(f'"{module}"', SPEC)


if __name__ == "__main__":
    unittest.main()
