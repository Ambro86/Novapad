from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

import audio_description_bridge as bridge
from audio_describer.core import audio_describer


ROOT = Path(__file__).resolve().parents[2]
WINDOW = (ROOT / "src" / "app_windows" / "audio_description_window.rs").read_text(
    encoding="utf-8"
)
RUST_PIPELINE = (ROOT / "src" / "audio_description.rs").read_text(encoding="utf-8")


class CharacterCatalogTests(unittest.TestCase):
    def test_bridge_forwards_saved_catalog_to_next_episode(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            media = root / "episode2.mkv"
            chunk = root / "chunk.mp4"
            media.write_bytes(b"video")
            chunk.write_bytes(b"chunk")
            captured = {}

            def generate(*_args, **kwargs):
                captured.update(kwargs)
                return (
                    [(1.0, 2.0, "Anna entra.")],
                    [
                        {
                            "id": "c1",
                            "name": "Anna",
                            "description": "Donna dai capelli scuri e occhi verdi",
                        },
                        {
                            "id": "c2",
                            "name": "Marco",
                            "description": "Uomo con barba corta",
                        },
                    ],
                    [],
                )

            request = {
                "input_path": str(media),
                "duration_sec": 20.0,
                "chunks": [
                    {"path": str(chunk), "start_sec": 0.0, "end_sec": 20.0}
                ],
                "audio_wav_path": "",
                "language": "it",
                "verbosity": "detailed",
                "allow_extended_pauses": True,
                "recognize_characters": True,
                "initial_character_glossary": [
                    {
                        "id": "c1",
                        "name": "  Anna  ",
                        "description": "Donna dai capelli scuri",
                    }
                ],
                "gemini_api_key": "test-key",
                "gemini_model": "gemini-test",
            }
            with mock.patch.object(bridge, "_configure_omni"), mock.patch.object(
                bridge, "_status"
            ), mock.patch.object(
                bridge.speech_detector, "format_intervals_for_prompt", return_value=""
            ), mock.patch.object(
                bridge.audio_describer,
                "generate_descriptions_chunked",
                side_effect=generate,
            ), mock.patch.object(
                bridge.audio_describer,
                "_remove_consecutive_duplicates",
                side_effect=lambda descriptions, _callback: descriptions,
            ), mock.patch.object(
                bridge.speech_detector,
                "align_descriptions_with_extended_pauses",
                side_effect=lambda descriptions, _intervals, _duration: (
                    descriptions,
                    0,
                    0,
                ),
            ), mock.patch.object(
                bridge.config_model,
                "get_setting",
                return_value="gemini-test",
            ):
                result = bridge.run(request)

            self.assertEqual(
                captured["initial_character_glossary"],
                [
                    {
                        "id": "c1",
                        "name": "Anna",
                        "description": "Donna dai capelli scuri",
                    }
                ],
            )
            self.assertEqual(
                [item["name"] for item in result["character_glossary"]],
                ["Anna", "Marco"],
            )

    def test_saved_characters_seed_chunk_continuity_and_are_updated(self):
        continuity = {}
        audio_describer._update_character_continuity(
            continuity,
            [
                {
                    "id": "c1",
                    "name": "Anna",
                    "description": "Donna dai capelli scuri",
                }
            ],
            max_characters=96,
        )
        audio_describer._update_character_continuity(
            continuity,
            [
                {
                    "id": "c1",
                    "name": "Anna",
                    "description": "Donna dai capelli scuri e occhi verdi",
                },
                {
                    "id": "c2",
                    "name": "Marco",
                    "description": "Uomo con barba corta",
                },
            ],
            max_characters=96,
        )
        self.assertEqual(set(continuity), {"id:c1", "id:c2"})
        self.assertIn("occhi verdi", continuity["id:c1"]["description"])
        self.assertEqual(continuity["id:c1"]["id"], "c1")


    def test_catalog_merge_rejects_repeated_and_corrupted_biography(self):
        existing = (
            "Padre di Flo, Anna è sua moglie. Uomo adulto sui quarant’anni, medico, "
            "alto e robusto, con capelli castano scuro corti, folta barba e baffi scuri."
        )
        observed = (
            "Padre di Dio, Anna è sua moglie. Uomo adulto sui quarant'anni, medico, "
            "alto e robusto, con capelli castano scuro corti, folta barba e baffi scuri. "
            "Padre di Flo, medico robusto con barba e baffi scuri."
        )
        self.assertEqual(
            audio_describer._merge_character_descriptions(existing, observed),
            existing,
        )

    def test_catalog_merge_appends_only_new_visual_information(self):
        existing = (
            "Madre di Flo, Franz e Jack e moglie di Ernest. Donna adulta con capelli "
            "castano-ramati raccolti ordinatamente dietro la testa."
        )
        observed = (
            "Madre di Flo, Franz e Jack e moglie di Ernest. "
            "Indossa un abito azzurro chiaro con colletto alto volantato."
        )
        merged = audio_describer._merge_character_descriptions(existing, observed)
        self.assertTrue(merged.startswith(existing))
        self.assertEqual(merged.count("Madre di Flo"), 1)
        self.assertIn("abito azzurro", merged)

    def test_saved_catalog_id_survives_shortened_gemini_alias(self):
        continuity = {}
        audio_describer._update_character_continuity(
            continuity,
            [
                {
                    "id": "anna_robinson",
                    "name": "Anna Robinson",
                    "description": "Madre di Flo con capelli raccolti e abiti ottocenteschi.",
                }
            ],
            max_characters=96,
        )
        audio_describer._update_character_continuity(
            continuity,
            [
                {
                    "id": "anna",
                    "name": "Anna",
                    "description": "Indossa un abito azzurro chiaro.",
                }
            ],
            max_characters=96,
        )
        self.assertEqual(list(continuity), ["id:anna_robinson"])
        anna = continuity["id:anna_robinson"]
        self.assertEqual(anna["id"], "anna_robinson")
        self.assertEqual(anna["name"], "Anna Robinson")
        self.assertIn("Madre di Flo", anna["description"])
        self.assertIn("abito azzurro", anna["description"])

    def test_short_name_matching_does_not_merge_ambiguous_eric_id_prefixes(self):
        continuity = {}
        audio_describer._update_character_continuity(
            continuity,
            [
                {"id": "eric_capretto", "name": "Capretto Eric", "description": "Capretto giovane."},
                {"id": "eric_beths", "name": "Eric Beths", "description": "Naufrago adulto."},
            ],
            max_characters=96,
        )
        audio_describer._update_character_continuity(
            continuity,
            [{"id": "eric", "name": "Eric", "description": "Figura vista nel filmato."}],
            max_characters=96,
        )
        self.assertEqual(len(continuity), 3)
        self.assertIn("id:eric", continuity)

    def test_initial_catalog_descriptions_are_not_truncated_to_240_characters(self):
        description = "Descrizione completa. " * 30
        normalized = bridge._normalise_initial_character_glossary(
            [{"id": "flo", "name": "Flo", "description": description}]
        )
        self.assertEqual(len(normalized), 1)
        self.assertGreater(len(normalized[0]["description"]), 240)
        self.assertEqual(
            normalized[0]["description"],
            " ".join(description.split()),
        )

    def test_prompt_explicitly_allows_prior_episode_catalog_names(self):
        _system, user = audio_describer._build_unified_prompts(
            "",
            "gemini-test",
            character_continuity_text='[{"name":"Anna","description":"capelli scuri"}]',
        )
        self.assertIn("SAVED SERIES CATALOG", user)
        self.assertIn("prior episode", user)
        self.assertIn("Anna", user)
        self.assertIn("AUTHORITATIVE", user)
        self.assertIn("EXACTLY the established `id`", user)
        self.assertIn("genuinely new visible appearance fact", user)
        self.assertIn("Do not 'correct' the catalog", user)

    def test_catalog_controls_depend_on_character_recognition(self):
        self.assertIn("fn update_character_catalog_visibility", WINDOW)
        self.assertIn(
            "let recognize = checkbox_checked(state.recognize_characters_checkbox);",
            WINDOW,
        )
        self.assertIn("if recognize { SW_SHOW } else { SW_HIDE }", WINDOW)
        self.assertIn("if keep { SW_SHOW } else { SW_HIDE }", WINDOW)
        self.assertIn("ID_RECOGNIZE_CHARACTERS if !state.running", WINDOW)
        self.assertIn("ID_KEEP_CHARACTER_CATALOG if !state.running", WINDOW)

    def test_rust_pipeline_loads_and_saves_catalog_around_generation(self):
        self.assertIn("initial_character_glossary", RUST_PIPELINE)
        self.assertIn("save_audio_description_character_catalog", RUST_PIPELINE)
        self.assertIn("analysis.character_glossary", RUST_PIPELINE)
        self.assertIn('join("Catalogs")', RUST_PIPELINE)


if __name__ == "__main__":
    unittest.main()
