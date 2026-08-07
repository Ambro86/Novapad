"""Headless translation shim.

The Rust UI owns localization. The worker keeps Omni's source strings as
status/error details and sends stable stage identifiers alongside them.
"""
from __future__ import annotations


def _(text: str) -> str:
    return text


def initialize_translations() -> str:
    return "en"
