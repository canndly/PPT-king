from __future__ import annotations

from pathlib import Path

import pytest

from ppt_king.cli import _build_parser, _run_doctor


def test_generate_defaults_to_materials_folder() -> None:
    args = _build_parser().parse_args(["generate"])

    assert args.command == "generate"
    assert args.source == "资料"


def test_doctor_ready(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    for key in (
        "PPT_KING_API_KEY",
        "DEEPSEEK_API_KEY",
        "OPENAI_API_KEY",
    ):
        monkeypatch.delenv(key, raising=False)
    (tmp_path / ".env").write_text("PPT_KING_API_KEY=sk-test-1234567890\n", encoding="utf-8")
    materials = tmp_path / "资料"
    materials.mkdir()
    (materials / "说明.md").write_text("# 资料\n", encoding="utf-8")

    args = _build_parser().parse_args(["doctor"])

    assert _run_doctor(args) == 0


def test_doctor_reports_missing_key_and_materials(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.chdir(tmp_path)
    for key in (
        "PPT_KING_API_KEY",
        "DEEPSEEK_API_KEY",
        "OPENAI_API_KEY",
    ):
        monkeypatch.delenv(key, raising=False)

    args = _build_parser().parse_args(["doctor"])

    assert _run_doctor(args) == 1
