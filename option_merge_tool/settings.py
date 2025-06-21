from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True, kw_only=True)
class Settings:
    test_mode: bool
    log_file: str
    input_file: str
    template_file: str
    first_column: str
    second_column: str
    join_by: str
    output_column: str
    column_to_dropna: str
    columns_to_drop_dulicates: list[str]
