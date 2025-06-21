from __future__ import annotations

import os
import sys

from glob import glob
from typing import TYPE_CHECKING

from option_merge_tool.log import LOGGER_FORMAT_STR, logger
from option_merge_tool.merge import merge


if TYPE_CHECKING:
    from option_merge_tool.settings import Settings


async def run(settings: Settings) -> None:
    logger.remove()
    if settings.test_mode:
        logger.add(
            sys.stderr,
            format=LOGGER_FORMAT_STR,
            level="DEBUG",
            colorize=True,
            enqueue=True,
        )
    else:
        logger.add(
            sys.stderr,
            format=LOGGER_FORMAT_STR,
            level="INFO",
            colorize=True,
            enqueue=True,
        )
    logger.add(
        settings.log_file,
        format=LOGGER_FORMAT_STR,
        enqueue=True,
        encoding="utf-8-sig",
        level="DEBUG",
    )

    input_file = settings.input_file

    if not os.path.exists(settings.input_file):
        input_file = glob("INPUT_*.xlsx")[0]

    merge(
        input_file,
        settings.template_file,
        settings.output_column,
        settings.first_column,
        settings.second_column,
        settings.join_by,
        settings.column_to_dropna,
        settings.columns_to_drop_dulicates,
    )
