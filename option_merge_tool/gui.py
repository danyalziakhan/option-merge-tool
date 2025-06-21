from __future__ import annotations

import os
import sys

from dataclasses import dataclass, field
from enum import IntEnum, auto
from glob import glob
from pathlib import Path
from typing import TYPE_CHECKING, Protocol, cast

import dearpygui.dearpygui as dpg

from option_merge_tool.log import LOGGER_FORMAT_STR, logger
from option_merge_tool.merge import TODAY_DATE, merge, read_excel


if TYPE_CHECKING:
    from typing import Any

    import pandas as pd

    from option_merge_tool.settings import Settings


@dataclass(slots=True, kw_only=True)
class Configuration:
    template_file: str
    input_file: str
    first_column: str
    second_column: str
    join_by: str
    output_column: str
    today_date: str
    column_to_dropna: str
    columns_to_drop_dulicates: list[str]


class ElementTag(IntEnum):
    output_column = auto()
    first_column = auto()
    second_column = auto()
    join_by = auto()
    selected_data_file = auto()
    data_file_dialog = auto()
    selected_template_file = auto()
    template_file_dialog = auto()


# ? Attributes of SplitOptions that we want to pass around DearPyGUI elements
class StatefulData(Protocol):
    configuration: Configuration
    dataframe: pd.DataFrame


@dataclass(slots=True)
class SplitOptions:
    configuration: Configuration
    dataframe: pd.DataFrame = field(init=False)

    def input_file_selected(self, sender: str, app_data: dict[str, dict[str, str]]):
        print("Selecting")
        dpg.set_value(
            ElementTag.selected_data_file,
            list(app_data["selections"].values())[0],
        )
        self.configuration.input_file = list(app_data["selections"].values())[0]
        self.dataframe = read_excel(self.configuration.input_file)
        logger.success(f"File selected: {self.configuration.input_file}")

    def template_file_selected(self, sender: str, app_data: dict[str, dict[str, str]]):
        print("Selecting")
        dpg.set_value(
            ElementTag.selected_template_file,
            list(app_data["selections"].values())[0],
        )
        self.configuration.template_file = list(app_data["selections"].values())[0]
        logger.success(f"File selected: {self.configuration.template_file}")

    def create(self, font: str):
        if not os.path.exists(self.configuration.input_file):
            self.configuration.input_file = glob("SABANGNET_*.xlsx")[0]

        self.dataframe = read_excel(self.configuration.input_file)

        # Input file selection
        with dpg.file_dialog(
            directory_selector=False,
            show=False,
            callback=self.input_file_selected,
            tag=ElementTag.data_file_dialog,
            height=500,
            width=700,
        ):
            dpg.add_file_extension(".xlsx", color=(255, 255, 0, 255))
            dpg.add_file_extension(".csv", color=(255, 0, 255, 255))

        dpg.add_button(
            label="Choose the input data file",
            callback=lambda: dpg.show_item(ElementTag.data_file_dialog),
        )
        dpg.add_text(
            tag=ElementTag.selected_data_file,
            default_value=self.configuration.input_file,
            color=(255, 0, 0, 255),
        )

        # Template file selection
        with dpg.file_dialog(
            directory_selector=False,
            show=False,
            callback=self.template_file_selected,
            tag=ElementTag.template_file_dialog,
            height=500,
            width=700,
        ):
            dpg.add_file_extension(".xlsx", color=(255, 255, 0, 255))
            dpg.add_file_extension(".csv", color=(255, 0, 255, 255))

        dpg.add_button(
            label="Choose the template data file",
            callback=lambda: dpg.show_item(ElementTag.template_file_dialog),
        )
        dpg.add_text(
            tag=ElementTag.selected_template_file,
            default_value=self.configuration.template_file,
            color=(255, 0, 0, 255),
        )

        with dpg.group(width=800, horizontal_spacing=50):
            default = self.configuration.first_column.replace("\n", "  ")
            columns: list[str] = list(self.dataframe.columns)

            dpg.add_text(f"First Column [DEFAULT: {default}]")
            dpg.add_combo(
                items=[c.replace("\n", "  ") for c in columns],
                default_value=self.configuration.first_column.replace("\n", "  "),
                tag=ElementTag.first_column,
            )

            default = self.configuration.second_column.replace("\n", "  ")
            dpg.add_text(f"Second Column [DEFAULT: {default}]")
            dpg.add_combo(
                items=[c.replace("\n", "  ") for c in columns],
                default_value=self.configuration.second_column.replace("\n", "  "),
                tag=ElementTag.second_column,
            )

            default = self.configuration.output_column.replace("\n", "  ")
            dpg.add_text(f"Output Column [DEFAULT: {default}]")
            dpg.add_combo(
                items=[c.replace("\n", "  ") for c in columns],
                default_value=self.configuration.output_column.replace("\n", "  "),
                tag=ElementTag.output_column,
            )

        with dpg.group(width=800):
            dpg.add_text(f"Join by [DEFAULT: {self.configuration.join_by}]")
            dpg.add_input_text(
                default_value=self.configuration.join_by,
                tag=ElementTag.join_by,
            )

        dpg.add_button(
            label="Proceed",
            callback=proceed_callback,
            user_data=self,
        )

        dpg.bind_font(font)


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

    configuration = Configuration(
        template_file=settings.template_file,
        input_file=settings.input_file,
        first_column=settings.first_column,
        second_column=settings.second_column,
        join_by=settings.join_by,
        output_column=settings.output_column,
        today_date=TODAY_DATE,
        column_to_dropna=settings.column_to_dropna,
        columns_to_drop_dulicates=settings.columns_to_drop_dulicates,
    )
    split_options = SplitOptions(configuration)
    logger.info(f"Template file: <RED>{Path(configuration.template_file).name}</RED>")
    logger.info(f"Today's date: <BLUE><white>{configuration.today_date}</white></BLUE>")

    dpg.create_context()

    font: Any
    with dpg.font_registry(tag="korean"):
        with dpg.font("NanumBarunGothic.otf", 18) as font:  # type: ignore
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Default)
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Korean)
            dpg.add_font_range(0x3100, 0x3FF0)
            dpg.add_font_chars([0x3105, 0x3107, 0x3108])
            dpg.add_char_remap(0x3084, 0x0025)

    with dpg.window(tag="Primary Window", autosize=True):
        split_options.create(font)  # type: ignore

    dpg.create_viewport(title="Merge Options", width=832, height=600)
    dpg.setup_dearpygui()
    dpg.show_viewport()

    dpg.set_primary_window("Primary Window", True)
    dpg.start_dearpygui()
    dpg.destroy_context()


def proceed_callback(sender: Any, app_data: Any, stateful: StatefulData):
    JOIN_RULE = dpg.get_value(ElementTag.join_by)
    output_column = cast(str, dpg.get_value(ElementTag.output_column)).replace(
        "  ", "\n"
    )
    first_column = cast(str, dpg.get_value(ElementTag.first_column)).replace(
        r"  ", "\n"
    )
    second_column = cast(str, dpg.get_value(ElementTag.second_column)).replace(
        r"  ", "\n"
    )
    print(f"{first_column = }")
    print(f"{second_column = }")
    print(f"{output_column = }")
    start_rule: str = JOIN_RULE

    stateful.dataframe = read_excel(stateful.configuration.input_file)
    merge(
        stateful.configuration.input_file,
        stateful.configuration.template_file,
        output_column,
        first_column,
        second_column,
        start_rule,
        stateful.configuration.column_to_dropna,
        stateful.configuration.columns_to_drop_dulicates,
    )
