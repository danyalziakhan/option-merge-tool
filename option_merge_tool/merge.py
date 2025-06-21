from __future__ import annotations

import os
import sys

from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

import numpy as np
import pandas as pd

from option_merge_tool.excel import (
    copy_dataframe_cells_to_excel_template,
    get_column_mapping,
)
from option_merge_tool.log import logger


if TYPE_CHECKING:
    from typing import Any, Final

TODAY_DATE = f"{datetime.now().strftime('%Y%m%d')}"
SCRIPT_PATH: Final[str] = os.path.dirname(os.path.realpath(sys.argv[0]))


def read_excel(file: str):
    return pd.read_excel(file, engine="openpyxl", dtype=str)


def merge(
    input_file: str,
    template_file: str,
    output_column: str,
    first_column: str,
    second_column: str,
    join_by: str,
    column_to_dropna: str,
    columns_to_drop_dulicates: list[str],
):
    dataframe = read_excel(input_file)

    logger.log("ACTION", "Creating MERGED_OPTIONS.xlsx ...")

    output_dir = os.path.join(SCRIPT_PATH, "output", TODAY_DATE)
    os.makedirs(output_dir, exist_ok=True)

    output_filename = os.path.join(output_dir, "MERGED_OPTIONS.xlsx")
    columns = [first_column, second_column]
    aggregation_functions = {output_column: join_by.join}

    not_empty = (dataframe[column_to_dropna] != "") & (
        dataframe[column_to_dropna].notna()
    )
    dataframe = dataframe[
        not_empty
    ]  # ? Remove the line 2 row which previously included meta information for columns (it is now removed in latest Excel DB file)
    dataframe = (
        dataframe.reset_index(drop=True).sort_values(by=columns).reset_index(drop=True)
    )

    aggregated = (
        dataframe.astype(str)
        .groupby(columns, as_index=False)
        .agg(aggregation_functions)
        .reset_index(drop=True)
        .sort_values(by=columns)
        .reset_index(drop=True)
    )

    # ? Look for index overlap
    dataframe_with_index = dataframe.set_index(columns)
    aggregated_with_index = aggregated.set_index(columns)

    common_in_both = dataframe[
        dataframe_with_index.index.isin(aggregated_with_index.index)
    ]
    common_in_both = (
        common_in_both.reset_index(drop=True)
        .sort_values(by=columns)
        .drop_duplicates(subset=columns, keep="first")
        .reset_index(drop=True)
    )

    predicate = dataframe_with_index[: len(common_in_both)].index.isin(
        aggregated_with_index[: len(common_in_both)].index
    )
    new_only_items = common_in_both[predicate]

    new_only_items = (
        new_only_items.reset_index(drop=True)
        .sort_values(by=columns)
        .reset_index(drop=True)
    )

    aggregated = aggregated.sort_values(by=columns).reset_index(drop=True)

    new_only_items_iter: Any = new_only_items.iterrows()  # type: ignore
    for idx, new_only_item in new_only_items_iter:
        is_matched = (aggregated[first_column] == new_only_item.loc[first_column]) & (
            aggregated[second_column] == new_only_item.loc[second_column]
        )
        matched_aggregated: Any = aggregated[is_matched]  # type: ignore
        if len(matched_aggregated) >= 1:  # type: ignore
            new_only_items.loc[idx, output_column] = matched_aggregated[
                output_column
            ].iloc[0]

    dataframe_with_merged_options = new_only_items
    dataframe_with_merged_options.to_excel(output_filename, index=False)

    # ? We need to remove the already existing file if present, otherwise shutil.copy fails
    if os.path.exists(output_filename):
        os.remove(output_filename)

    dataframe_with_merged_options = dataframe_with_merged_options.dropna(
        subset=[column_to_dropna], how="all"
    )
    dataframe_with_merged_options = dataframe_with_merged_options[
        (dataframe_with_merged_options[column_to_dropna] != "")
        & (dataframe_with_merged_options[column_to_dropna].notna())
    ]  # ? Remove the line 2 row which previously included meta information for columns (it is now removed in latest Excel DB file)
    dataframe_with_merged_options = dataframe_with_merged_options.replace(
        np.nan, "", regex=True
    ).replace("nan", "", regex=True)

    dataframe_with_merged_options.drop_duplicates(
        subset=columns_to_drop_dulicates
    ).to_excel(
        output_filename,
        index=False,
        engine="openpyxl",
    )

    logger.log(
        "ACTION",
        f"Formatting {Path(output_filename).name} ... <yellow>(it may take a few seconds, so wait for it to be finished.)</>",
    )

    column_mapping = get_column_mapping(template_file, output_filename)

    copy_dataframe_cells_to_excel_template(
        filename=output_filename,
        template_filename=os.path.abspath(template_file),
        column_mapping=column_mapping,
        column_to_dropna=column_to_dropna,
    )

    logger.success(f"File saved to <CYAN><white>{output_filename}</></>")
