from __future__ import annotations

import os

from dataclasses import dataclass

import numpy as np
import pandas as pd

from excelsheet import col_to_excel, write_to_excel_template_cell_openpyxl
from openpyxl import load_workbook


@dataclass(slots=True, frozen=True)
class ExcelColumn:
    name: str
    alphabet: str


def update_column_mapping(
    old_columns: tuple[str, ...], new_columns: tuple[str, ...]
) -> dict[int, ExcelColumn]:
    """
    Updates the mapping of column indices to ExcelColumn objects based on old and new column names.
    Given tuples of old and new column names, this function generates a mapping from the index of each old column
    to an ExcelColumn object, preserving the original Excel-style column alphabet (e.g., 'A', 'B', ...).
    Only columns with matching names in both old and new tuples are included in the mapping.

    Args:
        old_columns (tuple[str, ...]): Tuple containing the names of the old columns.
        new_columns (tuple[str, ...]): Tuple containing the names of the new columns.
    Returns:
        dict[int, ExcelColumn]: A dictionary mapping the index of each old column (int) to an ExcelColumn object,
        for columns whose names exist in both old_columns and new_columns.
    """
    col_dict_old: dict[str, str] = {
        old_columns[x - 1]: col_to_excel(x) for x in range(1, len(old_columns) + 1)
    }

    excel_columns_old = [
        ExcelColumn(name, alphabet) for name, alphabet in col_dict_old.items()
    ]

    col_dict_new: dict[str, str] = {
        new_columns[x - 1]: col_to_excel(x) for x in range(1, len(new_columns) + 1)
    }

    excel_columns_new = [
        ExcelColumn(name, alphabet) for name, alphabet in col_dict_new.items()
    ]

    new_column_mapping: dict[int, ExcelColumn] = {}
    for idx, old_column in enumerate(excel_columns_old):
        for _, new_column in enumerate(excel_columns_new):
            if new_column.name == old_column.name:
                new_column_mapping[idx] = ExcelColumn(
                    old_column.name, old_column.alphabet
                )

    return new_column_mapping


def get_column_mapping(
    old_data: pd.DataFrame | str, new_data: pd.DataFrame | str
) -> dict[int, ExcelColumn]:
    if isinstance(old_data, pd.DataFrame):
        old_columns = tuple(old_data.columns)
    else:
        old_columns = tuple(pd.read_excel(old_data).columns)

    if isinstance(new_data, pd.DataFrame):
        new_columns = tuple(new_data.columns)
    else:
        new_columns = tuple(pd.read_excel(new_data).columns)

    return update_column_mapping(old_columns, new_columns)


def copy_to_openpyxl_template(
    *,
    dataframe: pd.DataFrame,
    filename: str,
    template_filename: str,
    column_mapping: dict[int, ExcelColumn],
):
    """
    Copies data from a pandas DataFrame into an Excel file using a specified OpenPyXL template and column mapping.
    Args:
        dataframe (pd.DataFrame): The DataFrame containing the data to be copied into the Excel template.
        filename (str): The output filename where the resulting Excel file will be saved. The file will be saved with a '.xlsx' extension.
        template_filename (str): The path to the Excel template file to be used as a base for the output.
        column_mapping (dict[int, ExcelColumn]): A dictionary mapping column indices to ExcelColumn objects, which specify the DataFrame column name and the corresponding Excel column (alphabet).
    Notes:
        - The function assumes that the template file exists and is accessible.
        - The function writes each specified DataFrame column to the corresponding Excel column as defined in the column_mapping.
        - The output file will overwrite any existing file with the same name.
    """
    wb = load_workbook(template_filename)
    ws = wb.active

    def write_df_column_to_excel_template_cell_openpyxl(name: str, alphabet: str):
        write_to_excel_template_cell_openpyxl(
            worksheet=ws, df=dataframe, name=name, alphabet=alphabet
        )

    for _, attr in column_mapping.items():
        alphabet = attr.alphabet
        name = attr.name
        write_df_column_to_excel_template_cell_openpyxl(name=name, alphabet=alphabet)

    wb.save(filename.replace(".csv", ".xlsx"))


def copy_dataframe_cells_to_excel_template(
    *,
    filename: str,
    template_filename: str,
    column_mapping: dict[int, ExcelColumn],
    column_to_dropna: str,
    current_os: str = "Windows",
):
    """
    Copies data from a DataFrame (loaded from a CSV or Excel file) into an Excel template.
    Parameters:
        filename (str): Path to the input data file (.csv or .xlsx).
        template_filename (str): Absolute path to the Excel template file.
        column_mapping (dict[int, ExcelColumn]): Mapping from DataFrame column indices to Excel columns.
        column_to_dropna (str): Name of the column to use for dropping rows with all NaN values.
        current_os (str, optional): Operating system name, defaults to "Windows".
    Raises:
        OSError: If the template_filename is not an absolute path on Windows.
    Notes:
        - The function reads the input file into a DataFrame, drops rows where the specified column is all NaN,
          replaces remaining NaN values with empty strings, and then copies the data to the Excel template.
        - Requires absolute path for the template file on Windows due to win32com limitations.
    """
    if not os.path.isabs(template_filename) and current_os == "Windows":
        raise OSError(f"Absolute path is needed for win32com (got {template_filename})")

    if filename.endswith(".xlsx"):
        dataframe = pd.read_excel(filename, engine="openpyxl", dtype="str")
    else:
        dataframe = pd.read_csv(filename, encoding="utf-8-sig", dtype="str")

    dataframe = dataframe.dropna(subset=[column_to_dropna], how="all")
    dataframe = dataframe.replace(np.nan, "", regex=True)

    copy_to_openpyxl_template(
        dataframe=dataframe,
        filename=filename,
        template_filename=template_filename,
        column_mapping=column_mapping,
    )
