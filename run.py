from __future__ import annotations

import asyncio
import os

from argparse import ArgumentParser

from option_merge_tool.log import logger
from option_merge_tool.merge import TODAY_DATE
from option_merge_tool.settings import Settings


if __name__ == "__main__":
    parser = ArgumentParser()

    parser.add_argument(
        "--gui",
        help="Run the program in GUI mode",
        action="store_true",
    )
    parser.add_argument(
        "--test_mode",
        help="Run the program with verbose logging output",
        action="store_true",
    )
    parser.add_argument(
        "--log_file",
        help="Log file path which will override the default path",
        type=str,
        default=os.path.join("logs", f"{TODAY_DATE}.log"),
    )
    parser.add_argument(
        "--input_file",
        help="Input file",
        type=str,
        required=True,
    )
    parser.add_argument(
        "--template_file",
        help="Excel template file",
        type=str,
        required=True,
    )
    parser.add_argument(
        "--first_column",
        help="First Column",
        type=str,
        required=True,
    )
    parser.add_argument(
        "--second_column",
        help="Second Column",
        type=str,
        required=True,
    )
    parser.add_argument(
        "--join_by",
        help="Join by rule",
        type=str,
        required=True,
    )
    parser.add_argument(
        "--output_column",
        help="Column for saving merged options",
        type=str,
        required=True,
    )
    parser.add_argument(
        "--column_to_dropna",
        help="Column to drop NaN values",
        type=str,
        required=True,
    )
    parser.add_argument(
        "--columns_to_drop_dulicates",
        help="List of column names (separated by ,) to drop duplicates",
        type=str,
        required=True,
    )
    args = parser.parse_args()

    os.makedirs("logs", exist_ok=True)
    os.makedirs("output", exist_ok=True)

    settings = Settings(
        test_mode=args.test_mode,
        log_file=args.log_file,
        input_file=args.input_file,
        template_file=args.template_file,
        first_column=args.first_column.replace("\\n", "\n"),
        second_column=args.second_column.replace("\\n", "\n"),
        join_by=args.join_by,
        output_column=args.output_column.replace("\\n", "\n"),
        column_to_dropna=args.column_to_dropna.replace("\\n", "\n"),
        columns_to_drop_dulicates=args.columns_to_drop_dulicates.replace(
            "\\n", "\n"
        ).split(","),
    )

    if args.gui:
        from option_merge_tool.gui import run
    else:
        from option_merge_tool.non_gui import run

    try:
        asyncio.run(run(settings))
    except Exception as err:
        logger.log("UNHANDLED ERROR", err)
        raise err from err

    logger.success("Program has been run successfully")
