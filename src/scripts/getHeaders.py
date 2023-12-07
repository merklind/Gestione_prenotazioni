import json
from pathlib import Path

from openpyxl.worksheet.worksheet import Worksheet

from constants import RIEPILOGO_WS, COLUMN_FILE
from src import MASTER_FILE
from utils.excel_utils import open_workbook, open_worksheet
from utils.path_utils import get_root_path


def main():
    MAX_COLUMN = 81
    colNameDict = dict()

    folderPath: Path = get_root_path()

    wb_path: Path = folderPath.joinpath(MASTER_FILE)
    columnLabelPath = folderPath.joinpath(COLUMN_FILE)

    wb, _ = open_workbook(wb_path, data_only=True)
    ws: Worksheet = open_worksheet(wb, RIEPILOGO_WS)

    for col in range(1, MAX_COLUMN + 1):
        colName = ws.cell(1, col).value
        colNameDict.update({colName: col})

    with open(f"{columnLabelPath}", "w", encoding="utf-8") as outputFile:
        json.dump(colNameDict, outputFile, indent=4)


main()
