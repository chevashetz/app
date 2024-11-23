"""Полезные функции, которые ни к чему не принадлежат (общие)"""
import re

from PyQt6.QtWidgets import QTableWidget


def extract_number(text):
    text = text.replace(",", ".")
    match = re.search(r'\d+\.\d+', text)
    if match:
        return float(match.group(0))
    return float(text)

def countFilledRows(table: QTableWidget):
    filled_row_count = 0
    for row in range(table.rowCount()):
        for col in range(table.columnCount()):
            item = table.item(row, col)
            if not item or not item.text():
                break
        else:  # если for не прервался break
            filled_row_count += 1
    return filled_row_count