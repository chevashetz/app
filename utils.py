"""Полезные функции, которые ни к чему не принадлежат (общие)"""
import re

from PyQt6.QtWidgets import QTableWidget


def extract_number(text):
    text = text.replace(",", ".")
    match = re.search(r'\d+\.\d+', text)
    if match:
        return float(match.group(0))
    return float(text)


