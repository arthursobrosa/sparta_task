from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook
from datetime import datetime


def get_tab(tab_name: str, workbook: Workbook):
    try:
        tab = workbook[tab_name]
        return tab
    except KeyError:
        print(f"Erro: A aba {tab_name} não foi encontrada.")
        return None
    

def get_value_at_coordinate(coordinate: str, tab: Worksheet):
    cell = tab[coordinate]
    value = cell.value

    if "#REF!" in str(value):
        return None
    
    if value is isinstance(value, str):
        return value.strip()
    
    if isinstance(value, datetime):
        return value.date()
    
    return value