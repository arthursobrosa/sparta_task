from pathlib import Path
import unicodedata
import json


def get_suffix(file_name: str) -> str:
    path = Path(file_name)
    return path.suffix


def normalize(text: str) -> str:
    if not isinstance(text, str):
        return ""
    
    text = text.strip().upper()

    return ''.join(
        c for c in unicodedata.normalize('NFD', text)
        if unicodedata.category(c) != 'Mn'
    )


def get_json_data(json_path: str):
    with open(json_path, 'r', encoding='utf-8') as file:
        json_data = json.load(file)
        return json_data
