from openpyxl import Workbook
from .sheet_info import get_value_at_coordinate
from .utils import normalize
from datetime import date


def _get_process_date(workbook: Workbook):
    try:
        process_date_dn = workbook.defined_names['LnkTxtDRPData']
        
        for tab_name, cell_ref in process_date_dn.destinations:
            tab_origin = workbook[tab_name]
            return tab_origin[cell_ref].value
    except KeyError:
        pass
    except Exception as error:
        print(f"Unexpected error when reading defined name in {workbook}: {str(error)}")

    try:
        cover = workbook['CAPA']
        return get_value_at_coordinate('C10', cover)
    except Exception as error:
        print(f"Could not find process date on {workbook}: {str(error)}")
        return None
    

def get_process_year(workbook: Workbook):
    process_date = _get_process_date(workbook)

    if process_date and isinstance(process_date, date):
        return process_date.year
    else:
        return "-"
    

def get_contract_type(workbook, coordinate: str = 'C27'):
    try:
        cover = workbook['CAPA']
    except Exception as error:
        print(f"Could not get tab 'CAPA' at {workbook}: {str(error)}")
        return "-"

    try:
        contract_type = get_value_at_coordinate(coordinate, cover)
    except Exception as error:
        print(f"Could not find contract type on {workbook}: {str(error)}")
        return "-"
    
    if coordinate == 'C27':
        if normalize(contract_type) in ["ANTIGO", "NOVO"]:
            return contract_type
        else:
            return get_contract_type(
                workbook=workbook,
                coordinate='C28'
            )
        
    if not contract_type:
        return "-"
    
    return contract_type