from modules import process_distributors, process_data_bases
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

if __name__ == "__main__":
    process_distributors()
    process_data_bases()