
import pandas as pd

def load_dataset(file):
    if file.endswith(".csv"):
        df = pd.read_csv(file)
    elif file.endswith(".xlsx"):
        df = pd.read_excel(file)
    elif file.endswith(".json"):
        df = pd.read_json(file)
    else:
        df = pd.read_table(file)
    return df
