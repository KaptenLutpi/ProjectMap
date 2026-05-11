import pandas as pd
import os

def load_excel():
    file_path = os.path.join("..", "data", "data_pencairan1.xlsx")
    return pd.read_excel(file_path)
