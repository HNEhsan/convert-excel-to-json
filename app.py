""" This script reads all Excel files in the input directory, cleans the data, and saves it as JSON files in the output directory. """

# Import required libraries
import os
import json
import re
import pandas as pd
from typing import Any

# Define input and output directories
INPUT_DIR = "./data/input/"
OUTPUT_DIR = "./data/output/"


# Define a function to clean and convert values
def clean_and_convert(value: Any) -> Any:
    """
    Cleans and converts Persian numbers to English, removes extra spaces, and handles missing values.

    Args:
        value (Any): Input value from the dataset.

    Returns:
        Any: Cleaned and converted value.
    """
    if pd.isna(value) or value in {"NaN", "nan", "None"}:
        return None
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        value = re.sub(
            r"[\u200c\u202c\u202d\u202e\u2066\u2067\u2068\u2069]", "", value.strip()
        )  # Remove hidden characters
        value = re.sub(r"\s+", " ", value)  # Remove extra spaces
        persian_digits = "۰۱۲۳۴۵۶۷۸۹"
        english_digits = "0123456789"
        value = value.translate(str.maketrans(persian_digits, english_digits))
        return (
            value.replace("“", '"')
            .replace("”", '"')
            .replace("‘", "'")
            .replace("’", "'")
            .strip()
        )
    return value


# Run App
if __name__ == "__main__":
    # Ensure the output directory exists
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Get a list of all Excel files in the input directory
    excel_files = [f for f in os.listdir(INPUT_DIR) if f.endswith((".xls", ".xlsx"))]

    # Check if there are any Excel files in the input directory
    if not excel_files:
        print("No Excel files found in the input directory.")
        exit()

    # Process each Excel file
    for file in excel_files:
        file_path = os.path.join(INPUT_DIR, file)
        output_file = os.path.join(OUTPUT_DIR, os.path.splitext(file)[0] + ".json")

        # Read all sheets in the Excel file
        excel_data = pd.read_excel(file_path, sheet_name=None, dtype=str)

        # Convert to dictionary
        json_data = {
            sheet: df.map(clean_and_convert).to_dict(orient="records")
            for sheet, df in excel_data.items()
        }

        # Save as JSON
        with open(output_file, "w", encoding="utf-8") as json_file:
            json.dump(json_data, json_file, ensure_ascii=False, indent=4)
