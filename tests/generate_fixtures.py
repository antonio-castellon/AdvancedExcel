#!/usr/bin/env python
"""
Generate complex reference Excel files for unit tests.

Run: python tests/generate_fixtures.py
"""

import os

from openpyxl import Workbook

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "data")


def create_multi_table_studies():
    """
    Realistic multi-block file.
    After pd.read_excel (header=0), the first row becomes column names.
    We put a title row, then the actual "Batch Number" keys appear in the data area.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Raw data"

    # Title row (will become column headers after read)
    ws["A1"] = "Header"
    ws["B1"] = "Val"
    ws["C1"] = "Extra1"
    ws["D1"] = "Extra2"

    # Block 1
    ws["A2"] = "Batch Number"
    ws["B2"] = "001"
    ws["A3"] = "Study No"
    ws["B3"] = "S-123"
    ws["C3"] = "Plan No"
    ws["D3"] = "P-001"

    # Table headers (real data headers)
    ws["A5"] = "Sample"
    ws["B5"] = "Replicate"
    ws["C5"] = "Value1"
    ws["D5"] = "Value2"
    ws["A6"] = "S1"
    ws["B6"] = "R1"
    ws["C6"] = 10.5
    ws["D6"] = 20.1
    ws["A7"] = "S1"
    ws["B7"] = "R2"
    ws["C7"] = 11.2
    ws["D7"] = 19.8

    # Separator
    ws["A9"] = ""

    # Block 2
    ws["A10"] = "Batch Number"
    ws["B10"] = "002"
    ws["A11"] = "Study No"
    ws["B11"] = "S-456"
    ws["C11"] = "Plan No"
    ws["D11"] = "P-002"

    ws["A13"] = "Sample"
    ws["B13"] = "Replicate"
    ws["C13"] = "Value1"
    ws["D13"] = "Value2"
    ws["A14"] = "S3"
    ws["B14"] = "R1"
    ws["C14"] = 15.0
    ws["D14"] = 25.5
    ws["A15"] = "S3"
    ws["B15"] = "R2"
    ws["C15"] = 14.8
    ws["D15"] = 26.1

    # Second sheet - more key-value + table
    ws2 = wb.create_sheet("Data Raw")
    ws2["A1"] = "TitleRow"  # becomes header
    ws2["B1"] = "Val"
    ws2["A2"] = "PLATEFORM"
    ws2["B2"] = "LC-MS"
    ws2["A3"] = "PARAMETER"
    ws2["B3"] = "m/z"
    ws2["A4"] = "UNIT"
    ws2["B4"] = "Da"

    ws2["A6"] = "Sample"
    ws2["B6"] = "Area"
    ws2["C6"] = "Height"
    ws2["A7"] = "S1"
    ws2["B7"] = 12345
    ws2["C7"] = 9876
    ws2["A8"] = "S2"
    ws2["B8"] = 23456
    ws2["C8"] = 8765

    ws2["A10"] = "TBD"
    ws2["B10"] = "tbd"

    out_path = os.path.join(OUTPUT_DIR, "multi_table_studies.xlsx")
    wb.save(out_path)
    print(f"Created: {out_path}")
    return out_path


def create_complex_key_value_blocks():
    """Scattered keys + repeated Batch blocks + pair values + tables."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Complex"

    # First rows become headers after read
    ws["A1"] = "ColA"
    ws["B1"] = "ColB"

    # Key values
    ws["A2"] = "Project"
    ws["B2"] = "ADV-2026"
    ws["A3"] = "Operator"
    ws["B3"] = "J. Smith"

    # Block 1
    ws["A5"] = "Batch Number"
    ws["B5"] = "BATCH-001"
    ws["A6"] = "Instrument"
    ws["B6"] = "QTOF-1"

    ws["A8"] = "ID"
    ws["B8"] = "Compound"
    ws["C8"] = "RT"
    ws["D8"] = "Area"
    ws["A9"] = "C001"
    ws["B9"] = "Caffeine"
    ws["C9"] = 2.34
    ws["D9"] = 45231

    # Block 2 (repeated key)
    ws["A11"] = "Batch Number"
    ws["B11"] = "BATCH-002"
    ws["A12"] = "Instrument"
    ws["B12"] = "QTOF-2"

    ws["A14"] = "ID"
    ws["B14"] = "Compound"
    ws["C14"] = "RT"
    ws["D14"] = "Area"
    ws["A15"] = "C003"
    ws["B15"] = "Ibuprofen"
    ws["C15"] = 4.56
    ws["D15"] = 67123

    # Summary section for pair extraction and getAllTables
    ws["A17"] = "Summary"
    ws["A18"] = "Total Samples"
    ws["B18"] = 42
    ws["A19"] = "Notes"
    ws["B19"] = "See attached"

    out_path = os.path.join(OUTPUT_DIR, "complex_key_blocks.xlsx")
    wb.save(out_path)
    print(f"Created: {out_path}")
    return out_path


if __name__ == "__main__":
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    create_multi_table_studies()
    create_complex_key_value_blocks()
    print("\nReference fixtures (re)generated.")
