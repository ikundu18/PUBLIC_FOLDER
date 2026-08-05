import os
import re
import pandas as pd
import pdfplumber
from difflib import SequenceMatcher

# ======================================
# CONFIGURATION
# ======================================

REPORT_FOLDER = r"C:\Users\Jayantap\OneDrive - ICRA Analytics Ltd\Desktop\ALL REPORT\2 - Copy\INPUT_2"

OUTPUT_FILE = r"C:\Users\Jayantap\OneDrive - ICRA Analytics Ltd\Desktop\ALL REPORT\2 - Copy\OUTPUT_2\Column_Check_Output_3.xlsx"

# ======================================
# NORMALIZATION
# ======================================

def normalize_text(text):

    text = str(text).lower()

    text = re.sub(
        r'[^a-z0-9]+',
        ' ',
        text
    )

    text = re.sub(
        r'\s+',
        ' ',
        text
    )

    return text.strip()

# ======================================
# FUZZY MATCH
# ======================================

def similarity(text1, text2):

    return SequenceMatcher(
        None,
        normalize_text(text1),
        normalize_text(text2)
    ).ratio()

# ======================================
# TARGET COLUMNS
# ======================================

column_targets = {

    "cumulative_unpaid_net_wac_rate_carryover":
        "cumulative unpaid net wac rate carryover",

    "basis_risk":
        "basis risk",

    "basis_risk_paid":
        "basis risk paid",

    "basis_risk_unpaid":
        "basis risk unpaid",

    "cap_carryover_amount_unpaid":
        "cap carryover amount unpaid",

    "net_wac_shortfall_amounts":
        "net wac shortfall amounts"
}

# ======================================
# EXCEL HEADER READER
# ======================================

def read_excel_headers(filepath):

    headers = []

    try:

        excel_file = pd.ExcelFile(filepath)

        for sheet in excel_file.sheet_names:

            df = pd.read_excel(
                filepath,
                sheet_name=sheet,
                header=None,
                nrows=8
            )

            max_cols = df.shape[1]

            for col_idx in range(max_cols):

                parts = []

                for row_idx in range(min(8, len(df))):

                    value = df.iloc[row_idx, col_idx]

                    if pd.notna(value):

                        value = str(value).strip()

                        if value:
                            parts.append(value)

                header = " ".join(parts)

                header = re.sub(
                    r"\s+",
                    " ",
                    header
                ).strip()

                if header:
                    headers.append(header)

    except Exception as e:

        print(f"Error reading Excel: {filepath} -> {e}")

    return list(dict.fromkeys(headers))

# ======================================
# CSV HEADER READER
# ======================================

def read_csv_headers(filepath):

    headers = []

    try:

        df = pd.read_csv(
            filepath,
            header=None,
            nrows=8,
            low_memory=False
        )

        max_cols = df.shape[1]

        for col_idx in range(max_cols):

            parts = []

            for row_idx in range(min(8, len(df))):

                value = df.iloc[row_idx, col_idx]

                if pd.notna(value):

                    value = str(value).strip()

                    if value:
                        parts.append(value)

            header = " ".join(parts)

            header = re.sub(
                r"\s+",
                " ",
                header
            ).strip()

            if header:
                headers.append(header)

    except Exception as e:

        print(f"Error reading CSV: {filepath} -> {e}")

    return list(dict.fromkeys(headers))

# ======================================
# PDF HEADER READER
# ======================================

def read_pdf_headers(filepath):

    headers = []

    try:

        with pdfplumber.open(filepath) as pdf:

            for page in pdf.pages:

                # --------------------------
                # TABLE EXTRACTION
                # --------------------------

                tables = page.extract_tables()

                if tables:

                    for table in tables:

                        if not table:
                            continue

                        header_rows = table[:8]

                        try:

                            max_cols = max(
                                len(row)
                                for row in header_rows
                                if row
                            )

                        except:
                            continue

                        for col_idx in range(max_cols):

                            parts = []

                            for row in header_rows:

                                if row and col_idx < len(row):

                                    value = row[col_idx]

                                    if value:

                                        value = str(
                                            value
                                        ).strip()

                                        if value:
                                            parts.append(
                                                value
                                            )

                            header = " ".join(parts)

                            header = re.sub(
                                r"\s+",
                                " ",
                                header
                            ).strip()

                            if header:
                                headers.append(header)

                # --------------------------
                # PAGE TEXT EXTRACTION
                # --------------------------

                text = page.extract_text()

                if text:

                    lines = text.split("\n")

                    for line in lines:

                        line = line.strip()

                        if len(line) > 3:

                            headers.append(line)

                            for part in re.split(
                                r"[/|,;]",
                                line
                            ):

                                part = part.strip()

                                if len(part) > 3:
                                    headers.append(part)

    except Exception as e:

        print(f"Error reading PDF: {filepath} -> {e}")

    return list(dict.fromkeys(headers))

# ======================================
# COLUMN MATCHING
# ======================================

def check_columns(headers):

    result = {}

    normalized_headers = {
        h: normalize_text(h)
        for h in headers
    }

    for target_col, target_text in column_targets.items():

        target_norm = normalize_text(
            target_text
        )

        target_words = set(
            target_norm.split()
        )

        best_match = ""
        best_score = 0

        for header, header_norm in normalized_headers.items():

            matched_words = sum(
                1
                for word in target_words
                if word in header_norm
            )

            keyword_score = (
                matched_words
                / len(target_words)
            ) * 100

            fuzzy_score = similarity(
                target_norm,
                header_norm
            ) * 100

            final_score = (
                keyword_score * 0.80
                +
                fuzzy_score * 0.20
            )

            # Special Logic

            if target_col == "basis_risk":

                if (
                    "paid" in header_norm
                    or
                    "unpaid" in header_norm
                ):
                    continue

            if target_col == "basis_risk_paid":

                if not all(
                    x in header_norm
                    for x in ["basis", "risk", "paid"]
                ):
                    continue

            if target_col == "basis_risk_unpaid":

                if not all(
                    x in header_norm
                    for x in ["basis", "risk", "unpaid"]
                ):
                    continue

            if final_score > best_score:

                best_score = final_score
                best_match = header

        result[target_col] = (
            "Y"
            if best_score >= 75
            else "N"
        )

        result[
            f"{target_col}_matched_header"
        ] = (
            best_match
            if best_score >= 75
            else ""
        )

    return result

# ======================================
# MAIN PROCESS
# ======================================

output_rows = []

for file in os.listdir(REPORT_FOLDER):

    if file.startswith("~$"):
        continue

    filepath = os.path.join(
        REPORT_FOLDER,
        file
    )

    if not os.path.isfile(filepath):
        continue

    print(f"Processing: {file}")

    try:

        headers = []

        if file.lower().endswith(
            (".xlsx", ".xls")
        ):

            headers = read_excel_headers(
                filepath
            )

        elif file.lower().endswith(".csv"):

            headers = read_csv_headers(
                filepath
            )

        elif file.lower().endswith(".pdf"):

            headers = read_pdf_headers(
                filepath
            )

        else:
            continue

        result = check_columns(
            headers
        )

        row = {

            "File_Name": file,

            "Total_Headers_Found":
                len(headers),

            **result
        }

        output_rows.append(row)

    except Exception as e:

        print(
            f"Failed: {file} -> {e}"
        )

# ======================================
# EXPORT OUTPUT
# ======================================

output_df = pd.DataFrame(
    output_rows
)

os.makedirs(
    os.path.dirname(
        OUTPUT_FILE
    ),
    exist_ok=True
)

output_df.to_excel(
    OUTPUT_FILE,
    index=False,
    engine="openpyxl"
)

print("\nCompleted Successfully")
print(f"Output File: {OUTPUT_FILE}")