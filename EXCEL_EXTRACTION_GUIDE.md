# Guide to Extracting Unstructured Excel Files

Based on learning from the POC Mill OER Analysis example project, here's a comprehensive guide to extracting and normalizing unstructured Excel files.

## Overview: The Challenge

Unstructured Excel files often contain:
- **Multiple sheets** with different data layouts
- **Pivot tables** with complex row/column structures
- **Headers that span multiple rows** or are embedded in data
- **Missing values and inconsistent formatting**
- **Multiple data sections** side-by-side in the same sheet
- **Quality issues** like data not summing to expected totals

## Key Extraction Strategies

### 1. **Finding and Parsing Headers**

When headers are not in the standard first row, search for them:

```python
import pandas as pd
from pathlib import Path

def extract_mill_metadata(raw_file: Path) -> pd.DataFrame:
    """Extract data when headers are buried in the sheet."""
    df = pd.read_excel(raw_file, sheet_name="Sheet1", header=None)
    
    # Strategy: Search for a row containing known header keywords
    header_row = None
    for i, row in df.iterrows():
        if "PSM" in row.values and "Region" in row.values:
            header_row = i
            break
    
    if header_row is None:
        raise ValueError("Could not find header row")
    
    # Extract from header onward
    df_clean = df.iloc[header_row:].copy()
    df_clean.columns = df_clean.iloc[0]
    df_clean = df_clean.iloc[1:].reset_index(drop=True)
    
    # Clean column names (handle NaN columns)
    df_clean.columns = [
        str(c).strip() if pd.notna(c) else f"col_{i}" 
        for i, c in enumerate(df_clean.columns)
    ]
    
    return df_clean
```

**Key Techniques:**
- Read with `header=None` to get raw data
- Loop through rows looking for known keywords
- Set the found row as the actual header
- Clean and normalize column names after

### 2. **Parsing Pivot Tables with Hierarchical Structure**

When data has year/month hierarchies embedded in rows:

```python
def extract_nonlmm_oer(raw_file: Path) -> pd.DataFrame:
    """Parse pivot table with Year/Month rows and Mill Code columns."""
    df = pd.read_excel(raw_file, sheet_name="Trend OER 14", header=None)
    
    # Find the header row with mill codes
    header_row = None
    for i, row in df.iterrows():
        if "BAMM" in row.values or "SMGM" in row.values:
            header_row = i
            break
    
    # Extract mill codes from header
    mill_codes = df.iloc[header_row, 2:].dropna().tolist()
    mill_codes = [str(m).strip() for m in mill_codes if str(m).strip() != "nan"]
    
    records = []
    current_year = None
    
    # Parse year/month rows below header
    for i in range(header_row + 1, len(df)):
        row = df.iloc[i]
        
        # Detect year row (typically in column 1)
        year_val = row.iloc[1]
        if pd.notna(year_val):
            try:
                if isinstance(year_val, (int, float)) and year_val > 2000:
                    current_year = int(year_val)
                    continue
            except (ValueError, TypeError):
                pass
        
        # Detect month row (typically has month number 1-12 in column 1)
        month_val = row.iloc[1]
        if pd.notna(month_val) and current_year is not None:
            try:
                month = int(month_val)
                if 1 <= month <= 12:
                    # Extract values for each mill (from column 2 onward)
                    oer_values = row.iloc[2:2 + len(mill_codes)].tolist()
                    
                    for mill_idx, mill_code in enumerate(mill_codes):
                        oer_val = oer_values[mill_idx]
                        if pd.notna(oer_val) and oer_val != 0:
                            records.append({
                                "year": current_year,
                                "month": month,
                                "mill_code": mill_code,
                                "oer_actual": float(oer_val)
                            })
            except (ValueError, TypeError):
                pass
    
    df_result = pd.DataFrame(records)
    
    # Create proper datetime column
    if not df_result.empty:
        df_result["date"] = pd.to_datetime(
            df_result[["year", "month"]].assign(day=1)
        )
        df_result = df_result.sort_values(["mill_code", "date"]).reset_index(drop=True)
    
    return df_result
```

**Key Techniques:**
- Use row position indices (column 0, 1, 2) when header names are inconsistent
- Track state (current_year) as you iterate through rows
- Use row index positions instead of named columns: `row.iloc[1]`, `row.iloc[2:15]`
- Create proper datetime columns in the final dataframe

### 3. **Parsing Multiple Data Sections Side-by-Side**

When a sheet has multiple related data blocks in different column ranges:

```python
def extract_lmm_oer_with_hfc(raw_file: Path) -> pd.DataFrame:
    """Parse sheet with 3 data sections in different column ranges."""
    df = pd.read_excel(raw_file, sheet_name="LMM14+HFC", header=None)
    
    lmm_mills = [
        "HNAM", "INKM", "JLMM", "KUYM", "LIBM", "NSAM", 
        "PHLM", "PRDM", "SBYM", "SKOM", "SMLM", "SMRM", 
        "SRUM", "TNGM"
    ]
    
    # The sheet has 3 sections:
    # - Columns 2-15: CPOER values
    # - Columns 17-30: HFC adjustment values
    # - Columns 32-45: CPOER + HFC combined values
    
    records = []
    current_year = None
    
    for i in range(2, len(df)):
        row = df.iloc[i]
        
        # Extract year
        year_val = row.iloc[0]
        if pd.notna(year_val):
            try:
                if isinstance(year_val, (int, float)) and year_val >= 2022:
                    current_year = int(year_val)
            except (ValueError, TypeError):
                pass
        
        # Extract month
        month_val = row.iloc[1]
        if month_val == "Grand Total":
            continue
        
        if pd.notna(month_val) and current_year is not None:
            try:
                month = int(month_val)
                if 1 <= month <= 12:
                    # Extract from each section
                    cpoer_vals = row.iloc[2:16].tolist()      # Columns 2-15
                    hfc_vals = row.iloc[17:31].tolist()       # Columns 17-30
                    combined_vals = row.iloc[32:46].tolist()  # Columns 32-45
                    
                    for mill_idx, mill_code in enumerate(lmm_mills):
                        cpoer = cpoer_vals[mill_idx] if mill_idx < len(cpoer_vals) else None
                        hfc = hfc_vals[mill_idx] if mill_idx < len(hfc_vals) else None
                        combined = combined_vals[mill_idx] if mill_idx < len(combined_vals) else None
                        
                        if pd.notna(cpoer) and cpoer != 0:
                            records.append({
                                "year": current_year,
                                "month": month,
                                "mill_code": mill_code,
                                "cpoer": float(cpoer) if pd.notna(cpoer) else None,
                                "hfc": float(hfc) if pd.notna(hfc) and hfc != 0 else 0.0,
                                "oer_actual": float(combined) if pd.notna(combined) and combined != 0 else float(cpoer)
                            })
            except (ValueError, TypeError):
                pass
    
    df_result = pd.DataFrame(records)
    
    if not df_result.empty:
        df_result["date"] = pd.to_datetime(
            df_result[["year", "month"]].assign(day=1)
        )
        df_result = df_result.sort_values(["mill_code", "date"]).reset_index(drop=True)
    
    return df_result
```

**Key Techniques:**
- Map column ranges to specific sections using `row.iloc[start:end]`
- Process each section separately within the same loop
- Use explicit column indices rather than named columns

### 4. **Parsing Repeating Data Blocks**

When data has a repeating structure (e.g., each mill has a block of rows):

```python
def extract_ffb_source_mix(raw_file: Path) -> pd.DataFrame:
    """Parse sheet with repeating data blocks (one per mill)."""
    df = pd.read_excel(raw_file, sheet_name="FFB%", header=None)
    
    records = []
    row = 5  # Data starts at row 5
    
    while row < len(df):
        # Get mill code (column 2)
        mill_code = df.iloc[row, 2]
        if pd.isna(mill_code) or str(mill_code).strip() in ["", "nan"]:
            row += 1
            continue
        
        mill_code = str(mill_code).strip()
        
        # Find the "Inti" row (search within next 2 rows to be robust)
        inti_offset = None
        for offset in [1, 2]:
            if row + offset < len(df):
                label = df.iloc[row + offset, 3]
                if str(label).strip() == "Inti":
                    inti_offset = offset
                    break
        
        if inti_offset is None:
            row += 1
            continue
        
        # Get the data rows for this mill
        inti_row = df.iloc[row + inti_offset]
        plasma_row = df.iloc[row + inti_offset + 1]
        p3_row = df.iloc[row + inti_offset + 2]
        
        # Extract data for each month (columns 4-51 represent months across years)
        for col in range(4, 52):
            inti_val = inti_row[col]
            plasma_val = plasma_row[col]
            p3_val = p3_row[col]
            
            # Skip incomplete records
            if pd.isna(inti_val) or pd.isna(plasma_val) or pd.isna(p3_val):
                continue
            
            # Map column index to year/month
            # Columns 4-15: 2022 (Jan-Dec)
            # Columns 16-27: 2023 (Jan-Dec)
            # Columns 28-39: 2024 (Jan-Dec)
            # Columns 40-51: 2025 (Jan-Dec)
            if col <= 15:
                year = 2022
                month = col - 3
            elif col <= 27:
                year = 2023
                month = col - 15
            elif col <= 39:
                year = 2024
                month = col - 27
            else:
                year = 2025
                month = col - 39
            
            # Normalize to sum to 100% (data quality issue)
            total = inti_val + plasma_val + p3_val
            if total > 0:
                pct_inti = inti_val / total
                pct_plasma = plasma_val / total
                pct_3p = p3_val / total
            else:
                continue
            
            records.append({
                "year": year,
                "month": month,
                "mill_code": mill_code,
                "pct_inti": pct_inti,
                "pct_plasma": pct_plasma,
                "pct_3p": pct_3p
            })
        
        row += 4  # Each mill block is 4 rows, move to next mill
    
    df_result = pd.DataFrame(records)
    
    if not df_result.empty:
        df_result["date"] = pd.to_datetime(
            df_result[["year", "month"]].assign(day=1)
        )
        df_result = df_result.sort_values(["mill_code", "date"]).reset_index(drop=True)
    
    return df_result
```

**Key Techniques:**
- Use while loop with manual row advancement to handle repeating blocks
- Search for label rows to confirm structure (robust to minor variations)
- Build column-to-date mappings for month data spread across columns
- Normalize/validate data (e.g., ensure percentages sum to 100%)

### 5. **Data Type Conversions and Cleaning**

Always handle type conversions carefully:

```python
# Convert boolean columns
bool_cols = ["has_auto_p1", "has_hss", "has_eb_press"]
for col in bool_cols:
    if col in df.columns:
        df[col] = df[col].apply(
            lambda x: 1 if pd.notna(x) and x not in [0, "No", "no", ""] else 0
        )

# Convert numeric columns (coerce errors to NaN)
numeric_cols = ["capacity_mt", "utilization_pct"]
for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# Create datetime columns safely
df["date"] = pd.to_datetime(
    df[["year", "month"]].assign(day=1),
    errors="coerce"
)

# Handle NaN values explicitly
df[col] = df[col].fillna(value)
df = df[df[col].notna()]  # Filter out rows with missing values
```

## Workflow Pattern

A typical extraction workflow follows this pattern:

```python
def main():
    parser = argparse.ArgumentParser(description="Extract and normalize raw data")
    parser.add_argument("--input", "-i", type=str, help="Input raw Excel file")
    parser.add_argument("--output-dir", "-o", type=str, default="data/processed")
    args = parser.parse_args()
    
    # Find input file
    if args.input:
        raw_file = Path(args.input)
    else:
        raw_dir = Path("data/raws")
        raw_files = list(raw_dir.glob("raw_*.xlsx"))
        raw_file = sorted(raw_files)[-1]  # Get latest
    
    print(f"Processing: {raw_file}")
    
    # Extract version from filename
    version = re.search(r"(\d{8})", raw_file.name).group(1)
    
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Step 1: Extract metadata
    print("1. Extracting metadata...")
    df_metadata = extract_metadata(raw_file)
    df_metadata.to_csv(output_dir / f"01_metadata_{version}.csv", index=False)
    
    # Step 2: Extract time series data
    print("2. Extracting time series...")
    df_ts = extract_timeseries(raw_file)
    df_ts.to_csv(output_dir / f"02_timeseries_{version}.csv", index=False)
    
    # Step 3: Extract supplementary data
    print("3. Extracting supplementary data...")
    df_supp = extract_supplementary(raw_file)
    df_supp.to_csv(output_dir / f"03_supplementary_{version}.csv", index=False)
    
    print("✓ Complete!")

if __name__ == "__main__":
    main()
```

## Best Practices

1. **Read without headers first** (`header=None`) to inspect raw structure
2. **Use `row.iloc[index]`** for positional access when headers are unreliable
3. **Search for known keywords** to locate data sections dynamically
4. **Validate data quality** and normalize where needed (e.g., percentages)
5. **Create proper datetime columns** from year/month components
6. **Sort and reset indices** after building dataframes
7. **Handle NaN/None values** explicitly with type checks
8. **Use version strings** from filenames for traceability
9. **Add progress logging** for debugging multi-step extractions
10. **Catch exceptions** around type conversions to handle edge cases

## Common Pitfalls to Avoid

- ❌ Assuming fixed column positions without verification
- ❌ Not handling missing values (NaN) before type conversion
- ❌ Mixing string and numeric comparisons without explicit conversion
- ❌ Not creating proper datetime columns (keep year/month separate or combine properly)
- ❌ Forgetting to sort final dataframes chronologically
- ❌ Not validating extracted data against expected schemas
- ✅ Always inspect the raw Excel file first to understand structure
- ✅ Build extraction functions to be as robust as possible to minor variations
- ✅ Test with sample data before running on full dataset
