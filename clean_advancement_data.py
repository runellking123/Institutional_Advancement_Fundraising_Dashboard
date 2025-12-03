"""
clean_advancement_data.py

A comprehensive data cleaning and transformation script for advancement/fundraising
CSV exports. Prepares data for Power BI with Copilot by creating star-schema
dimension and fact tables.

Usage:
    python clean_advancement_data.py

Author: Data Engineering Team
"""

from pathlib import Path
from typing import Optional
import re
import pandas as pd

# =============================================================================
# CONFIGURATION
# =============================================================================

# Edit this path to point to your raw CSV files folder
BASE_DIR = Path(r"C:\Users\ruking\Institutional_Advancement_Fundraising_Dashboard")

# Source file names
SOURCE_FILES = {
    "alumni_master": "ALUMNI_MASTER.csv",
    "campaign": "CAMPAIGN.csv",
    "donor_camp_sum": "DONOR_CAMP_SUM.csv",
    "donor_master": "DONOR_MASTER.csv",
    "donor_year_sum": "DONOR_YEAR_SUM.csv",
    "gift_category": "GIFT_CATEGORY.csv",
    "gift_master": "GIFT_MASTER.csv",
    "gift_tran": "GIFT_TRAN.csv",
    "name_master": "NAME_MASTER.csv",
    "solicit_def": "SOLICIT_DEF.csv",
}

# Columns to drop (technical metadata)
METADATA_COLUMNS_TO_DROP = ["USER_NAME", "JOB_NAME", "JOB_TIME", "APPROWVERSION"]

# ID/code columns that should always be treated as strings
ID_CODE_COLUMNS = [
    "ID_NUM",
    "DONOR_ID",
    "CAMPAIGN_CDE",
    "SOLICIT_CDE",
    "GIFT_GROUP_NUM",
    "GIFT_NUM",
    "GIFT_TRAN_NUM",
    "APPID",
    "CAT_COMP_1",
    "CAT_COMP_2",
]

# Natural keys for deduplication per table
NATURAL_KEYS = {
    "alumni_master": ["ID_NUM"],
    "campaign": ["CAMPAIGN_CDE"],
    "donor_camp_sum": ["ID_NUM", "CAMPAIGN_CDE", "YR_CDE", "YEAR_TYPE"],
    "donor_master": ["ID_NUM"],
    "donor_year_sum": ["ID_NUM", "YR_CDE", "YEAR_TYPE"],
    "gift_category": ["APPID", "CAT_COMP_1", "CAT_COMP_2"],
    "gift_master": ["GIFT_GROUP_NUM", "GIFT_NUM"],
    "gift_tran": ["GIFT_GROUP_NUM", "GIFT_NUM", "GIFT_TRAN_NUM"],
    "name_master": ["ID_NUM"],
    "solicit_def": ["SOLICIT_CDE"],
}

# Track parsing issues for summary
parsing_issues: dict[str, list[str]] = {}


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================


def to_upper_snake_case(column_name: str) -> str:
    """
    Convert a column name to upper snake case.

    Examples:
        "Id Num" -> "ID_NUM"
        "gift_amt" -> "GIFT_AMT"
        "FirstName" -> "FIRST_NAME"

    Args:
        column_name: The original column name.

    Returns:
        The column name in upper snake case.
    """
    # Strip leading/trailing whitespace
    name = column_name.strip()

    # Replace spaces with underscores
    name = name.replace(" ", "_")

    # Insert underscore before uppercase letters that follow lowercase letters
    name = re.sub(r"([a-z])([A-Z])", r"\1_\2", name)

    # Replace multiple underscores with single underscore
    name = re.sub(r"_+", "_", name)

    # Convert to uppercase
    return name.upper()


def standardize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardize all column names to upper snake case.

    Args:
        df: The input DataFrame.

    Returns:
        DataFrame with standardized column names.
    """
    df.columns = [to_upper_snake_case(col) for col in df.columns]
    return df


def strip_string_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Strip leading/trailing whitespace from all string columns.

    Args:
        df: The input DataFrame.

    Returns:
        DataFrame with stripped string values.
    """
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def drop_all_null_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Drop columns that are entirely null (all NaN values).

    Args:
        df: The input DataFrame.

    Returns:
        DataFrame with all-null columns removed.
    """
    return df.dropna(axis=1, how="all")


def drop_metadata_columns(df: pd.DataFrame, columns_to_drop: list[str]) -> pd.DataFrame:
    """
    Drop specified technical metadata columns if they exist.

    Args:
        df: The input DataFrame.
        columns_to_drop: List of column names to remove.

    Returns:
        DataFrame with metadata columns removed.
    """
    existing_cols_to_drop = [col for col in columns_to_drop if col in df.columns]
    return df.drop(columns=existing_cols_to_drop, errors="ignore")


def ensure_string_dtype(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """
    Ensure specified columns are loaded as string dtype and stripped of whitespace.

    Args:
        df: The input DataFrame.
        columns: List of column names to convert to string.

    Returns:
        DataFrame with specified columns as strings.
    """
    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            # Replace 'nan' strings (from NaN values) with empty string
            df[col] = df[col].replace("nan", "")
    return df


def deduplicate_by_key(df: pd.DataFrame, key_columns: list[str]) -> pd.DataFrame:
    """
    Deduplicate DataFrame based on specified natural key columns.
    Keeps the first occurrence of each unique key combination.

    Args:
        df: The input DataFrame.
        key_columns: List of columns forming the natural key.

    Returns:
        Deduplicated DataFrame.
    """
    existing_keys = [col for col in key_columns if col in df.columns]
    if existing_keys:
        return df.drop_duplicates(subset=existing_keys, keep="first")
    return df


def left_pad_id(value: str, width: int = 9) -> str:
    """
    Left-pad a numeric-looking ID with zeros to a specified width.
    This function is provided for optional use - it is NOT called automatically.

    Args:
        value: The ID value to pad.
        width: The target width (default 9 digits).

    Returns:
        Zero-padded ID string, or original value if not numeric.

    Example:
        left_pad_id("1234") -> "000001234"
        left_pad_id("ABC") -> "ABC"
    """
    if value and value.replace("-", "").isdigit():
        return value.zfill(width)
    return value


def parse_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Automatically convert date columns to datetime.

    Converts columns whose names:
    - End with '_DTE'
    - Contain 'DATE' or 'DTE'

    Also explicitly ensures common date fields are parsed correctly.

    Args:
        df: The input DataFrame.

    Returns:
        DataFrame with date columns converted to datetime.
    """
    # Explicit date columns to always parse
    explicit_date_cols = [
        "GIFT_DTE",
        "MATURITY_DTE",
        "FIRST_GIFT_DTE",
        "LAST_GIFT_DTE",
        "LARGEST_GIFT_DTE",
        "SMALLEST_GIFT_DTE",
        "CAMP_START_DATE",
        "CAMP_END_DATE",
        "VALID_UNTIL_DTE",
        "ACK_DATE",
        "EXPIRE_DTE",
        "LAST_UPDATE_DTE",
        "LAST_VISIT_DTE",
        "NEXT_VISIT_DTE",
        "REASON_DTE",
        "UDEF_DTE_1",
        "UDEF_DTE_2",
        "UDEF_DTE_3",
        "USER_DEF_DTE_1",
        "USER_DEF_DTE_2",
        "USER_DEF_DTE_3",
    ]

    date_pattern = re.compile(r"(_DTE$|DATE|DTE)", re.IGNORECASE)

    for col in df.columns:
        should_parse = (
            date_pattern.search(col) is not None or col in explicit_date_cols
        )

        if should_parse and col in df.columns:
            try:
                df[col] = pd.to_datetime(df[col], format="mixed", errors="coerce")
            except Exception:
                # Track parsing issues
                if "date_parsing" not in parsing_issues:
                    parsing_issues["date_parsing"] = []
                parsing_issues["date_parsing"].append(col)

    return df


def convert_numeric_columns(
    df: pd.DataFrame, include_patterns: list[str]
) -> pd.DataFrame:
    """
    Convert columns matching specified patterns to numeric type.

    Args:
        df: The input DataFrame.
        include_patterns: List of patterns to match column names (e.g., "AMT", "_NUM").

    Returns:
        DataFrame with matched columns converted to numeric.
    """
    for col in df.columns:
        for pattern in include_patterns:
            if pattern.upper() in col.upper():
                # Skip ID/code columns that should remain as strings
                if col in ID_CODE_COLUMNS:
                    continue
                try:
                    original_non_null = df[col].notna().sum()
                    df[col] = pd.to_numeric(df[col], errors="coerce")
                    new_non_null = df[col].notna().sum()

                    # Track if values were coerced to NaN
                    if new_non_null < original_non_null:
                        if "numeric_coercion" not in parsing_issues:
                            parsing_issues["numeric_coercion"] = []
                        parsing_issues["numeric_coercion"].append(
                            f"{col} ({original_non_null - new_non_null} values coerced to NaN)"
                        )
                except Exception:
                    if "numeric_conversion" not in parsing_issues:
                        parsing_issues["numeric_conversion"] = []
                    parsing_issues["numeric_conversion"].append(col)
                break

    return df


def load_csv_with_strings(filepath: Path, id_columns: list[str]) -> pd.DataFrame:
    """
    Load a CSV file ensuring specified ID columns are loaded as strings.

    Args:
        filepath: Path to the CSV file.
        id_columns: List of column names to load as strings.

    Returns:
        Loaded DataFrame with ID columns as strings.
    """
    # First, load to get column names
    df_sample = pd.read_csv(filepath, nrows=0)
    existing_id_cols = [col for col in id_columns if col in df_sample.columns]

    # Build dtype dict for string columns
    dtype_dict = {col: str for col in existing_id_cols}

    # Load full file with specified dtypes
    df = pd.read_csv(filepath, dtype=dtype_dict, low_memory=False)

    return df


# =============================================================================
# GLOBAL CLEANUP FUNCTION
# =============================================================================


def apply_global_cleanup(
    df: pd.DataFrame, table_name: str
) -> pd.DataFrame:
    """
    Apply all global cleanup rules to a DataFrame.

    Performs the following in order:
    1. Standardize column names to upper snake case
    2. Strip whitespace from string cells
    3. Ensure ID/code columns are strings
    4. Drop all-null columns
    5. Drop technical metadata columns
    6. Deduplicate based on natural key

    Args:
        df: The input DataFrame.
        table_name: The logical name of the table (for key lookup).

    Returns:
        Cleaned DataFrame.
    """
    # 1. Standardize column names
    df = standardize_column_names(df)

    # 2. Strip whitespace from strings
    df = strip_string_columns(df)

    # 3. Ensure ID/code columns are strings
    df = ensure_string_dtype(df, ID_CODE_COLUMNS)

    # 4. Drop all-null columns
    df = drop_all_null_columns(df)

    # 5. Drop metadata columns
    df = drop_metadata_columns(df, METADATA_COLUMNS_TO_DROP)

    # 6. Deduplicate
    if table_name in NATURAL_KEYS:
        df = deduplicate_by_key(df, NATURAL_KEYS[table_name])

    return df


# =============================================================================
# DATA LOADING FUNCTIONS
# =============================================================================


def load_all_source_files(base_dir: Path) -> dict[str, pd.DataFrame]:
    """
    Load all source CSV files into DataFrames.

    Args:
        base_dir: Base directory containing the source files.

    Returns:
        Dictionary mapping table names to DataFrames.
    """
    dataframes = {}

    for table_name, filename in SOURCE_FILES.items():
        filepath = base_dir / filename
        if filepath.exists():
            print(f"Loading {filename}...")
            df = load_csv_with_strings(filepath, ID_CODE_COLUMNS)
            df = apply_global_cleanup(df, table_name)
            df = parse_date_columns(df)
            dataframes[table_name] = df
            print(f"  Loaded {len(df):,} rows, {len(df.columns)} columns")
        else:
            print(f"WARNING: File not found: {filepath}")

    return dataframes


# =============================================================================
# DIMENSION TABLE BUILDERS
# =============================================================================


def build_dim_constituent(
    name_master: pd.DataFrame,
    donor_master: pd.DataFrame,
    alumni_master: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """
    Build DimConstituent dimension table.

    Sources: NAME_MASTER joined with DONOR_MASTER (and optionally ALUMNI_MASTER).
    Natural key: ID_NUM.

    Args:
        name_master: NAME_MASTER DataFrame.
        donor_master: DONOR_MASTER DataFrame.
        alumni_master: Optional ALUMNI_MASTER DataFrame for alumni flag.

    Returns:
        DimConstituent DataFrame with one row per ID_NUM.
    """
    print("Building DimConstituent...")

    # Select columns from NAME_MASTER
    name_cols = ["ID_NUM", "LAST_NAME", "FIRST_NAME", "MIDDLE_NAME", "PREFIX", "SUFFIX"]
    name_cols = [c for c in name_cols if c in name_master.columns]
    dim = name_master[name_cols].copy()

    # Select columns from DONOR_MASTER for join
    donor_cols_to_keep = [
        "ID_NUM",
        "PREF_MAIL_NM",
        "PREF_FST_NM",
        "PREF_LST_NM",
        "DONOR_MAIL_CODE",
        "ESTATE_PLAN_FLR",
        "ESTATE_PLAN_DOC",
        "YRS_CONT_GIVING",
        "NUM_GIVING_YRS",
        "AVG_GIFT_SIZE",
        "ANON_DONOR_FLAG",
        "DONOR_SELECT",
        "ACTIVE_FLAG",
        "STOP_MAIL_FLAG",
        "CURR_ADDR_CDE",
        "RECEIPT_SALUT",
        "SINGLE_SALUT",
        "JOINT_SALUT",
        "FIRST_GIFT_DTE",
        "FIRST_GIFT_AMT",
        "LAST_GIFT_DTE",
        "LAST_GIFT_AMT",
        "LARGEST_GIFT_DTE",
        "LARGEST_GIFT_AMT",
        "SMALLEST_GIFT_DTE",
        "SMALLEST_GIFT_AMT",
    ]
    donor_cols = [c for c in donor_cols_to_keep if c in donor_master.columns]
    donor_subset = donor_master[donor_cols].copy()

    # Convert numeric columns in donor subset
    donor_subset = convert_numeric_columns(donor_subset, ["AMT", "YRS", "NUM", "SIZE"])

    # Join on ID_NUM
    dim = dim.merge(donor_subset, on="ID_NUM", how="left")

    # If ALUMNI_MASTER available, add alumni flag
    if alumni_master is not None and "ID_NUM" in alumni_master.columns:
        alumni_ids = set(alumni_master["ID_NUM"].unique())
        dim["IS_ALUMNI"] = dim["ID_NUM"].isin(alumni_ids)

    # Deduplicate by ID_NUM
    dim = dim.drop_duplicates(subset=["ID_NUM"], keep="first")

    print(f"  Created {len(dim):,} constituent records")
    return dim


def build_dim_alumni(alumni_master: pd.DataFrame) -> pd.DataFrame:
    """
    Build DimAlumni dimension table.

    Source: ALUMNI_MASTER.
    Natural key: ID_NUM.

    Args:
        alumni_master: ALUMNI_MASTER DataFrame.

    Returns:
        DimAlumni DataFrame.
    """
    print("Building DimAlumni...")

    # Core alumni columns to keep
    core_cols = ["ID_NUM", "REUNION_YR_1", "REUNION_YR_2", "REUNION_YR_3"]

    # Education interest indicators
    edu_cols = [f"EDUCATION_INT_{i:02d}" for i in range(1, 7)]

    # Other valuable columns
    other_cols = [
        "LAST_UPDATE_DTE",
        "LAST_VISIT_DTE",
        "NEXT_VISIT_DTE",
        "CURR_ATTITUDE",
        "STOP_ALUM_MAIL",
        "CURR_ADDR_CDE",
        "WORK_ADDR_CDE",
        "EMAIL_1",
        "EMAIL_2",
        "CITY_SIZE",
    ]

    all_desired_cols = core_cols + edu_cols + other_cols
    available_cols = [c for c in all_desired_cols if c in alumni_master.columns]

    dim = alumni_master[available_cols].copy()

    # Deduplicate by ID_NUM
    dim = dim.drop_duplicates(subset=["ID_NUM"], keep="first")

    print(f"  Created {len(dim):,} alumni records")
    return dim


def build_dim_campaign(campaign: pd.DataFrame) -> pd.DataFrame:
    """
    Build DimCampaign dimension table.

    Source: CAMPAIGN.
    Key: CAMPAIGN_CDE.

    Args:
        campaign: CAMPAIGN DataFrame.

    Returns:
        DimCampaign DataFrame.
    """
    print("Building DimCampaign...")

    # Columns to keep
    cols_to_keep = [
        "CAMPAIGN_CDE",
        "CAMPAIGN_DESC",
        "PIECES_RET_GOAL",
        "PIECES_RET_ACTUAL",
        "EXPENSES_BUDGET",
        "EXPENSES_ACTUAL",
        "CAMPAN_AMT_GOAL",
        "CAMPAN_AMT_ACTUAL",
        "CAMP_CONTACT_ID_NUM",
        "CAMP_START_DATE",
        "CAMP_END_DATE",
        "ONLINE_GIVING_AVAIL",
        "ONLINE_GIVING_DESC",
    ]

    available_cols = [c for c in cols_to_keep if c in campaign.columns]
    dim = campaign[available_cols].copy()

    # Convert numeric columns
    dim = convert_numeric_columns(dim, ["AMT", "GOAL", "ACTUAL", "BUDGET", "PIECES"])

    # Deduplicate by CAMPAIGN_CDE
    dim = dim.drop_duplicates(subset=["CAMPAIGN_CDE"], keep="first")

    print(f"  Created {len(dim):,} campaign records")
    return dim


def build_dim_gift_category(gift_category: pd.DataFrame) -> pd.DataFrame:
    """
    Build DimGiftCategory dimension table.

    Source: GIFT_CATEGORY.
    Keys: APPID, CAT_COMP_1, CAT_COMP_2.

    Args:
        gift_category: GIFT_CATEGORY DataFrame.

    Returns:
        DimGiftCategory DataFrame.
    """
    print("Building DimGiftCategory...")

    # Columns to keep
    cols_to_keep = [
        "APPID",
        "CAT_COMP_1",
        "CAT_COMP_2",
        "GIFT_CAT_DESC",
        "CAMPAIGN_CDE",
        "FIN_AID_ELEMENT",
        "MEM_HONOR_CDE",
        "FUND_TYPE",
        "GIVING_CLUB_CDE",
        "DESIGNATION_CDE",
        "CASE_GROUP",
        "CHARITABLE_FLAG",
        "ONLINE_GIVING_AVAIL",
        "ONLINE_GIVING_DESC",
        "PROJECT_CODE",
    ]

    available_cols = [c for c in cols_to_keep if c in gift_category.columns]
    dim = gift_category[available_cols].copy()

    # Deduplicate by composite key
    key_cols = ["APPID", "CAT_COMP_1", "CAT_COMP_2"]
    existing_keys = [c for c in key_cols if c in dim.columns]
    if existing_keys:
        dim = dim.drop_duplicates(subset=existing_keys, keep="first")

    print(f"  Created {len(dim):,} gift category records")
    return dim


def build_dim_solicitation(solicit_def: pd.DataFrame) -> pd.DataFrame:
    """
    Build DimSolicitation dimension table.

    Source: SOLICIT_DEF.
    Key: SOLICIT_CDE.

    Args:
        solicit_def: SOLICIT_DEF DataFrame.

    Returns:
        DimSolicitation DataFrame.
    """
    print("Building DimSolicitation...")

    # Columns to keep
    cols_to_keep = ["SOLICIT_CDE", "DESCRIPTION"]

    available_cols = [c for c in cols_to_keep if c in solicit_def.columns]
    dim = solicit_def[available_cols].copy()

    # Deduplicate by SOLICIT_CDE
    dim = dim.drop_duplicates(subset=["SOLICIT_CDE"], keep="first")

    print(f"  Created {len(dim):,} solicitation records")
    return dim


# =============================================================================
# FACT TABLE BUILDERS
# =============================================================================


def build_fact_gift_transaction(
    gift_tran: pd.DataFrame, gift_master: pd.DataFrame
) -> pd.DataFrame:
    """
    Build FactGiftTransaction fact table.

    Sources: GIFT_TRAN (line-level) joined with GIFT_MASTER (header).
    Join on: GIFT_GROUP_NUM and GIFT_NUM.

    Args:
        gift_tran: GIFT_TRAN DataFrame.
        gift_master: GIFT_MASTER DataFrame.

    Returns:
        FactGiftTransaction DataFrame.
    """
    print("Building FactGiftTransaction...")

    # Columns from GIFT_TRAN
    tran_cols = [
        "GIFT_GROUP_NUM",
        "GIFT_NUM",
        "GIFT_TRAN_NUM",
        "DONOR_ID",
        "CAT_COMP_1",
        "CAT_COMP_2",
        "MEM_HONOR_CDE",
        "AID_ELEMENT",
        "CAMPAIGN_CDE",
        "SOLICIT_CDE",
        "GIVING_CLUB_CDE",
        "GIFT_DTE",
        "GIFT_TRAN_AMT",
        "GIVING_RELATION",
        "CHARITABLE_YN",
        "GIFT_TRAN_STS",
        "SOFT_CREDIT_YN",
        "GIFT_CLASS",
        "SUB_CLASS_CDE",
        "GIFT_SET_CDE",
        "ANON_GIFT_TRAN",
        "NOTATION_1",
        "NOTATION_2",
        "RESTR_TYPE",
        "CONTRIB_TYPE",
        "MATURITY_AMT",
        "MATURITY_DTE",
        "MATCH_CO_ID",
        "SOLICITOR_ID",
    ]

    available_tran_cols = [c for c in tran_cols if c in gift_tran.columns]
    fact = gift_tran[available_tran_cols].copy()

    # Convert numeric columns in transactions
    fact = convert_numeric_columns(fact, ["AMT"])

    # Columns from GIFT_MASTER to add
    master_cols = [
        "GIFT_GROUP_NUM",
        "GIFT_NUM",
        "GIFT_AMT",
        "LETTER_TO_SEND",
        "PROPOSAL_NUM",
        "BANK_ID",
        "CHECK_CC_NUM",
        "LEGAL_TENDER",
        "EXPIRE_DTE",
        "ACK_DATE",
        "PRINT_RECPT_YN",
        "IMMED_LETTER_YN",
        "BOOK_YN",
        "GIFT_MASTER_STS",
    ]

    available_master_cols = [c for c in master_cols if c in gift_master.columns]
    master_subset = gift_master[available_master_cols].copy()

    # Convert GIFT_AMT to numeric
    master_subset = convert_numeric_columns(master_subset, ["AMT"])

    # Deduplicate master for join (keep only relevant columns)
    join_keys = ["GIFT_GROUP_NUM", "GIFT_NUM"]
    cols_to_add = [c for c in available_master_cols if c not in join_keys]

    if all(k in master_subset.columns for k in join_keys):
        master_for_join = master_subset.drop_duplicates(subset=join_keys, keep="first")

        # Merge
        fact = fact.merge(
            master_for_join,
            on=join_keys,
            how="left",
            suffixes=("", "_MASTER"),
        )

    # Remove rows with missing critical keys
    if "DONOR_ID" in fact.columns:
        before_count = len(fact)
        fact = fact[fact["DONOR_ID"].notna() & (fact["DONOR_ID"] != "")]
        removed = before_count - len(fact)
        if removed > 0:
            print(f"  Removed {removed:,} rows with missing DONOR_ID")

    if "GIFT_NUM" in fact.columns:
        before_count = len(fact)
        fact = fact[fact["GIFT_NUM"].notna() & (fact["GIFT_NUM"] != "")]
        removed = before_count - len(fact)
        if removed > 0:
            print(f"  Removed {removed:,} rows with missing GIFT_NUM")

    print(f"  Created {len(fact):,} gift transaction records")
    return fact


def build_fact_donor_year_summary(donor_year_sum: pd.DataFrame) -> pd.DataFrame:
    """
    Build FactDonorYearSummary fact table.

    Source: DONOR_YEAR_SUM.
    Natural key: ID_NUM, YR_CDE, YEAR_TYPE.

    Args:
        donor_year_sum: DONOR_YEAR_SUM DataFrame.

    Returns:
        FactDonorYearSummary DataFrame.
    """
    print("Building FactDonorYearSummary...")

    fact = donor_year_sum.copy()

    # Convert numeric columns (amounts and counts)
    fact = convert_numeric_columns(fact, ["AMT", "_NUM"])

    print(f"  Created {len(fact):,} donor year summary records")
    return fact


def build_fact_donor_campaign_summary(donor_camp_sum: pd.DataFrame) -> pd.DataFrame:
    """
    Build FactDonorCampaignSummary fact table.

    Source: DONOR_CAMP_SUM.
    Natural key: ID_NUM, CAMPAIGN_CDE, YR_CDE, YEAR_TYPE.

    Args:
        donor_camp_sum: DONOR_CAMP_SUM DataFrame.

    Returns:
        FactDonorCampaignSummary DataFrame.
    """
    print("Building FactDonorCampaignSummary...")

    fact = donor_camp_sum.copy()

    # Convert numeric columns (amounts and counts)
    fact = convert_numeric_columns(fact, ["AMT", "_NUM"])

    print(f"  Created {len(fact):,} donor campaign summary records")
    return fact


# =============================================================================
# EXPORT FUNCTIONS
# =============================================================================


def export_tables(tables: dict[str, pd.DataFrame], output_dir: Path) -> None:
    """
    Export all cleaned tables to Excel files.

    Args:
        tables: Dictionary mapping table names to DataFrames.
        output_dir: Directory to write output files to.
    """
    # Create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nExporting tables to {output_dir}...")

    for table_name, df in tables.items():
        output_path = output_dir / f"{table_name}.xlsx"
        df.to_excel(output_path, index=False, engine="openpyxl")
        print(f"  Exported {table_name}.xlsx ({len(df):,} rows)")


def print_summary(tables: dict[str, pd.DataFrame]) -> None:
    """
    Print a summary of all tables and any parsing issues.

    Args:
        tables: Dictionary mapping table names to DataFrames.
    """
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)

    print("\nFinal Table Row Counts:")
    print("-" * 40)
    for table_name, df in tables.items():
        print(f"  {table_name}: {len(df):,} rows, {len(df.columns)} columns")

    if parsing_issues:
        print("\nParsing/Conversion Issues:")
        print("-" * 40)
        for issue_type, columns in parsing_issues.items():
            print(f"\n  {issue_type}:")
            for col in columns:
                print(f"    - {col}")
    else:
        print("\nNo parsing or conversion issues detected.")

    print("\n" + "=" * 60)


# =============================================================================
# MAIN EXECUTION
# =============================================================================


def main() -> None:
    """
    Main execution function.

    Orchestrates the full data cleaning and transformation pipeline:
    1. Load all source files
    2. Build dimension tables
    3. Build fact tables
    4. Export all cleaned tables
    5. Print summary
    """
    print("=" * 60)
    print("ADVANCEMENT DATA CLEANING PIPELINE")
    print("=" * 60)
    print(f"Base directory: {BASE_DIR}")
    print()

    # Validate base directory exists
    if not BASE_DIR.exists():
        raise FileNotFoundError(f"Base directory does not exist: {BASE_DIR}")

    # Step 1: Load all source files
    print("STEP 1: Loading source files...")
    print("-" * 40)
    dataframes = load_all_source_files(BASE_DIR)
    print()

    # Check we have required files
    required_tables = ["name_master", "donor_master", "gift_tran", "gift_master"]
    missing = [t for t in required_tables if t not in dataframes]
    if missing:
        print(f"WARNING: Missing required tables: {missing}")

    # Step 2: Build dimension tables
    print("STEP 2: Building dimension tables...")
    print("-" * 40)

    output_tables = {}

    # DimConstituent
    if "name_master" in dataframes and "donor_master" in dataframes:
        output_tables["DimConstituent"] = build_dim_constituent(
            dataframes["name_master"],
            dataframes["donor_master"],
            dataframes.get("alumni_master"),
        )

    # DimAlumni
    if "alumni_master" in dataframes:
        output_tables["DimAlumni"] = build_dim_alumni(dataframes["alumni_master"])

    # DimCampaign
    if "campaign" in dataframes:
        output_tables["DimCampaign"] = build_dim_campaign(dataframes["campaign"])

    # DimGiftCategory
    if "gift_category" in dataframes:
        output_tables["DimGiftCategory"] = build_dim_gift_category(
            dataframes["gift_category"]
        )

    # DimSolicitation
    if "solicit_def" in dataframes:
        output_tables["DimSolicitation"] = build_dim_solicitation(
            dataframes["solicit_def"]
        )

    print()

    # Step 3: Build fact tables
    print("STEP 3: Building fact tables...")
    print("-" * 40)

    # FactGiftTransaction
    if "gift_tran" in dataframes and "gift_master" in dataframes:
        output_tables["FactGiftTransaction"] = build_fact_gift_transaction(
            dataframes["gift_tran"], dataframes["gift_master"]
        )

    # FactDonorYearSummary
    if "donor_year_sum" in dataframes:
        output_tables["FactDonorYearSummary"] = build_fact_donor_year_summary(
            dataframes["donor_year_sum"]
        )

    # FactDonorCampaignSummary
    if "donor_camp_sum" in dataframes:
        output_tables["FactDonorCampaignSummary"] = build_fact_donor_campaign_summary(
            dataframes["donor_camp_sum"]
        )

    print()

    # Step 4: Export all tables
    print("STEP 4: Exporting cleaned tables...")
    print("-" * 40)
    output_dir = BASE_DIR / "cleaned"
    export_tables(output_tables, output_dir)

    # Step 5: Print summary
    print_summary(output_tables)

    print("\nPipeline completed successfully!")
    print(f"Cleaned files are available in: {output_dir}")


if __name__ == "__main__":
    main()
