"""
rename_columns.py - Rename columns to human-readable names for Power BI Copilot
"""
import pandas as pd
from pathlib import Path

cleaned_dir = Path("cleaned")
renamed_dir = Path("Renamed Excel Model Files")

# ============================================================
# COLUMN RENAME MAPPINGS
# ============================================================

dim_constituent_renames = {
    "ID_NUM": "Constituent ID",
    "LAST_NAME": "Last Name",
    "FIRST_NAME": "First Name",
    "MIDDLE_NAME": "Middle Name",
    "PREFIX": "Prefix",
    "SUFFIX": "Suffix",
    "YRS_CONT_GIVING": "Years Continuous Giving",
    "NUM_GIVING_YRS": "Total Giving Years",
    "AVG_GIFT_SIZE": "Average Gift Size",
    "ANON_DONOR_FLAG": "Is Anonymous Donor",
    "ACTIVE_FLAG": "Is Active",
    "STOP_MAIL_FLAG": "Stop Mail",
    "CURR_ADDR_CDE": "Current Address Code",
    "FIRST_GIFT_DTE": "First Gift Date",
    "FIRST_GIFT_AMT": "First Gift Amount",
    "LAST_GIFT_DTE": "Last Gift Date",
    "LAST_GIFT_AMT": "Last Gift Amount",
    "LARGEST_GIFT_DTE": "Largest Gift Date",
    "LARGEST_GIFT_AMT": "Largest Gift Amount",
    "SMALLEST_GIFT_DTE": "Smallest Gift Date",
    "SMALLEST_GIFT_AMT": "Smallest Gift Amount",
    "IS_ALUMNI": "Is Alumni",
}

dim_alumni_renames = {
    "ID_NUM": "Constituent ID",
    "REUNION_YR_1": "Reunion Year 1",
    "REUNION_YR_2": "Reunion Year 2",
    "REUNION_YR_3": "Reunion Year 3",
    "EDUCATION_INT_01": "Education Interest 1",
    "EDUCATION_INT_02": "Education Interest 2",
    "EDUCATION_INT_03": "Education Interest 3",
    "EDUCATION_INT_04": "Education Interest 4",
    "EDUCATION_INT_05": "Education Interest 5",
    "EDUCATION_INT_06": "Education Interest 6",
    "LAST_UPDATE_DTE": "Last Update Date",
    "LAST_VISIT_DTE": "Last Visit Date",
    "NEXT_VISIT_DTE": "Next Visit Date",
    "CURR_ATTITUDE": "Current Attitude",
    "STOP_ALUM_MAIL": "Stop Alumni Mail",
    "CURR_ADDR_CDE": "Current Address Code",
    "WORK_ADDR_CDE": "Work Address Code",
    "EMAIL_1": "Email Primary",
    "EMAIL_2": "Email Secondary",
    "CITY_SIZE": "City Size",
}

dim_campaign_renames = {
    "CAMPAIGN_CDE": "Campaign Code",
    "CAMPAIGN_DESC": "Campaign Name",
    "PIECES_RET_GOAL": "Pieces Return Goal",
    "PIECES_RET_ACTUAL": "Pieces Return Actual",
    "EXPENSES_BUDGET": "Expenses Budget",
    "EXPENSES_ACTUAL": "Expenses Actual",
    "CAMPAN_AMT_GOAL": "Campaign Goal Amount",
    "CAMPAN_AMT_ACTUAL": "Campaign Actual Amount",
    "CAMP_CONTACT_ID_NUM": "Campaign Contact ID",
    "CAMP_START_DATE": "Campaign Start Date",
    "CAMP_END_DATE": "Campaign End Date",
    "ONLINE_GIVING_AVAIL": "Online Giving Available",
    "ONLINE_GIVING_DESC": "Online Giving Description",
}

dim_gift_category_renames = {
    "APPID": "Category ID",
    "CAT_COMP_1": "Category Type",
    "CAT_COMP_2": "Category Code",
    "GIFT_CAT_DESC": "Category Description",
    "CAMPAIGN_CDE": "Campaign Code",
    "FIN_AID_ELEMENT": "Financial Aid Element",
    "MEM_HONOR_CDE": "Memorial Honor Code",
    "FUND_TYPE": "Fund Type",
    "GIVING_CLUB_CDE": "Giving Club Code",
    "DESIGNATION_CDE": "Designation Code",
    "CASE_GROUP": "Case Group",
    "CHARITABLE_FLAG": "Is Charitable",
    "ONLINE_GIVING_AVAIL": "Online Giving Available",
    "ONLINE_GIVING_DESC": "Online Giving Description",
    "PROJECT_CODE": "Project Code",
}

dim_solicitation_renames = {
    "SOLICIT_CDE": "Solicitation Code",
    "DESCRIPTION": "Solicitation Description",
}

dim_year_renames = {
    "YR_CDE": "Fiscal Year Code",
    "FISCAL_YEAR": "Fiscal Year Short",
    "FISCAL_YEAR_FULL": "Fiscal Year",
    "CALENDAR_YEAR": "Calendar Year",
    "YEAR_TYPE": "Year Type",
}

fact_gift_transaction_renames = {
    "GIFT_GROUP_NUM": "Gift Group ID",
    "GIFT_NUM": "Gift ID",
    "GIFT_TRAN_NUM": "Transaction ID",
    "DONOR_ID": "Donor ID",
    "YR_CDE": "Fiscal Year",
    "CAT_COMP_1": "Category Type",
    "CAT_COMP_2": "Category Code",
    "MEM_HONOR_CDE": "Memorial Honor Code",
    "CAMPAIGN_CDE": "Campaign Code",
    "SOLICIT_CDE": "Solicitation Code",
    "GIVING_CLUB_CDE": "Giving Club Code",
    "GIFT_DTE": "Gift Date",
    "GIFT_TRAN_AMT": "Transaction Amount",
    "GIVING_RELATION": "Giving Relationship",
    "CHARITABLE_YN": "Is Charitable",
    "GIFT_TRAN_STS": "Transaction Status",
    "SOFT_CREDIT_YN": "Is Soft Credit",
    "GIFT_CLASS": "Gift Class",
    "SUB_CLASS_CDE": "Gift Subclass",
    "GIFT_SET_CDE": "Gift Set Code",
    "ANON_GIFT_TRAN": "Is Anonymous",
    "NOTATION_1": "Note 1",
    "NOTATION_2": "Note 2",
    "RESTR_TYPE": "Restriction Type",
    "CONTRIB_TYPE": "Contribution Type",
    "MATURITY_AMT": "Maturity Amount",
    "SOLICITOR_ID": "Solicitor ID",
    "GIFT_AMT": "Gift Amount",
    "BANK_ID": "Bank ID",
    "CHECK_CC_NUM": "Check or Card Number",
    "EXPIRE_DTE": "Expiration Date",
    "PRINT_RECPT_YN": "Print Receipt",
    "IMMED_LETTER_YN": "Immediate Letter",
    "BOOK_YN": "Is Booked",
    "GIFT_MASTER_STS": "Gift Status",
}

fact_donor_year_summary_renames = {
    "ID_NUM": "Donor ID",
    "YR_CDE": "Fiscal Year",
    "YEAR_TYPE": "Year Type",
    "CASH_GIFT_NUM": "Cash Gift Count",
    "CASH_GIFT_AMT": "Cash Gift Amount",
    "PROMISE_NUM": "Pledge Count",
    "PROMISE_AMT": "Pledge Amount",
    "PROMISE_PMT_NUM": "Pledge Payment Count",
    "PROMISE_PMT_AMT": "Pledge Payment Amount",
    "MATCH_GIFT_NUM": "Matching Gift Count",
    "MATCH_GIFT_AMT": "Matching Gift Amount",
    "MATCH_PMT_NUM": "Matching Payment Count",
    "MATCH_PMT_AMT": "Matching Payment Amount",
    "SOFT_CREDIT_NUM": "Soft Credit Count",
    "SOFT_CREDIT_AMT": "Soft Credit Amount",
    "NON_CASH_NUM": "Non-Cash Gift Count",
    "NON_CASH_AMT": "Non-Cash Gift Amount",
    "GIFTS_NUM": "Total Gift Count",
    "GIFTS_AMT": "Total Gift Amount",
    "DEFER_NUM": "Deferred Gift Count",
    "DEFER_AMT": "Deferred Gift Amount",
    "DEFER_PMT_NUM": "Deferred Payment Count",
    "DEFER_PMT_AMT": "Deferred Payment Amount",
    "DON_RESTR_NUM": "Donor Restricted Count",
    "DON_RESTR_AMT": "Donor Restricted Amount",
    "ORG_RESTR_NUM": "Org Restricted Count",
    "ORG_RESTR_AMT": "Org Restricted Amount",
    "FIRST_GIFT_DTE": "First Gift Date",
    "FIRST_GIFT_AMT": "First Gift Amount",
    "LAST_GIFT_DTE": "Last Gift Date",
    "LAST_GIFT_AMT": "Last Gift Amount",
    "LARGEST_GIFT_DTE": "Largest Gift Date",
    "LARGEST_GIFT_AMT": "Largest Gift Amount",
    "SMALLEST_GIFT_DTE": "Smallest Gift Date",
    "SMALLEST_GIFT_AMT": "Smallest Gift Amount",
}

fact_donor_campaign_summary_renames = {
    "ID_NUM": "Donor ID",
    "CAMPAIGN_CDE": "Campaign Code",
    "YR_CDE": "Fiscal Year",
    "YEAR_TYPE": "Year Type",
    "CASH_GIFT_NUM": "Cash Gift Count",
    "CASH_GIFT_AMT": "Cash Gift Amount",
    "PROMISE_NUM": "Pledge Count",
    "PROMISE_AMT": "Pledge Amount",
    "PROMISE_PMT_NUM": "Pledge Payment Count",
    "PROMISE_PMT_AMT": "Pledge Payment Amount",
    "MATCH_GIFT_NUM": "Matching Gift Count",
    "MATCH_GIFT_AMT": "Matching Gift Amount",
    "MATCH_PMT_NUM": "Matching Payment Count",
    "MATCH_PMT_AMT": "Matching Payment Amount",
    "SOFT_CREDIT_NUM": "Soft Credit Count",
    "SOFT_CREDIT_AMT": "Soft Credit Amount",
    "NON_CASH_NUM": "Non-Cash Gift Count",
    "NON_CASH_AMT": "Non-Cash Gift Amount",
    "GIFTS_NUM": "Total Gift Count",
    "GIFTS_AMT": "Total Gift Amount",
    "DEFER_NUM": "Deferred Gift Count",
    "DEFER_AMT": "Deferred Gift Amount",
    "DEFER_PMT_NUM": "Deferred Payment Count",
    "DEFER_PMT_AMT": "Deferred Payment Amount",
    "DON_RESTR_NUM": "Donor Restricted Count",
    "DON_RESTR_AMT": "Donor Restricted Amount",
    "ORG_RESTR_NUM": "Org Restricted Count",
    "ORG_RESTR_AMT": "Org Restricted Amount",
    "FIRST_GIFT_DTE": "First Gift Date",
    "FIRST_GIFT_AMT": "First Gift Amount",
    "LAST_GIFT_DTE": "Last Gift Date",
    "LAST_GIFT_AMT": "Last Gift Amount",
    "LARGEST_GIFT_DTE": "Largest Gift Date",
    "LARGEST_GIFT_AMT": "Largest Gift Amount",
    "SMALLEST_GIFT_DTE": "Smallest Gift Date",
    "SMALLEST_GIFT_AMT": "Smallest Gift Amount",
}

# ============================================================
# PROCESS EACH FILE
# ============================================================

files_config = [
    ("DimConstituent.xlsx", "Constituents.xlsx", dim_constituent_renames),
    ("DimAlumni.xlsx", "Alumni.xlsx", dim_alumni_renames),
    ("DimCampaign.xlsx", "Campaigns.xlsx", dim_campaign_renames),
    ("DimGiftCategory.xlsx", "Gift Categories.xlsx", dim_gift_category_renames),
    ("DimSolicitation.xlsx", "Solicitations.xlsx", dim_solicitation_renames),
    ("DimYear.xlsx", "Fiscal Years.xlsx", dim_year_renames),
    ("FactGiftTransaction.xlsx", "Gift Transactions.xlsx", fact_gift_transaction_renames),
    ("FactDonorYearSummary.xlsx", "Donor Year Summary.xlsx", fact_donor_year_summary_renames),
    ("FactDonorCampaignSummary.xlsx", "Donor Campaign Summary.xlsx", fact_donor_campaign_summary_renames),
]

if __name__ == "__main__":
    print("Renaming columns and saving files...")
    print("=" * 60)

    for src_file, dest_file, rename_map in files_config:
        src_path = cleaned_dir / src_file
        dest_path = renamed_dir / dest_file

        if src_path.exists():
            df = pd.read_excel(src_path)

            # Count columns that will be renamed
            cols_renamed = sum(1 for old_col in rename_map if old_col in df.columns)

            df = df.rename(columns=rename_map)

            # Save
            df.to_excel(dest_path, index=False, engine="openpyxl")
            print(f"\n{src_file} -> {dest_file}")
            print(f"  Rows: {len(df):,}")
            print(f"  Columns renamed: {cols_renamed}")
        else:
            print(f"WARNING: {src_file} not found")

    print("\n" + "=" * 60)
    print("All files saved to 'Renamed Excel Model Files' folder")
