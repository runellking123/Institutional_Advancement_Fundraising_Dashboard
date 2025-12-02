// ============================================================================
// Tabular Editor 2 - Bulk Create Measures Script
// Institutional Advancement Fundraising Dashboard
// ============================================================================
// Instructions:
// 1. Open your Power BI model in Tabular Editor 2
// 2. Go to File > Open > From DB or open your .bim file
// 3. Go to Advanced Scripting tab (C# Script)
// 4. Paste this entire script and click Run (F5)
// 5. Save your model when complete
// ============================================================================

using System.Collections.Generic;

// Counter for created measures
int measureCount = 0;
List<string> createdMeasures = new List<string>();
List<string> errorMessages = new List<string>();

// Helper function to safely create a measure
Action<string, string, string, string, string> CreateMeasure = (tableName, measureName, expression, formatString, displayFolder) =>
{
    try
    {
        var table = Model.Tables[tableName];
        if (table == null)
        {
            errorMessages.Add("Table not found: " + tableName + " (for measure: " + measureName + ")");
            return;
        }

        // Check if measure already exists
        if (table.Measures.Contains(measureName))
        {
            // Update existing measure
            table.Measures[measureName].Expression = expression;
            if (!string.IsNullOrEmpty(formatString))
                table.Measures[measureName].FormatString = formatString;
            if (!string.IsNullOrEmpty(displayFolder))
                table.Measures[measureName].DisplayFolder = displayFolder;
            createdMeasures.Add(measureName + " (updated)");
        }
        else
        {
            // Create new measure
            var measure = table.AddMeasure(measureName, expression);
            if (!string.IsNullOrEmpty(formatString))
                measure.FormatString = formatString;
            if (!string.IsNullOrEmpty(displayFolder))
                measure.DisplayFolder = displayFolder;
            createdMeasures.Add(measureName);
        }
        measureCount++;
    }
    catch (Exception ex)
    {
        errorMessages.Add("Error creating " + measureName + ": " + ex.Message);
    }
};

// ============================================================================
// BASIC GIVING METRICS
// ============================================================================

CreateMeasure("GIFT_TRAN", "Total Gifts",
@"SUM(GIFT_TRAN[GIFT_TRAN_AMT])",
"$#,##0.00", "01. Core Metrics");

CreateMeasure("GIFT_TRAN", "Total Gifts (Excluding Voided)",
@"CALCULATE(
    SUM(GIFT_TRAN[GIFT_TRAN_AMT]),
    GIFT_TRAN[GIFT_TRAN_STS] <> ""V""
)",
"$#,##0.00", "01. Core Metrics");

CreateMeasure("GIFT_TRAN", "Gift Count",
@"COUNTROWS(GIFT_TRAN)",
"#,##0", "01. Core Metrics");

CreateMeasure("GIFT_TRAN", "Gift Count (Excluding Voided)",
@"CALCULATE(
    COUNTROWS(GIFT_TRAN),
    GIFT_TRAN[GIFT_TRAN_STS] <> ""V""
)",
"#,##0", "01. Core Metrics");

CreateMeasure("GIFT_TRAN", "Average Gift Size",
@"DIVIDE(
    [Total Gifts (Excluding Voided)],
    [Gift Count (Excluding Voided)],
    0
)",
"$#,##0.00", "01. Core Metrics");

CreateMeasure("GIFT_TRAN", "Median Gift",
@"MEDIAN(GIFT_TRAN[GIFT_TRAN_AMT])",
"$#,##0.00", "01. Core Metrics");

CreateMeasure("GIFT_TRAN", "Largest Gift",
@"MAX(GIFT_TRAN[GIFT_TRAN_AMT])",
"$#,##0.00", "01. Core Metrics");

CreateMeasure("GIFT_TRAN", "Smallest Gift",
@"MIN(GIFT_TRAN[GIFT_TRAN_AMT])",
"$#,##0.00", "01. Core Metrics");

// ============================================================================
// DONOR METRICS
// ============================================================================

CreateMeasure("GIFT_TRAN", "Total Donors",
@"DISTINCTCOUNT(GIFT_TRAN[DONOR_ID])",
"#,##0", "02. Donor Metrics");

CreateMeasure("GIFT_TRAN", "Total Donors (Excluding Voided)",
@"CALCULATE(
    DISTINCTCOUNT(GIFT_TRAN[DONOR_ID]),
    GIFT_TRAN[GIFT_TRAN_STS] <> ""V""
)",
"#,##0", "02. Donor Metrics");

CreateMeasure("GIFT_TRAN", "Average Gift Per Donor",
@"DIVIDE(
    [Total Gifts (Excluding Voided)],
    [Total Donors (Excluding Voided)],
    0
)",
"$#,##0.00", "02. Donor Metrics");

CreateMeasure("GIFT_TRAN", "Gifts Per Donor",
@"DIVIDE(
    [Gift Count (Excluding Voided)],
    [Total Donors (Excluding Voided)],
    0
)",
"#,##0.00", "02. Donor Metrics");

CreateMeasure("GIFT_TRAN", "New Donors (Current Year)",
@"VAR CurrentYear = YEAR(TODAY())
RETURN
CALCULATE(
    DISTINCTCOUNT(GIFT_TRAN[DONOR_ID]),
    GIFT_TRAN[GIFT_TRAN_STS] <> ""V"",
    FILTER(
        ALL(DONOR_MASTER),
        YEAR(DONOR_MASTER[FIRST_GIFT_DTE]) = CurrentYear
    )
)",
"#,##0", "02. Donor Metrics");

CreateMeasure("GIFT_TRAN", "Current Year Donors",
@"VAR CurrentYear = YEAR(TODAY())
RETURN
CALCULATE(
    DISTINCTCOUNT(GIFT_TRAN[DONOR_ID]),
    GIFT_TRAN[GIFT_TRAN_STS] <> ""V"",
    YEAR(GIFT_TRAN[GIFT_DTE]) = CurrentYear
)",
"#,##0", "02. Donor Metrics");

CreateMeasure("GIFT_TRAN", "Prior Year Donors",
@"VAR PriorYear = YEAR(TODAY()) - 1
RETURN
CALCULATE(
    DISTINCTCOUNT(GIFT_TRAN[DONOR_ID]),
    GIFT_TRAN[GIFT_TRAN_STS] <> ""V"",
    YEAR(GIFT_TRAN[GIFT_DTE]) = PriorYear
)",
"#,##0", "02. Donor Metrics");

CreateMeasure("GIFT_TRAN", "Retained Donors",
@"VAR CurrentYear = YEAR(TODAY())
VAR PriorYear = CurrentYear - 1
VAR CurrentYearDonors =
    CALCULATETABLE(
        VALUES(GIFT_TRAN[DONOR_ID]),
        GIFT_TRAN[GIFT_TRAN_STS] <> ""V"",
        YEAR(GIFT_TRAN[GIFT_DTE]) = CurrentYear
    )
VAR PriorYearDonors =
    CALCULATETABLE(
        VALUES(GIFT_TRAN[DONOR_ID]),
        GIFT_TRAN[GIFT_TRAN_STS] <> ""V"",
        YEAR(GIFT_TRAN[GIFT_DTE]) = PriorYear
    )
RETURN
    COUNTROWS(INTERSECT(CurrentYearDonors, PriorYearDonors))",
"#,##0", "02. Donor Metrics");

CreateMeasure("GIFT_TRAN", "Lapsed Donors",
@"VAR CurrentYear = YEAR(TODAY())
VAR PriorYear = CurrentYear - 1
VAR CurrentYearDonors =
    CALCULATETABLE(
        VALUES(GIFT_TRAN[DONOR_ID]),
        GIFT_TRAN[GIFT_TRAN_STS] <> ""V"",
        YEAR(GIFT_TRAN[GIFT_DTE]) = CurrentYear
    )
VAR PriorYearDonors =
    CALCULATETABLE(
        VALUES(GIFT_TRAN[DONOR_ID]),
        GIFT_TRAN[GIFT_TRAN_STS] <> ""V"",
        YEAR(GIFT_TRAN[GIFT_DTE]) = PriorYear
    )
RETURN
    COUNTROWS(EXCEPT(PriorYearDonors, CurrentYearDonors))",
"#,##0", "02. Donor Metrics");

CreateMeasure("GIFT_TRAN", "Donor Retention Rate",
@"DIVIDE(
    [Retained Donors],
    [Prior Year Donors],
    0
)",
"0.0%", "02. Donor Metrics");

CreateMeasure("GIFT_TRAN", "Donor Acquisition Rate",
@"DIVIDE(
    [New Donors (Current Year)],
    [Current Year Donors],
    0
)",
"0.0%", "02. Donor Metrics");

// ============================================================================
// TIME INTELLIGENCE MEASURES
// ============================================================================

CreateMeasure("GIFT_TRAN", "YTD Gifts",
@"VAR CurrentDate = TODAY()
VAR StartOfYear = DATE(YEAR(CurrentDate), 1, 1)
RETURN
CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_DTE] >= StartOfYear,
    GIFT_TRAN[GIFT_DTE] <= CurrentDate
)",
"$#,##0.00", "03. Time Intelligence");

CreateMeasure("GIFT_TRAN", "Prior Year Total Gifts",
@"VAR PriorYear = YEAR(TODAY()) - 1
RETURN
CALCULATE(
    [Total Gifts (Excluding Voided)],
    YEAR(GIFT_TRAN[GIFT_DTE]) = PriorYear
)",
"$#,##0.00", "03. Time Intelligence");

CreateMeasure("GIFT_TRAN", "Prior YTD Gifts",
@"VAR CurrentDate = TODAY()
VAR PriorYearSameDate = DATE(YEAR(CurrentDate) - 1, MONTH(CurrentDate), DAY(CurrentDate))
VAR PriorYearStart = DATE(YEAR(CurrentDate) - 1, 1, 1)
RETURN
CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_DTE] >= PriorYearStart,
    GIFT_TRAN[GIFT_DTE] <= PriorYearSameDate
)",
"$#,##0.00", "03. Time Intelligence");

CreateMeasure("GIFT_TRAN", "YoY Growth Amount",
@"[YTD Gifts] - [Prior YTD Gifts]",
"$#,##0.00", "03. Time Intelligence");

CreateMeasure("GIFT_TRAN", "YoY Growth %",
@"DIVIDE(
    [YTD Gifts] - [Prior YTD Gifts],
    [Prior YTD Gifts],
    0
)",
"0.0%", "03. Time Intelligence");

CreateMeasure("GIFT_TRAN", "Current Month Gifts",
@"VAR CurrentDate = TODAY()
VAR StartOfMonth = DATE(YEAR(CurrentDate), MONTH(CurrentDate), 1)
RETURN
CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_DTE] >= StartOfMonth,
    GIFT_TRAN[GIFT_DTE] <= CurrentDate
)",
"$#,##0.00", "03. Time Intelligence");

CreateMeasure("GIFT_TRAN", "Prior Month Gifts",
@"VAR CurrentDate = TODAY()
VAR StartOfPriorMonth = EOMONTH(CurrentDate, -2) + 1
VAR EndOfPriorMonth = EOMONTH(CurrentDate, -1)
RETURN
CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_DTE] >= StartOfPriorMonth,
    GIFT_TRAN[GIFT_DTE] <= EndOfPriorMonth
)",
"$#,##0.00", "03. Time Intelligence");

CreateMeasure("GIFT_TRAN", "MoM Growth %",
@"DIVIDE(
    [Current Month Gifts] - [Prior Month Gifts],
    [Prior Month Gifts],
    0
)",
"0.0%", "03. Time Intelligence");

CreateMeasure("GIFT_TRAN", "Fiscal YTD Gifts",
@"VAR CurrentDate = TODAY()
VAR CurrentFiscalYearStart =
    IF(
        MONTH(CurrentDate) >= 7,
        DATE(YEAR(CurrentDate), 7, 1),
        DATE(YEAR(CurrentDate) - 1, 7, 1)
    )
RETURN
CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_DTE] >= CurrentFiscalYearStart,
    GIFT_TRAN[GIFT_DTE] <= CurrentDate
)",
"$#,##0.00", "03. Time Intelligence");

CreateMeasure("GIFT_TRAN", "Prior Fiscal Year Gifts",
@"VAR CurrentDate = TODAY()
VAR PriorFiscalYearStart =
    IF(
        MONTH(CurrentDate) >= 7,
        DATE(YEAR(CurrentDate) - 1, 7, 1),
        DATE(YEAR(CurrentDate) - 2, 7, 1)
    )
VAR PriorFiscalYearEnd =
    IF(
        MONTH(CurrentDate) >= 7,
        DATE(YEAR(CurrentDate), 6, 30),
        DATE(YEAR(CurrentDate) - 1, 6, 30)
    )
RETURN
CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_DTE] >= PriorFiscalYearStart,
    GIFT_TRAN[GIFT_DTE] <= PriorFiscalYearEnd
)",
"$#,##0.00", "03. Time Intelligence");

CreateMeasure("GIFT_TRAN", "Current Quarter Gifts",
@"VAR CurrentDate = TODAY()
VAR CurrentQuarter = QUARTER(CurrentDate)
VAR CurrentYear = YEAR(CurrentDate)
VAR QuarterStart = DATE(CurrentYear, (CurrentQuarter - 1) * 3 + 1, 1)
RETURN
CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_DTE] >= QuarterStart,
    GIFT_TRAN[GIFT_DTE] <= CurrentDate
)",
"$#,##0.00", "03. Time Intelligence");

// ============================================================================
// GIFT TYPE METRICS
// ============================================================================

CreateMeasure("GIFT_TRAN", "Cash Gifts Total",
@"CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_CLASS] = ""C""
)",
"$#,##0.00", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Cash Gifts Count",
@"CALCULATE(
    [Gift Count (Excluding Voided)],
    GIFT_TRAN[GIFT_CLASS] = ""C""
)",
"#,##0", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Pledge Total",
@"CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_CLASS] = ""P""
)",
"$#,##0.00", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Pledge Count",
@"CALCULATE(
    [Gift Count (Excluding Voided)],
    GIFT_TRAN[GIFT_CLASS] = ""P""
)",
"#,##0", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Matching Gifts Total",
@"CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_CLASS] = ""M""
)",
"$#,##0.00", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Matching Gifts Count",
@"CALCULATE(
    [Gift Count (Excluding Voided)],
    GIFT_TRAN[GIFT_CLASS] = ""M""
)",
"#,##0", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Non-Cash Gifts Total",
@"CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_CLASS] = ""N""
)",
"$#,##0.00", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Non-Cash Gifts Count",
@"CALCULATE(
    [Gift Count (Excluding Voided)],
    GIFT_TRAN[GIFT_CLASS] = ""N""
)",
"#,##0", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Soft Credit Total",
@"CALCULATE(
    SUM(GIFT_TRAN[GIFT_TRAN_AMT]),
    GIFT_TRAN[GIFT_TRAN_STS] <> ""V"",
    GIFT_TRAN[SOFT_CREDIT_YN] = ""Y""
)",
"$#,##0.00", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Hard Credit Total",
@"CALCULATE(
    SUM(GIFT_TRAN[GIFT_TRAN_AMT]),
    GIFT_TRAN[GIFT_TRAN_STS] <> ""V"",
    GIFT_TRAN[SOFT_CREDIT_YN] = ""N""
)",
"$#,##0.00", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Cash Gift %",
@"DIVIDE(
    [Cash Gifts Total],
    [Total Gifts (Excluding Voided)],
    0
)",
"0.0%", "04. Gift Types");

CreateMeasure("GIFT_TRAN", "Pledge %",
@"DIVIDE(
    [Pledge Total],
    [Total Gifts (Excluding Voided)],
    0
)",
"0.0%", "04. Gift Types");

// ============================================================================
// CAMPAIGN METRICS
// ============================================================================

CreateMeasure("CAMPAIGN", "Campaign Goal Total",
@"SUM(CAMPAIGN[CAMPAN_AMT_GOAL])",
"$#,##0.00", "05. Campaign Metrics");

CreateMeasure("CAMPAIGN", "Campaign Actual Total",
@"SUM(CAMPAIGN[CAMPAN_AMT_ACTUAL])",
"$#,##0.00", "05. Campaign Metrics");

CreateMeasure("CAMPAIGN", "Campaign Goal Achievement %",
@"DIVIDE(
    [Campaign Actual Total],
    [Campaign Goal Total],
    0
)",
"0.0%", "05. Campaign Metrics");

CreateMeasure("CAMPAIGN", "Campaign Count",
@"COUNTROWS(CAMPAIGN)",
"#,##0", "05. Campaign Metrics");

CreateMeasure("CAMPAIGN", "Campaign Gap to Goal",
@"[Campaign Goal Total] - [Campaign Actual Total]",
"$#,##0.00", "05. Campaign Metrics");

CreateMeasure("CAMPAIGN", "Average Campaign Goal",
@"AVERAGE(CAMPAIGN[CAMPAN_AMT_GOAL])",
"$#,##0.00", "05. Campaign Metrics");

CreateMeasure("CAMPAIGN", "Average Campaign Raised",
@"AVERAGE(CAMPAIGN[CAMPAN_AMT_ACTUAL])",
"$#,##0.00", "05. Campaign Metrics");

CreateMeasure("CAMPAIGN", "Campaigns at Goal",
@"COUNTROWS(
    FILTER(
        CAMPAIGN,
        CAMPAIGN[CAMPAN_AMT_ACTUAL] >= CAMPAIGN[CAMPAN_AMT_GOAL] &&
        CAMPAIGN[CAMPAN_AMT_GOAL] > 0
    )
)",
"#,##0", "05. Campaign Metrics");

CreateMeasure("CAMPAIGN", "Campaigns Below Goal",
@"COUNTROWS(
    FILTER(
        CAMPAIGN,
        CAMPAIGN[CAMPAN_AMT_ACTUAL] < CAMPAIGN[CAMPAN_AMT_GOAL] &&
        CAMPAIGN[CAMPAN_AMT_GOAL] > 0
    )
)",
"#,##0", "05. Campaign Metrics");

// ============================================================================
// FUND/CATEGORY METRICS
// ============================================================================

CreateMeasure("GIFT_CATEGORY", "Fund Goal Total",
@"SUM(GIFT_CATEGORY[FUND_GOAL_AMT])",
"$#,##0.00", "06. Fund Metrics");

CreateMeasure("GIFT_CATEGORY", "Annual Goal Total",
@"SUM(GIFT_CATEGORY[ANNUAL_GOAL_AMT])",
"$#,##0.00", "06. Fund Metrics");

CreateMeasure("GIFT_CATEGORY", "Fund Count",
@"COUNTROWS(GIFT_CATEGORY)",
"#,##0", "06. Fund Metrics");

CreateMeasure("GIFT_CATEGORY", "Category Gifts Total",
@"SUM(GIFT_CATEGORY[GIFTS_AMT])",
"$#,##0.00", "06. Fund Metrics");

CreateMeasure("GIFT_CATEGORY", "Category Cash Gifts",
@"SUM(GIFT_CATEGORY[CASH_GIFT_AMT])",
"$#,##0.00", "06. Fund Metrics");

// ============================================================================
// DONOR MASTER LIFETIME METRICS
// ============================================================================

CreateMeasure("DONOR_MASTER", "Total Lifetime Giving",
@"SUM(DONOR_MASTER[GIFTS_AMT])",
"$#,##0.00", "07. Donor Lifetime Metrics");

CreateMeasure("DONOR_MASTER", "Average Donor Lifetime Value",
@"AVERAGE(DONOR_MASTER[GIFTS_AMT])",
"$#,##0.00", "07. Donor Lifetime Metrics");

CreateMeasure("DONOR_MASTER", "Total Lifetime Gift Count",
@"SUM(DONOR_MASTER[GIFTS_NUM])",
"#,##0", "07. Donor Lifetime Metrics");

CreateMeasure("DONOR_MASTER", "Average Donor Gift Count",
@"AVERAGE(DONOR_MASTER[GIFTS_NUM])",
"#,##0.0", "07. Donor Lifetime Metrics");

CreateMeasure("DONOR_MASTER", "Average Years Continuous Giving",
@"AVERAGE(DONOR_MASTER[YRS_CONT_GIVING])",
"#,##0.0", "07. Donor Lifetime Metrics");

CreateMeasure("DONOR_MASTER", "Max Years Continuous Giving",
@"MAX(DONOR_MASTER[YRS_CONT_GIVING])",
"#,##0", "07. Donor Lifetime Metrics");

CreateMeasure("DONOR_MASTER", "Active Donors",
@"COUNTROWS(
    FILTER(
        DONOR_MASTER,
        DONOR_MASTER[ACTIVE_FLAG] = ""Y""
    )
)",
"#,##0", "07. Donor Lifetime Metrics");

CreateMeasure("DONOR_MASTER", "Inactive Donors",
@"COUNTROWS(
    FILTER(
        DONOR_MASTER,
        DONOR_MASTER[ACTIVE_FLAG] <> ""Y""
    )
)",
"#,##0", "07. Donor Lifetime Metrics");

CreateMeasure("DONOR_MASTER", "Anonymous Donors",
@"COUNTROWS(
    FILTER(
        DONOR_MASTER,
        DONOR_MASTER[ANON_DONOR_FLAG] = ""Y""
    )
)",
"#,##0", "07. Donor Lifetime Metrics");

CreateMeasure("DONOR_MASTER", "Donors with Pledges",
@"COUNTROWS(
    FILTER(
        DONOR_MASTER,
        DONOR_MASTER[PROMISE_NUM] > 0
    )
)",
"#,##0", "07. Donor Lifetime Metrics");

CreateMeasure("DONOR_MASTER", "Total Outstanding Pledges",
@"SUM(DONOR_MASTER[PROMISE_AMT]) - SUM(DONOR_MASTER[PROMISE_PMT_AMT])",
"$#,##0.00", "07. Donor Lifetime Metrics");

// ============================================================================
// GIVING LEVEL SEGMENTATION
// ============================================================================

CreateMeasure("GIFT_TRAN", "Major Gifts ($10K+)",
@"CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_TRAN_AMT] >= 10000
)",
"$#,##0.00", "08. Giving Levels");

CreateMeasure("GIFT_TRAN", "Major Gift Count ($10K+)",
@"CALCULATE(
    [Gift Count (Excluding Voided)],
    GIFT_TRAN[GIFT_TRAN_AMT] >= 10000
)",
"#,##0", "08. Giving Levels");

CreateMeasure("GIFT_TRAN", "Major Gift Donors ($10K+)",
@"CALCULATE(
    DISTINCTCOUNT(GIFT_TRAN[DONOR_ID]),
    GIFT_TRAN[GIFT_TRAN_STS] <> ""V"",
    GIFT_TRAN[GIFT_TRAN_AMT] >= 10000
)",
"#,##0", "08. Giving Levels");

CreateMeasure("GIFT_TRAN", "Mid-Level Gifts ($1K-$10K)",
@"CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_TRAN_AMT] >= 1000,
    GIFT_TRAN[GIFT_TRAN_AMT] < 10000
)",
"$#,##0.00", "08. Giving Levels");

CreateMeasure("GIFT_TRAN", "Mid-Level Gift Count ($1K-$10K)",
@"CALCULATE(
    [Gift Count (Excluding Voided)],
    GIFT_TRAN[GIFT_TRAN_AMT] >= 1000,
    GIFT_TRAN[GIFT_TRAN_AMT] < 10000
)",
"#,##0", "08. Giving Levels");

CreateMeasure("GIFT_TRAN", "Annual Fund Gifts (<$1K)",
@"CALCULATE(
    [Total Gifts (Excluding Voided)],
    GIFT_TRAN[GIFT_TRAN_AMT] < 1000
)",
"$#,##0.00", "08. Giving Levels");

CreateMeasure("GIFT_TRAN", "Annual Fund Gift Count (<$1K)",
@"CALCULATE(
    [Gift Count (Excluding Voided)],
    GIFT_TRAN[GIFT_TRAN_AMT] < 1000
)",
"#,##0", "08. Giving Levels");

CreateMeasure("GIFT_TRAN", "Major Gift % of Total",
@"DIVIDE(
    [Major Gifts ($10K+)],
    [Total Gifts (Excluding Voided)],
    0
)",
"0.0%", "08. Giving Levels");

// ============================================================================
// YEAR SUMMARY METRICS
// ============================================================================

CreateMeasure("DONOR_YEAR_SUM", "Year Summary Total Gifts",
@"SUM(DONOR_YEAR_SUM[GIFTS_AMT])",
"$#,##0.00", "09. Year Summary");

CreateMeasure("DONOR_YEAR_SUM", "Year Summary Gift Count",
@"SUM(DONOR_YEAR_SUM[GIFTS_NUM])",
"#,##0", "09. Year Summary");

CreateMeasure("DONOR_YEAR_SUM", "Year Summary Cash Gifts",
@"SUM(DONOR_YEAR_SUM[CASH_GIFT_AMT])",
"$#,##0.00", "09. Year Summary");

CreateMeasure("DONOR_YEAR_SUM", "Year Summary Donors",
@"DISTINCTCOUNT(DONOR_YEAR_SUM[ID_NUM])",
"#,##0", "09. Year Summary");

// ============================================================================
// CAMPAIGN SUMMARY METRICS
// ============================================================================

CreateMeasure("DONOR_CAMP_SUM", "Campaign Summary Total Gifts",
@"SUM(DONOR_CAMP_SUM[GIFTS_AMT])",
"$#,##0.00", "10. Campaign Summary");

CreateMeasure("DONOR_CAMP_SUM", "Campaign Summary Gift Count",
@"SUM(DONOR_CAMP_SUM[GIFTS_NUM])",
"#,##0", "10. Campaign Summary");

CreateMeasure("DONOR_CAMP_SUM", "Campaign Summary Donors",
@"DISTINCTCOUNT(DONOR_CAMP_SUM[ID_NUM])",
"#,##0", "10. Campaign Summary");

// ============================================================================
// KPI MEASURES FOR CARDS
// ============================================================================

CreateMeasure("GIFT_TRAN", "KPI - Total Raised",
@"[Total Gifts (Excluding Voided)]",
"$#,##0", "11. KPI Cards");

CreateMeasure("GIFT_TRAN", "KPI - Total Donors",
@"[Total Donors (Excluding Voided)]",
"#,##0", "11. KPI Cards");

CreateMeasure("GIFT_TRAN", "KPI - Avg Gift",
@"[Average Gift Size]",
"$#,##0", "11. KPI Cards");

CreateMeasure("GIFT_TRAN", "KPI - YoY Growth",
@"[YoY Growth %]",
"0.0%", "11. KPI Cards");

CreateMeasure("GIFT_TRAN", "KPI - Retention Rate",
@"[Donor Retention Rate]",
"0.0%", "11. KPI Cards");

CreateMeasure("CAMPAIGN", "KPI - Campaign Progress",
@"[Campaign Goal Achievement %]",
"0.0%", "11. KPI Cards");

// ============================================================================
// DISPLAY COMPLETION MESSAGE
// ============================================================================

string resultMessage = "===== MEASURE CREATION COMPLETE =====\n\n";
resultMessage += "Total Measures Created/Updated: " + measureCount + "\n\n";

if (createdMeasures.Count > 0)
{
    resultMessage += "MEASURES:\n";
    resultMessage += "─────────────────────────────────\n";
    foreach (var m in createdMeasures)
    {
        resultMessage += "✓ " + m + "\n";
    }
}

if (errorMessages.Count > 0)
{
    resultMessage += "\n\nERRORS:\n";
    resultMessage += "─────────────────────────────────\n";
    foreach (var e in errorMessages)
    {
        resultMessage += "✗ " + e + "\n";
    }
}

resultMessage += "\n─────────────────────────────────\n";
resultMessage += "Remember to SAVE your model!\n";
resultMessage += "(File > Save or Ctrl+S)";

Info(resultMessage);
