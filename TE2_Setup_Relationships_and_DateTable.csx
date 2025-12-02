// ============================================================================
// Tabular Editor 2 - Setup Relationships and Date Table Script
// Institutional Advancement Fundraising Dashboard
// VERSION 2 - Fixed Ambiguous Path Issues
// ============================================================================
// Instructions:
// 1. Open your Power BI model in Tabular Editor 2
// 2. Go to File > Open > From DB or open your .bim file
// 3. Go to Advanced Scripting tab (C# Script)
// 4. Paste this entire script and click Run (F5)
// 5. Save your model when complete
// ============================================================================

using System.Collections.Generic;

int itemsCreated = 0;
List<string> successLog = new List<string>();
List<string> errorLog = new List<string>();

// ============================================================================
// STEP 1: CREATE DATE TABLE (DimDate)
// ============================================================================

try
{
    // Check if DimDate already exists
    if (Model.Tables.Contains("DimDate"))
    {
        successLog.Add("DimDate table already exists - skipping creation");
    }
    else
    {
        // Create the calculated table with proper DAX syntax
        var dimDateExpression = @"
VAR StartDate = DATE(2015, 1, 1)
VAR EndDate = DATE(2030, 12, 31)
RETURN
ADDCOLUMNS(
    CALENDAR(StartDate, EndDate),
    ""Year"", YEAR([Date]),
    ""YearText"", FORMAT([Date], ""YYYY""),
    ""Month"", FORMAT([Date], ""MMMM""),
    ""MonthNum"", MONTH([Date]),
    ""MonthYear"", FORMAT([Date], ""MMM YYYY""),
    ""Quarter"", ""Q"" & FORMAT(QUARTER([Date]), ""0""),
    ""QuarterNum"", QUARTER([Date]),
    ""YearQuarter"", FORMAT([Date], ""YYYY"") & ""-Q"" & FORMAT(QUARTER([Date]), ""0""),
    ""YearMonth"", FORMAT([Date], ""YYYY-MM""),
    ""YearMonthNum"", YEAR([Date]) * 100 + MONTH([Date]),
    ""DayOfWeek"", WEEKDAY([Date]),
    ""DayName"", FORMAT([Date], ""dddd""),
    ""DayOfMonth"", DAY([Date]),
    ""WeekNum"", WEEKNUM([Date]),
    ""IsWeekend"", IF(WEEKDAY([Date]) IN {1, 7}, ""Weekend"", ""Weekday""),
    ""FiscalYear"", IF(MONTH([Date]) >= 7, YEAR([Date]) + 1, YEAR([Date])),
    ""FiscalYearText"", ""FY"" & FORMAT(IF(MONTH([Date]) >= 7, YEAR([Date]) + 1, YEAR([Date])), ""0""),
    ""FiscalQuarterNum"",
        SWITCH(
            TRUE(),
            MONTH([Date]) IN {7, 8, 9}, 1,
            MONTH([Date]) IN {10, 11, 12}, 2,
            MONTH([Date]) IN {1, 2, 3}, 3,
            4
        ),
    ""FiscalQuarter"",
        ""FQ"" & FORMAT(
            SWITCH(
                TRUE(),
                MONTH([Date]) IN {7, 8, 9}, 1,
                MONTH([Date]) IN {10, 11, 12}, 2,
                MONTH([Date]) IN {1, 2, 3}, 3,
                4
            ), ""0""),
    ""FiscalMonth"",
        SWITCH(
            TRUE(),
            MONTH([Date]) >= 7, MONTH([Date]) - 6,
            MONTH([Date]) + 6
        ),
    ""IsCurrentYear"", IF(YEAR([Date]) = YEAR(TODAY()), ""Current Year"", ""Prior Years""),
    ""IsCurrentMonth"", IF(YEAR([Date]) = YEAR(TODAY()) && MONTH([Date]) = MONTH(TODAY()), ""Current Month"", ""Other""),
    ""IsYTD"", IF([Date] <= TODAY() && YEAR([Date]) = YEAR(TODAY()), ""YTD"", ""Not YTD"")
)";

        var dimDate = Model.AddCalculatedTable("DimDate", dimDateExpression);

        // Mark as Date table
        dimDate.DataCategory = "Time";

        // Set the Date column as the key
        if (dimDate.Columns.Contains("Date"))
        {
            dimDate.Columns["Date"].IsKey = true;
            dimDate.Columns["Date"].DataType = DataType.DateTime;
            dimDate.Columns["Date"].FormatString = "yyyy-MM-dd";
        }

        // Format other columns
        foreach (var col in dimDate.Columns)
        {
            if (col.Name.Contains("Num") || col.Name == "Year" || col.Name == "FiscalYear" || col.Name == "DayOfMonth" || col.Name == "WeekNum" || col.Name == "DayOfWeek")
            {
                col.DataType = DataType.Int64;
            }
        }

        successLog.Add("Created DimDate calculated table with 22 columns");
        itemsCreated++;
    }
}
catch (Exception ex)
{
    errorLog.Add("Error creating DimDate: " + ex.Message);
}

// ============================================================================
// STEP 2: CREATE RELATIONSHIPS
// ============================================================================

// Helper function to create relationships safely
Action<string, string, string, string, string, bool, string> CreateRelationship =
    (fromTable, fromColumn, toTable, toColumn, crossFilter, isActive, description) =>
{
    try
    {
        // Check if tables exist
        if (!Model.Tables.Contains(fromTable))
        {
            errorLog.Add("Table not found: " + fromTable);
            return;
        }
        if (!Model.Tables.Contains(toTable))
        {
            errorLog.Add("Table not found: " + toTable);
            return;
        }

        var from = Model.Tables[fromTable];
        var to = Model.Tables[toTable];

        // Check if columns exist
        if (!from.Columns.Contains(fromColumn))
        {
            errorLog.Add("Column not found: " + fromTable + "[" + fromColumn + "]");
            return;
        }
        if (!to.Columns.Contains(toColumn))
        {
            errorLog.Add("Column not found: " + toTable + "[" + toColumn + "]");
            return;
        }

        // Check if relationship already exists
        foreach (var rel in Model.Relationships)
        {
            if (rel.FromTable.Name == fromTable &&
                rel.FromColumn.Name == fromColumn &&
                rel.ToTable.Name == toTable &&
                rel.ToColumn.Name == toColumn)
            {
                successLog.Add("Relationship already exists: " + description);
                return;
            }
        }

        // Create the relationship
        var relationship = Model.AddRelationship();
        relationship.FromColumn = from.Columns[fromColumn];
        relationship.ToColumn = to.Columns[toColumn];

        // Set cross-filter direction
        if (crossFilter == "Both")
        {
            relationship.CrossFilteringBehavior = CrossFilteringBehavior.BothDirections;
        }
        else
        {
            relationship.CrossFilteringBehavior = CrossFilteringBehavior.OneDirection;
        }

        // Set active status
        relationship.IsActive = isActive;

        string activeStatus = isActive ? "ACTIVE" : "INACTIVE";
        successLog.Add("Created relationship (" + activeStatus + "): " + description);
        itemsCreated++;
    }
    catch (Exception ex)
    {
        errorLog.Add("Error creating relationship (" + description + "): " + ex.Message);
    }
};

// ============================================================================
// PRIMARY RELATIONSHIPS (ACTIVE) - These form the main star schema
// ============================================================================

// 1. GIFT_TRAN → DONOR_MASTER (Core relationship - ACTIVE)
CreateRelationship(
    "GIFT_TRAN", "DONOR_ID",
    "DONOR_MASTER", "ID_NUM",
    "Both", true,
    "GIFT_TRAN[DONOR_ID] → DONOR_MASTER[ID_NUM]"
);

// 2. DONOR_MASTER → NAME_MASTER (Donor names - ACTIVE)
CreateRelationship(
    "DONOR_MASTER", "ID_NUM",
    "NAME_MASTER", "ID_NUM",
    "Both", true,
    "DONOR_MASTER[ID_NUM] → NAME_MASTER[ID_NUM]"
);

// 3. GIFT_TRAN → CAMPAIGN (Campaign dimension - ACTIVE)
CreateRelationship(
    "GIFT_TRAN", "CAMPAIGN_CDE",
    "CAMPAIGN", "CAMPAIGN_CDE",
    "Single", true,
    "GIFT_TRAN[CAMPAIGN_CDE] → CAMPAIGN[CAMPAIGN_CDE]"
);

// 4. GIFT_TRAN → GIFT_CATEGORY (Fund/Designation dimension - ACTIVE)
CreateRelationship(
    "GIFT_TRAN", "CAT_COMP_1",
    "GIFT_CATEGORY", "CAT_COMP_1",
    "Single", true,
    "GIFT_TRAN[CAT_COMP_1] → GIFT_CATEGORY[CAT_COMP_1]"
);

// 5. GIFT_TRAN → SOLICIT_DEF (Solicitation type lookup - ACTIVE)
CreateRelationship(
    "GIFT_TRAN", "SOLICIT_CDE",
    "SOLICIT_DEF", "SOLICIT_CDE",
    "Single", true,
    "GIFT_TRAN[SOLICIT_CDE] → SOLICIT_DEF[SOLICIT_CDE]"
);

// 6. DimDate → GIFT_TRAN (Date dimension - ACTIVE)
CreateRelationship(
    "DimDate", "Date",
    "GIFT_TRAN", "GIFT_DTE",
    "Single", true,
    "DimDate[Date] → GIFT_TRAN[GIFT_DTE]"
);

// ============================================================================
// SECONDARY RELATIONSHIPS (INACTIVE) - Use with USERELATIONSHIP() in DAX
// These are marked INACTIVE to avoid ambiguous paths
// ============================================================================

// 7. DONOR_YEAR_SUM → DONOR_MASTER (INACTIVE - use USERELATIONSHIP when needed)
CreateRelationship(
    "DONOR_YEAR_SUM", "ID_NUM",
    "DONOR_MASTER", "ID_NUM",
    "Single", false,
    "DONOR_YEAR_SUM[ID_NUM] → DONOR_MASTER[ID_NUM] (INACTIVE)"
);

// 8. DONOR_CAMP_SUM → DONOR_MASTER (INACTIVE - use USERELATIONSHIP when needed)
CreateRelationship(
    "DONOR_CAMP_SUM", "ID_NUM",
    "DONOR_MASTER", "ID_NUM",
    "Single", false,
    "DONOR_CAMP_SUM[ID_NUM] → DONOR_MASTER[ID_NUM] (INACTIVE)"
);

// 9. DONOR_CAMP_SUM → CAMPAIGN (INACTIVE - use USERELATIONSHIP when needed)
CreateRelationship(
    "DONOR_CAMP_SUM", "CAMPAIGN_CDE",
    "CAMPAIGN", "CAMPAIGN_CDE",
    "Single", false,
    "DONOR_CAMP_SUM[CAMPAIGN_CDE] → CAMPAIGN[CAMPAIGN_CDE] (INACTIVE)"
);

// 10. ALUMNI_MASTER → DONOR_MASTER (INACTIVE - use USERELATIONSHIP when needed)
CreateRelationship(
    "ALUMNI_MASTER", "ID_NUM",
    "DONOR_MASTER", "ID_NUM",
    "Single", false,
    "ALUMNI_MASTER[ID_NUM] → DONOR_MASTER[ID_NUM] (INACTIVE)"
);

// ============================================================================
// STEP 3: MARK DATE TABLE
// ============================================================================

try
{
    if (Model.Tables.Contains("DimDate"))
    {
        var dimDate = Model.Tables["DimDate"];

        // Set as Date table using the Date column
        if (dimDate.Columns.Contains("Date"))
        {
            // In Tabular Editor 2, we set DataCategory to mark as date table
            dimDate.DataCategory = "Time";
            successLog.Add("Marked DimDate as Date table");
        }
    }
}
catch (Exception ex)
{
    errorLog.Add("Error marking Date table: " + ex.Message);
}

// ============================================================================
// STEP 4: CREATE FULL NAME CALCULATED COLUMN IN NAME_MASTER
// ============================================================================

try
{
    if (Model.Tables.Contains("NAME_MASTER"))
    {
        var nameMaster = Model.Tables["NAME_MASTER"];

        if (!nameMaster.Columns.Contains("Full_Name"))
        {
            var fullNameCol = nameMaster.AddCalculatedColumn(
                "Full_Name",
                @"TRIM([FIRST_NAME] & "" "" & [LAST_NAME])"
            );
            fullNameCol.DataType = DataType.String;
            successLog.Add("Created Full_Name calculated column in NAME_MASTER");
            itemsCreated++;
        }
        else
        {
            successLog.Add("Full_Name column already exists in NAME_MASTER");
        }
    }
}
catch (Exception ex)
{
    errorLog.Add("Error creating Full_Name column: " + ex.Message);
}

// ============================================================================
// STEP 5: SORT COLUMNS (if DimDate exists)
// ============================================================================

try
{
    if (Model.Tables.Contains("DimDate"))
    {
        var dimDate = Model.Tables["DimDate"];

        // Sort Month by MonthNum
        if (dimDate.Columns.Contains("Month") && dimDate.Columns.Contains("MonthNum"))
        {
            dimDate.Columns["Month"].SortByColumn = dimDate.Columns["MonthNum"];
            successLog.Add("Set Month to sort by MonthNum");
        }

        // Sort DayName by DayOfWeek
        if (dimDate.Columns.Contains("DayName") && dimDate.Columns.Contains("DayOfWeek"))
        {
            dimDate.Columns["DayName"].SortByColumn = dimDate.Columns["DayOfWeek"];
            successLog.Add("Set DayName to sort by DayOfWeek");
        }

        // Sort Quarter by QuarterNum
        if (dimDate.Columns.Contains("Quarter") && dimDate.Columns.Contains("QuarterNum"))
        {
            dimDate.Columns["Quarter"].SortByColumn = dimDate.Columns["QuarterNum"];
            successLog.Add("Set Quarter to sort by QuarterNum");
        }

        // Sort FiscalQuarter by FiscalQuarterNum
        if (dimDate.Columns.Contains("FiscalQuarter") && dimDate.Columns.Contains("FiscalQuarterNum"))
        {
            dimDate.Columns["FiscalQuarter"].SortByColumn = dimDate.Columns["FiscalQuarterNum"];
            successLog.Add("Set FiscalQuarter to sort by FiscalQuarterNum");
        }
    }
}
catch (Exception ex)
{
    errorLog.Add("Error setting sort orders: " + ex.Message);
}

// ============================================================================
// DISPLAY COMPLETION MESSAGE
// ============================================================================

string resultMessage = "═══════════════════════════════════════════════════════════════\n";
resultMessage += "   RELATIONSHIP & DATE TABLE SETUP COMPLETE (v2)\n";
resultMessage += "═══════════════════════════════════════════════════════════════\n\n";
resultMessage += "Items Created/Configured: " + itemsCreated + "\n\n";

if (successLog.Count > 0)
{
    resultMessage += "SUCCESS LOG:\n";
    resultMessage += "─────────────────────────────────────────────────────────────\n";
    foreach (var s in successLog)
    {
        resultMessage += "✓ " + s + "\n";
    }
}

if (errorLog.Count > 0)
{
    resultMessage += "\n\nERRORS:\n";
    resultMessage += "─────────────────────────────────────────────────────────────\n";
    foreach (var e in errorLog)
    {
        resultMessage += "✗ " + e + "\n";
    }
}

resultMessage += "\n═══════════════════════════════════════════════════════════════\n";
resultMessage += "ACTIVE RELATIONSHIPS (Main Star Schema):\n";
resultMessage += "─────────────────────────────────────────────────────────────\n";
resultMessage += "1. GIFT_TRAN[DONOR_ID] → DONOR_MASTER[ID_NUM] (Both)\n";
resultMessage += "2. DONOR_MASTER[ID_NUM] → NAME_MASTER[ID_NUM] (Both)\n";
resultMessage += "3. GIFT_TRAN[CAMPAIGN_CDE] → CAMPAIGN[CAMPAIGN_CDE]\n";
resultMessage += "4. GIFT_TRAN[CAT_COMP_1] → GIFT_CATEGORY[CAT_COMP_1]\n";
resultMessage += "5. GIFT_TRAN[SOLICIT_CDE] → SOLICIT_DEF[SOLICIT_CDE]\n";
resultMessage += "6. DimDate[Date] → GIFT_TRAN[GIFT_DTE]\n";
resultMessage += "\n";
resultMessage += "INACTIVE RELATIONSHIPS (Use USERELATIONSHIP in DAX):\n";
resultMessage += "─────────────────────────────────────────────────────────────\n";
resultMessage += "7. DONOR_YEAR_SUM[ID_NUM] → DONOR_MASTER[ID_NUM]\n";
resultMessage += "8. DONOR_CAMP_SUM[ID_NUM] → DONOR_MASTER[ID_NUM]\n";
resultMessage += "9. DONOR_CAMP_SUM[CAMPAIGN_CDE] → CAMPAIGN[CAMPAIGN_CDE]\n";
resultMessage += "10. ALUMNI_MASTER[ID_NUM] → DONOR_MASTER[ID_NUM]\n";
resultMessage += "═══════════════════════════════════════════════════════════════\n";
resultMessage += "\nRemember to SAVE your model! (File > Save or Ctrl+S)\n";

Info(resultMessage);
