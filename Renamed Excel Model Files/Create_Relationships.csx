// =============================================================================
// Tabular Editor 2 C# Script - Create Relationships for Advancement Data Model
// =============================================================================
//
// INSTRUCTIONS:
// 1. Open your Power BI model in Tabular Editor 2
// 2. Go to File > Open > From DB (connect to your Power BI Desktop model)
// 3. Go to C# Script tab (or press Ctrl+Shift+C)
// 4. Paste this entire script
// 5. Press F5 or click Run to execute
// 6. Save changes back to Power BI (Model > Deploy)
//
// IMPORTANT: Table names must match exactly as imported in Power BI.
// If your table names differ, update them in the script below.
// =============================================================================

// Define table names (update these if your Power BI table names differ)
var tblConstituents = "Constituents";
var tblAlumni = "Alumni";
var tblCampaigns = "Campaigns";
var tblGiftCategories = "Gift Categories";
var tblSolicitations = "Solicitations";
var tblFiscalYears = "Fiscal Years";
var tblGiftTransactions = "Gift Transactions";
var tblDonorYearSummary = "Donor Year Summary";
var tblDonorCampaignSummary = "Donor Campaign Summary";

// Helper function to safely create a relationship
Action<string, string, string, string, string> CreateRelationship = (fromTable, fromColumn, toTable, toColumn, relationshipName) =>
{
    try
    {
        // Check if tables exist
        if (!Model.Tables.Contains(fromTable))
        {
            Warning("Table not found: " + fromTable);
            return;
        }
        if (!Model.Tables.Contains(toTable))
        {
            Warning("Table not found: " + toTable);
            return;
        }

        // Check if columns exist
        if (!Model.Tables[fromTable].Columns.Contains(fromColumn))
        {
            Warning("Column not found: " + fromTable + "[" + fromColumn + "]");
            return;
        }
        if (!Model.Tables[toTable].Columns.Contains(toColumn))
        {
            Warning("Column not found: " + toTable + "[" + toColumn + "]");
            return;
        }

        // Check if relationship already exists
        var existingRel = Model.Relationships
            .FirstOrDefault(r =>
                r.FromTable.Name == toTable &&
                r.FromColumn.Name == toColumn &&
                r.ToTable.Name == fromTable &&
                r.ToColumn.Name == fromColumn);

        if (existingRel != null)
        {
            Info("Relationship already exists: " + relationshipName);
            return;
        }

        // Create the relationship (From = Many side, To = One side)
        var rel = Model.AddRelationship();
        rel.FromTable = Model.Tables[toTable];
        rel.FromColumn = Model.Tables[toTable].Columns[toColumn];
        rel.ToTable = Model.Tables[fromTable];
        rel.ToColumn = Model.Tables[fromTable].Columns[fromColumn];
        rel.IsActive = true;

        Info("Created relationship: " + relationshipName);
    }
    catch (Exception ex)
    {
        Error("Failed to create relationship '" + relationshipName + "': " + ex.Message);
    }
};

Info("=== Starting Relationship Creation ===");
Info("");

// =============================================================================
// FISCAL YEARS RELATIONSHIPS (Dimension → Fact Tables)
// =============================================================================
Info("Creating Fiscal Years relationships...");

CreateRelationship(
    tblFiscalYears, "Fiscal Year Code",
    tblGiftTransactions, "Fiscal Year",
    "Fiscal Years → Gift Transactions"
);

CreateRelationship(
    tblFiscalYears, "Fiscal Year Code",
    tblDonorYearSummary, "Fiscal Year",
    "Fiscal Years → Donor Year Summary"
);

CreateRelationship(
    tblFiscalYears, "Fiscal Year Code",
    tblDonorCampaignSummary, "Fiscal Year",
    "Fiscal Years → Donor Campaign Summary"
);

// =============================================================================
// CONSTITUENTS RELATIONSHIPS (Dimension → Fact Tables + Alumni)
// =============================================================================
Info("");
Info("Creating Constituents relationships...");

CreateRelationship(
    tblConstituents, "Constituent ID",
    tblGiftTransactions, "Donor ID",
    "Constituents → Gift Transactions"
);

CreateRelationship(
    tblConstituents, "Constituent ID",
    tblDonorYearSummary, "Donor ID",
    "Constituents → Donor Year Summary"
);

CreateRelationship(
    tblConstituents, "Constituent ID",
    tblDonorCampaignSummary, "Donor ID",
    "Constituents → Donor Campaign Summary"
);

CreateRelationship(
    tblConstituents, "Constituent ID",
    tblAlumni, "Constituent ID",
    "Constituents → Alumni"
);

// =============================================================================
// CAMPAIGNS RELATIONSHIPS (Dimension → Fact Tables)
// =============================================================================
Info("");
Info("Creating Campaigns relationships...");

CreateRelationship(
    tblCampaigns, "Campaign Code",
    tblGiftTransactions, "Campaign Code",
    "Campaigns → Gift Transactions"
);

CreateRelationship(
    tblCampaigns, "Campaign Code",
    tblDonorCampaignSummary, "Campaign Code",
    "Campaigns → Donor Campaign Summary"
);

// =============================================================================
// SOLICITATIONS RELATIONSHIPS (Dimension → Fact Tables)
// =============================================================================
Info("");
Info("Creating Solicitations relationships...");

CreateRelationship(
    tblSolicitations, "Solicitation Code",
    tblGiftTransactions, "Solicitation Code",
    "Solicitations → Gift Transactions"
);

// =============================================================================
// SUMMARY
// =============================================================================
Info("");
Info("=== Relationship Creation Complete ===");
Info("");
Info("NOTES:");
Info("1. Gift Categories requires a composite key (Category Type + Category Code).");
Info("   You may need to create a calculated column combining these fields");
Info("   in both Gift Categories and Gift Transactions tables for a proper relationship.");
Info("");
Info("2. Remember to save changes back to Power BI (Model > Deploy or Ctrl+Shift+D)");
