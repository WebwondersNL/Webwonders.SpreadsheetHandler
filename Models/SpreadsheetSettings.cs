using System.Collections.Generic;

namespace Webwonders.SpreadsheetHandler;


public class SpreadsheetColumnSettings
{
    public string? PropertyName { get; set; }

    public string? ColumnName { get; set; }

    public bool ColumnRequired { get; set; }

    // unused in the DataTable implementation: datatables have a fixed number of columns
    public bool RepeatedColumn { get; set; }
}

public class SpreadsheetSettings
{
    public bool EmptyCellsAllowed { get; set; }

    // unused in the DataTable implementation: datatables have a fixed number of columns
    public int? RepeatedFromColumn { get; set; }

    public List<SpreadsheetColumnSettings> ColumnDefinitions { get; set; }

    /// <summary>
    /// Used to configure what columns are used. By default, this is an empty list in which case all columns are exported.
    /// </summary>
    public List<string> IncludedColumns { get; set; } = new List<string>();

    public SpreadsheetSettings()
    {
        ColumnDefinitions = new List<SpreadsheetColumnSettings>();
    }
}