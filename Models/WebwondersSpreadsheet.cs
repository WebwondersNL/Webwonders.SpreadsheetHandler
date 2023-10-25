using System.Collections.Generic;

namespace Webwonders.SpreadsheetHandler;



public class WebwondersSpreadsheetCell
{
    public string? ColumName { get; set; }

    public string? ColumnValue { get; set; }

    public string? PropertyName { get; set; }

    public bool IsRequired { get; set; }
}

public class WebwondersSpreadsheetRow
{
    public int Number { get; set; }

    public List<WebwondersSpreadsheetCell> Cells { get; set; }

    public WebwondersSpreadsheetRow(int number)
    {
        Number = number;
        Cells = new List<WebwondersSpreadsheetCell>();
    }
}

public class WebwondersSpreadsheet
{
    public List<WebwondersSpreadsheetRow> Rows { get; set; }
    public WebwondersSpreadsheet()
    {
        Rows = new List<WebwondersSpreadsheetRow>();
    }
}
