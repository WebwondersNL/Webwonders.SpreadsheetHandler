namespace Webwonders.SpreadsheetHandler;



/// <summary>
/// Attribute that maps a class to map a spreadsheet
/// </summary>
[System.AttributeUsage(System.AttributeTargets.Class)]
public class SpreadsheetAttribute : System.Attribute
{
	/// <summary>
	/// Are empty cells allowed in the spreadsheet?
	/// </summary>
	public bool EmptyCellsAllowed { get; set; }

	/// <summary>
	/// From this column the data gets repeated until there is no more data. 
	/// This needs to be the last column and the type needs to be an Array, IEnumerable or a List
	/// Columns start at base 0.
	/// </summary>
	public int RepeatedFromColumn { get; set; }
}



/// <summary>
/// Attribute that maps a property to a column in a spreadsheet
/// Columns are written to the spreadheet in the order of the properties in the class
/// </summary>
[System.AttributeUsage(System.AttributeTargets.Property)]
public class SpreadsheetColumnAttribute : System.Attribute
{
	/// <summary>
	/// Name of the column in the spreadsheet
	/// </summary>
	public string? ColumnName { get; set; }

	/// <summary>
	/// Is this column required to have a value?
	/// </summary>
	public bool ColumnRequired { get; set; }

	/// <summary>
	/// Does this column get repeated until there is no more data?
	/// Only applicable for the last column in the spreadsheet
	/// </summary>
	public bool RepeatedColumn { get; set; }
}
