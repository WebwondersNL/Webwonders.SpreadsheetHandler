# Webwonders.SpreadsheetHandler


## About
Read spreadsheets into datatables or custom class.
Write custom data from class or datatable into spreadsheets.


## How to use
The package defines a service: IWebwondersSpreadsheetHandler which can be injected into your project.
This service has the following methods: 
	
	IEnumerable<T>? ReadSpreadsheet<T>(string SpreadsheetFile, Func<WebwondersSpreadsheet, IEnumerable<T>> mapper, int sheetNumber = 0, bool StopOnError = false) where T : class;
	  - Reads a spreadsheet into a custom class, using a mapper function to map the spreadsheet to the class.
	
	WebwondersSpreadsheet? ReadSpreadsheet<T>(string SpreadsheetFile, int sheetNumber = 0, bool StopOnError = false) where T : class;
	  - Reads a spreadsheet into a WebwondersSpreadsheet object, which contains a list of WebwondersSpreadsheetRow objects.
	
	MemoryStream? WriteSpreadsheet<T>(IEnumerable<T> data, bool StopOnError = false) where T : class;
	  - Writes a spreadsheet from a list of custom class objects.

	DataTable? ReadSpreadsheet(string spreadsheetFile, SpreadsheetSettings? spreadsheetSettings = null, int sheetNumber = 0, bool stopOnError = false);
      - Reads a spreadsheet into a datatable.
	
	MemoryStream? WriteSpreadsheet(DataTable data, SpreadsheetSettings? spreadsheetSettings = null, bool StopOnError = false);
      - Writes a spreadsheet from a datatable.

All methods have a StopOnError parameter. If this is set to true, the method will return null if an error occurs. 
All errors will be written to the log, independent of the StopOnError parameter.

The ReadSpreadsheet methods have a sheetNumber parameter. This is the index of the sheet to read. The default is 0, which is the first sheet.

The WriteSpreadsheet methods both return a Memorystream of the spreadsheetfile. This can be used to write the file to a filestream or to a httpresponse.

Mapping of the custom class to the Spreadsheet-definition is done by attributes added to the type definition of the class:

[Spreadsheet(EmptyCellsAllowed = true, RepeatedFromColumn = 5)]
[SpreadsheetColumn(ColumnName = "street", ColumnRequired = true, RepeatedColumn = false )]

The Spreadsheet attribute defines the settings for the spreadsheet. The SpreadsheetColumn attribute defines the settings for the columns in the spreadsheet.
The RepeatedFromColumnAttribute and the RepeatedColumn can be used when the last columns of the row are repeated. The RepeatedFromColumnAttribute defines the 
column from which the repeated columns start. The RepeatedColumn attribute defines if the column is repeated. The default is false. When a column is repeated, it
expects an array for the property connected to that column.


### Example
```csharp

// Class definition
[Spreadsheet(EmptyCellsAllowed = true)]
public class Address {
	
	[SpreadsheetColumn(ColumnName = "Street", ColumnRequired = true )]
	public string Street { get; set; }

	[SpreadsheetColumn(ColumnName = "Housenumber", ColumnRequired = true)]
	public string HouseNumber { get; set; }

	[SpreadsheetColumn(ColumnName = "Addition")]
	public string Addition { get; set; }

	[SpreadsheetColumn(ColumnName = "Postal code")]
	public string PostalCode { get; set; }

	[SpreadsheetColumn(ColumnName = "City")]
	public string City { get; set; }
}



// Mapper for spreadsheet
readonly Func<WebwondersSpreadsheet, IEnumerable<Address>> MapAddress = (spreadsheet) =>
{
	var addresses = new List<Address>();
	foreach (var row in spreadsheet.Rows)
	{
		var address = new Address
		{
			Street = row.Cells[0].ColumnValue,
			HouseNumber = row.Cells[1].ColumnValue,
			HouseAddition = row.Cells[2].ColumnValue,
			PostalCode = row.Cells[3].ColumnValue,
			City = row.Cells[4].ColumnValue,
		};
		addresses.Add(address);
	}
	return addresses;
};



// Code that reads and writes spreadsheet:
string spreadsheetFilePath = Path.Combine(webrootPath, "Spreadsheets", "AddressList.xlsx");
var addresses = _spreadsheet.ReadSpreadsheet<Address>(spreadsheetFilePath, MapAddress);
		
if (addresses != null && addresses.Any())
{
	using var memoryStream = _spreadsheet.WriteSpreadsheet(test);
	if (memoryStream != null && memoryStream.Length > 0)
	{
		return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "addresslist.xlsx");
	}
}


```
