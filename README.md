# Webwonders.SpreadsheetHandler

**Webwonders.SpreadsheetHandler** is a .NET library for reading and writing spreadsheets using datatables or custom classes. It provides an injectable `IWebwondersSpreadsheetHandler` service with methods for mapping spreadsheets to classes, handling multiple sheets, and managing errors. Attributes define spreadsheet settings and column mappings, supporting repeated columns. The library returns spreadsheets as `MemoryStream` objects for easy file handling.

---

## Features
- Read spreadsheets into DataTables or custom classes
- Write DataTables to spreadsheets
- Attribute-based column mapping
- Supports multiple sheets and repeated columns
- Injectable service for easy integration

---

## Installation

Use NuGet to install:
```sh
Install-Package Webwonders.SpreadsheetHandler
```

---

## Usage

### Dependency Injection Setup
```csharp
services.AddTransient<IWebwondersSpreadsheetHandler, WebwondersSpreadsheetHandler>();
```

### Reading a Spreadsheet
```csharp
using (var stream = new FileStream("file.xlsx", FileMode.Open))
{
    var result = _spreadsheetHandler.ReadSpreadsheet<MyClass>(stream);
}
```

### Writing a Spreadsheet
```csharp
var dataTable = new DataTable();
// Populate DataTable
var stream = _spreadsheetHandler.WriteSpreadsheet(dataTable);
File.WriteAllBytes("output.xlsx", stream.ToArray());
```

### Defining Attributes in a Model Class
```csharp
[SpreadsheetSheet(Name = "Sheet1")]
public class MyClass
{
    [SpreadsheetColumn(Name = "ID")]
    public int Id { get; set; }

    [SpreadsheetColumn(Name = "Name")]
    public string Name { get; set; }
}
```

---

## Contributing
1. Fork the repository.
2. Create a new branch.
3. Commit your changes.
4. Push the branch.
5. Create a Pull Request.

---

## License
This project is licensed under the MIT License. See the LICENSE file for details.

---

[GitHub Repository](https://github.com/WebwondersNL/Webwonders.SpreadsheetHandler)

