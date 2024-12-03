using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Extensions.Logging;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Umbraco.Extensions;


namespace Webwonders.SpreadsheetHandler;

public interface IWebwondersSpreadsheetHandler
{
    /// <summary>
    /// Reads the spreadsheet and returns an IEnumerable of the mapped class
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="SpreadsheetFile">Spreadsheet to read</param>
    /// <param name="mapper">Function to map the spreadsheet to an enumerable of class T</param>
    /// <param name="sheetNumber">Sheetnumber to read, default 0 (first)</param>
    /// <param name="StopOnError">If true: stops reading after an error, result will be null, default false</param>
    /// <returns>WWSpreadsheet containing data</returns>
    IEnumerable<T>? ReadSpreadsheet<T>(string SpreadsheetFile, Func<WebwondersSpreadsheet, IEnumerable<T>> mapper, int sheetNumber = 0, bool StopOnError = false) where T : class;

    /// <summary>
    /// Reads the spreadsheet and returns a class with the rows and columns
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="SpreadsheetFile">Spreadsheet to read</param>
    /// <param name="sheetNumber">Sheetnumber to read, default 0 (first)</param>
    /// <param name="StopOnError">If true: stops reading after an error, result will be null, default false</param>
    /// <returns>WWSpreadsheet containing data</returns>
    WebwondersSpreadsheet? ReadSpreadsheet<T>(string SpreadsheetFile, int sheetNumber = 0, bool StopOnError = false) where T : class;

    /// <summary>
    /// Writes data to a spreadsheet, returning a memory stream
    /// </summary>
    /// <typeparam name="T">type of class to write</typeparam>
    /// <param name="data">IEnumerable of <typeparamref name="T"/> containing data</param>
    /// <param name="StopOnError">If true: stops writing after an error, result will be null, default false</param>
    /// <returns>Memorystream with spreadsheet</returns>
    MemoryStream? WriteSpreadsheet<T>(IEnumerable<T> data, bool StopOnError = false) where T : class;


    /// <summary>
    /// Reads the spreadsheet and returns a DataTable
    /// </summary>
    /// <param name="spreadsheetFile">Spreadsheet to read</param>
    /// <param name="spreadsheetSettings">Spreadsheet settings, default null (repeatedcolumns are neglected, since this returns a datatable)</param>
    /// <param name="sheetNumber">Sheet to read, default 0 (first)</param>
    /// <param name="stopOnError">If true: stops reading after an error, result will be null, default false</param>
    /// <returns>DataTable containing data</returns>
    DataTable? ReadSpreadsheet(string spreadsheetFile, SpreadsheetSettings? spreadsheetSettings = null, int sheetNumber = 0, bool stopOnError = false);



    /// <summary>
    /// Writes the data from the Datatable to a spreadsheet, returning a memorystream
    /// </summary>
    /// <param name="data">Datatable to write</param>
    /// <param name="spreadsheetSettings">Spreadsheetsettings, default null (repeatedcolumns are neglected, since this reads a DataTable)</param>
    /// <param name="StopOnError">If true: stops writing after an error, result will be null</param>
    /// <returns>Memorystream with spreadsheet</returns>
    MemoryStream? WriteSpreadsheet(DataTable data, SpreadsheetSettings? spreadsheetSettings = null, bool StopOnError = false);

    /// <summary>
    /// Creates the default SpreadsheetSettings based on the given type.
    /// </summary>
    /// <typeparam name="T">Type of class to create settings for.</typeparam>
    /// <returns></returns>
    SpreadsheetSettings CreateDefaultSettings<T>();

    /// <summary>
    /// Writes data to a spreadsheet, returning a memory stream.
    /// </summary>
    /// <typeparam name="T">type of class to write</typeparam>
    /// <param name="data">IEnumerable of <typeparamref name="T"/> containing data</param>
    /// <param name="spreadsheetSettings">Spreadsheet settings that can be supplied manually.</param>
    /// <param name="StopOnError">If true: stops writing after an error, result will be null, default false</param>
    /// <returns>Memorystream with spreadsheet</returns>
    MemoryStream? WriteSpreadsheet<T>(IEnumerable<T> data, SpreadsheetSettings spreadsheetSettings, bool StopOnError = false) where T : class;
}

public class WebwondersSpreadsheetHandler : IWebwondersSpreadsheetHandler
{

    private readonly ILogger<WebwondersSpreadsheetHandler> _logger;



    public WebwondersSpreadsheetHandler(ILogger<WebwondersSpreadsheetHandler> logger)
    {
        _logger = logger;
    }



    ///<inheritdoc />
    public IEnumerable<T>? ReadSpreadsheet<T>(string SpreadsheetFile, Func<WebwondersSpreadsheet, IEnumerable<T>> mapper, int sheetNumber = 0, bool StopOnError = false) where T : class
    {
        var result = ReadSpreadsheet<T>(SpreadsheetFile, sheetNumber, StopOnError);
        if (result != null)
        {
            return mapper(result);
        }
        return null;
    }


    ///<inheritdoc />
    public WebwondersSpreadsheet? ReadSpreadsheet<T>(string SpreadsheetFile, int sheetNumber = 0, bool StopOnError = false) where T : class
    {
        var spreadsheetFile = new FileInfo(SpreadsheetFile);
        if (spreadsheetFile == null || !spreadsheetFile.Exists)
        {
            _logger.LogError("Webwonders.Spreadsheethandler: File not found {FileName}", spreadsheetFile?.Name ?? string.Empty);
            return null;
        }

        // First: get the definition of the spreadsheet and its rows
        var spreadsheetSettings = CreateDefaultSettings<T>();

        // Then: read the spreadsheet and try to fit in the definition
        WebwondersSpreadsheet? result = ReadSpreadSheet(spreadsheetFile, spreadsheetSettings, sheetNumber, StopOnError);
        return result;
    }


    public MemoryStream? WriteSpreadsheet<T>(IEnumerable<T> data, SpreadsheetSettings spreadsheetSettings, bool StopOnError = false) where T : class
    {
        if (data == null || !data.Any())
        {
            _logger.LogError("Webwonders.Spreadsheethandler, WriteSpreadsheet: No data to write");
            return null;
        }

        if (spreadsheetSettings.IncludedColumns != null && spreadsheetSettings.IncludedColumns.Any())
        {
            spreadsheetSettings.ColumnDefinitions = spreadsheetSettings.ColumnDefinitions.Where(x => spreadsheetSettings.IncludedColumns.Contains(x.ColumnName)).ToList();
        }

        if (spreadsheetSettings.ColumnDefinitions.Count == 0)
        {
            _logger.LogError("Webwonders.Spreadsheethandler, WriteSpreadsheet: No columns defined for spreadsheet on Type {T}", typeof(T));
            return null;
        }

        // Next: write the data to the spreadsheet
        var workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet();
        IRow headerRow = sheet.CreateRow(0);
        for (int i = 0; i < spreadsheetSettings.ColumnDefinitions.Count; i++)
        {
            headerRow.CreateCell(i).SetCellValue(spreadsheetSettings.ColumnDefinitions[i].ColumnName);
        }
        foreach (T element in data)
        {
            IRow row = sheet.CreateRow(sheet.LastRowNum + 1);
            for (int i = 0; i < spreadsheetSettings.ColumnDefinitions.Count; i++)
            {
                var columnDefinition = spreadsheetSettings.ColumnDefinitions[i];

                object? value = null;
                if (!String.IsNullOrWhiteSpace(columnDefinition.PropertyName))
                {
                    value = element.GetType().GetProperty(columnDefinition.PropertyName)?.GetValue(element);
                }

                if (!spreadsheetSettings.EmptyCellsAllowed && value == null)
                {
                    _logger.LogError("Webwonders.Spreadsheethandler, WriteSpreadsheet: Empty cell found in column {Column} of row {Row}, but empty cells are not allowed", columnDefinition.ColumnName, row.RowNum);
                    if (StopOnError)
                    {
                        return null;
                    }
                }
                if (columnDefinition.ColumnRequired && value == null)
                {
                    _logger.LogError("Webwonders.Spreadsheethandler, WriteSpreadsheet: Required cell {Column} of row {Row} is empty", columnDefinition.ColumnName, row.RowNum);
                    if (StopOnError)
                    {
                        return null;
                    }
                }


                if (columnDefinition.RepeatedColumn && i == spreadsheetSettings.ColumnDefinitions.Count - 1)
                {
                    // repeated column (can only be the last one): write the value as an array
                    object[]? array = null;
                    if (value != null)
                    {
                        // make sure that value is an array or enumerable
                        if (value.GetType().IsArray)
                        {
                            array = (object[])value;
                        }
                        else if (value is IEnumerable enumerable)
                        {
                            array = enumerable.Cast<object>().ToArray();
                        }
                        if (array != null)
                        {
                            //create cells for each element in the array
                            for (int j = 0; j < array.Length; j++)
                            {
                                row.CreateCell(i + j).SetCellValue(array[j]?.ToString() ?? "");
                            }
                        }
                    }
                }
                else
                {
                    // Normal column, just write the value
                    row.CreateCell(i).SetCellValue(value?.ToString() ?? "");
                }
            }
        }

        var exportData = new MemoryStream();
        workbook.Write(exportData, true); // true to leave stream open
        return exportData;
    }

    ///<inheritdoc />
    public MemoryStream? WriteSpreadsheet<T>(IEnumerable<T> data, bool StopOnError = false) where T : class
    {
        return WriteSpreadsheet(data, CreateDefaultSettings<T>(), StopOnError);
    }


    ///<inheritdoc />
    public DataTable? ReadSpreadsheet(string spreadsheetFile, SpreadsheetSettings? spreadsheetSettings = null, int sheetNumber = 0, bool stopOnError = false)
    {
        DataTable? result = new();

        var file = new FileInfo(spreadsheetFile);
        if (file == null || !file.Exists)
        {
            _logger.LogError("Webwonders.Spreadsheethandler: File not found {FileName}", file?.Name ?? string.Empty);
            return null;
        }

        using var stream = file.Open(FileMode.Open);
        stream.Position = 0;

        var xssWorkbook = new XSSFWorkbook(stream);
        ISheet sheet = xssWorkbook.GetSheetAt(sheetNumber);

        // get header and create columns with header names
        IRow headerRow = sheet.GetRow(sheet.FirstRowNum);
        int cellCount = headerRow.LastCellNum;

        foreach (ICell cell in headerRow.Cells)
        {
            result.Columns.Add(new DataColumn(cell.StringCellValue));
        }


        for (int r = (sheet.FirstRowNum + 1); r <= sheet.LastRowNum; r++)
        {
            IRow row = sheet.GetRow(r);

            if (row == null)
            {
                continue;
            }

            // check on empty cells
            if (spreadsheetSettings != null && !spreadsheetSettings.EmptyCellsAllowed && row.Cells.Any(x => x.CellType == CellType.Blank))
            {
                _logger.LogError("Webwonders.Spreadsheethandler, ReadSpreadsheet: Empty cell found in row {Row}, but empty cells are not allowed", r);
                if (stopOnError)
                {
                    return null;
                }
            }

            DataRow dataRow = result.NewRow();
            for (int j = row.FirstCellNum; j <= cellCount; j++)
            {
                // if there are spreadsheetsettings: check for value of required cell
                if (spreadsheetSettings != null
                    && spreadsheetSettings.ColumnDefinitions.Any(x => x.ColumnRequired && x.ColumnName == headerRow.GetCell(j).StringCellValue)
                    && (row.GetCell(j) == null || string.IsNullOrEmpty(row.GetCell(j).ToString())))
                {
                    _logger.LogError("Webwonders.Spreadsheethandler, ReadSpreadsheet: Required cell {Column} of row {Row} is empty", headerRow.GetCell(j).StringCellValue, r);
                    if (stopOnError)
                    {
                        return null;
                    }
                }

                if (row.GetCell(j) != null)
                {
                    dataRow[j] = row.GetCell(j).ToString();
                }
            }
            result.Rows.Add(dataRow);
        }

        return result;
    }




    ///<inheritdoc />
    public MemoryStream? WriteSpreadsheet(DataTable data, SpreadsheetSettings? spreadsheetSettings = null, bool StopOnError = false)
    {

        // can check for emptycellsisallowed and required cells here
        if (data.Rows == null || data.Rows.Count <= 0)
        {
            _logger.LogError("Webwonders.Spreadsheethandler, WriteSpreadsheet: No data to write");
            return null;
        }
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet();
        var headerRow = sheet.CreateRow(0);

        for (int i = 0; i < data.Columns.Count; i++)
        {
            headerRow.CreateCell(i).SetCellValue(data.Columns[i].ColumnName);
        }
        for (int i = 0; i < data.Rows.Count; i++)
        {
            var row = sheet.CreateRow(i + 1);
            for (int j = 0; j < data.Columns.Count; j++)
            {
                if (spreadsheetSettings != null
                    && spreadsheetSettings.ColumnDefinitions.Any(x => x.ColumnRequired && x.ColumnName == data.Columns[j].ColumnName)
                    && (data.Rows[i][j] == null || string.IsNullOrEmpty(data.Rows[i][j].ToString())))
                {
                    _logger.LogError("Webwonders.Spreadsheethandler, WriteSpreadsheet: Required cell {Column} of row {Row} is empty", data.Columns[j].ColumnName, i + 1);
                    if (StopOnError)
                    {
                        return null;
                    }
                }
                row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
            }
        }

        var stream = new MemoryStream();
        workbook.Write(stream, true); // true to leave stream open
        return stream;

    }



    /// <summary>
    /// Read the Spreadsheet and return as class Spreadsheet
    /// where the rows have cells that couple the column- and propertyname and the values
    /// </summary>
    /// <param name="spreadsheetFile"></param>
    /// <param name="spreadsheetSettings"></param>
    /// <param name="StopOnError"></param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    private WebwondersSpreadsheet? ReadSpreadSheet(FileInfo spreadsheetFile, SpreadsheetSettings spreadsheetSettings, int sheetNumber, bool StopOnError)
    {

        WebwondersSpreadsheet? result = null;

        if (spreadsheetFile == null || !spreadsheetFile.Exists)
        {
            _logger.LogError("Webwonders.Spreadsheethandler: File not found {FileName}", spreadsheetFile?.Name ?? string.Empty);
            return null;
        }

        using var stream = spreadsheetFile.Open(FileMode.Open);
        stream.Position = 0;

        var xssWorkbook = new XSSFWorkbook(stream);
        ISheet sheet = xssWorkbook.GetSheetAt(sheetNumber);


        var ColumnNames = new List<string>();
        var ColumnValues = new Dictionary<int, string>();


        IRow headerRow = sheet.GetRow(0);
        int cellCount = headerRow.LastCellNum;


        for (int r = (sheet.FirstRowNum); r <= sheet.LastRowNum; r++)
        {

            // Reinit the values
            ColumnValues = new Dictionary<int, string>();

            result ??= new WebwondersSpreadsheet();

            // TODO empty rows
            if (sheet.GetRow(r) is IRow currentRow)
            {
                if (!spreadsheetSettings.EmptyCellsAllowed && currentRow.Cells.Any(c => c.CellType == CellType.Blank))
                {
                    _logger.LogError("Error in reading spreadsheet, row {Row} contains empty cells", r);
                    if (StopOnError)
                    {
                        return null;
                    }
                }

                bool firstRowErrorLogged = false;
                // Iterate columns 
                for (int j = currentRow.FirstCellNum; j < cellCount; j++)
                {
                    if (currentRow.GetCell(j) is ICell currentCell)
                    {
                        string currentCellValue = (currentCell.CellType != CellType.Blank) ? (currentCell.ToString() ?? "")
                                                                                           : "";
                        // Is row the first: add to header, otherwise add to value
                        if (currentRow.RowNum == headerRow.RowNum)
                        {
                            // Headerrow can not contain empty cells: it would be impossible to match the columns to the properties
                            // so log an error and discard the column or stop depending on the StopOnError flag
                            if (currentCell.CellType == CellType.Blank)
                            {
                                if (!firstRowErrorLogged)
                                {
                                    _logger.LogError("Error in reading spreadsheet, first row contains empty column. Column is skipped.");
                                    firstRowErrorLogged = true;
                                }
                                if (StopOnError)
                                {
                                    return null;
                                }
                            }
                            ColumnNames.Add(currentCellValue);
                        }
                        else
                        {
                            ColumnValues.Add(currentCell.ColumnIndex, currentCellValue);
                        }

                    }
                }

                // for all rows except the header: create a WebwondersSpreadsheetRow
                // and all contained WWSpreadsheetCells
                if (currentRow.RowNum != headerRow.RowNum)
                {
                    if (ColumnValues.Where(x => !String.IsNullOrWhiteSpace(x.Value)).Any())
                    {
                        result.Rows.Add(new WebwondersSpreadsheetRow(currentRow.RowNum - headerRow.RowNum + 1)); // number in spreadsheetrow is actual number + 2,
                                                                                                                 // so it skips the title and starts at 1 instead of 0
                                                                                                                 // this wil make the muber the same as the actual spreadsheet row number
                        for (int i = 0; i < ColumnNames.Count; i++)
                        {
                            if (!String.IsNullOrEmpty(ColumnNames[i]))
                            {
                                string currentColumnsValue = "";
                                if (ColumnValues.ContainsKey(i))
                                {
                                    currentColumnsValue = ColumnValues[i];
                                }
                                // if the last columns are to be repeated and we are in one of those columns:
                                // get the definition of the column to be repeated.
                                // otherwise: get the definition of the current column
                                SpreadsheetColumnSettings? propDef = null;
                                if (spreadsheetSettings.RepeatedFromColumn != null
                                    && spreadsheetSettings.RepeatedFromColumn.Value > 0
                                    && i >= spreadsheetSettings.RepeatedFromColumn.Value)
                                {
                                    propDef = spreadsheetSettings.ColumnDefinitions.FirstOrDefault(x => x.RepeatedColumn);
                                }
                                else
                                {
                                    propDef = spreadsheetSettings.ColumnDefinitions.FirstOrDefault(x => x.ColumnName?.ToLower() == ColumnNames[i].ToLower());
                                }

                                // if the definition is found:
                                // add the row when valid
                                // use the correct columnname by getting it from the previously collected names
                                if (propDef != null)
                                {
                                    if (propDef.ColumnRequired && currentColumnsValue.IsNullOrWhiteSpace())
                                    {
                                        _logger.LogError("Error in reading spreadsheet, row {row},  column {counter}. Required column is empty, row is skipped.", currentRow.RowNum - headerRow.RowNum + 1, i + 1);
                                        if (StopOnError)
                                        {
                                            return result;
                                        }
                                    }
                                    else
                                    {
                                        // only add if required and given or not required
                                        //result.Rows[currentRow.RowNum - headerRow.RowNum - 1].Cells.Add(new WWSpreadsheetCell { ColumName = ColumnNames[i], ColumnValue = ColumnValues[i], PropertyName = propDef.PropertyName, IsRequired = propDef.ColumnRequired });
                                        // add to the last row, we always go from top to bottom in the spreadsheet
                                        result.Rows.Last().Cells.Add(new WebwondersSpreadsheetCell { ColumName = ColumnNames[i], ColumnValue = currentColumnsValue, PropertyName = propDef.PropertyName, IsRequired = propDef.ColumnRequired });
                                    }
                                }
                            }
                            // in later rows: if the property is found in the definition:
                            // make cell of columname, value and propertyName and add to current row in fullspreadsheet
                            // only if the column has a name to avoid probles with KeyValue (columns without title cannot be mapped)
                            // row - startRow - 1, because first row contains columnames and is ignored in result
                        }
                    }
                } // Save row
            }


        }

        return result;
    }

    public SpreadsheetSettings CreateDefaultSettings<T>()
    {
        var spreadsheetSettings = new SpreadsheetSettings();

        if (typeof(T).GetCustomAttribute<SpreadsheetAttribute>() is SpreadsheetAttribute spreadsheetAttribute)
        {
            spreadsheetSettings.EmptyCellsAllowed = spreadsheetAttribute.EmptyCellsAllowed;
            spreadsheetSettings.RepeatedFromColumn = spreadsheetAttribute?.RepeatedFromColumn;
        }
        foreach (PropertyInfo propInfo in typeof(T).GetProperties())
        {
            if (propInfo.GetCustomAttribute<SpreadsheetColumnAttribute>() is SpreadsheetColumnAttribute columnAttribute)
            {
                spreadsheetSettings.ColumnDefinitions.Add(new SpreadsheetColumnSettings
                {
                    PropertyName = propInfo.Name,
                    ColumnName = columnAttribute?.ColumnName ?? "",
                    ColumnRequired = columnAttribute?.ColumnRequired ?? false,
                    RepeatedColumn = columnAttribute?.RepeatedColumn ?? false,
                });
            }
        }

        return spreadsheetSettings;
    }
}
