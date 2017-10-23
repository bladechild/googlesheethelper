class Spreadsheet
{
    constructor(options)
    {
        this.properties = options.properties; // spreadsheetProperties
        this.sheets = options.sheets; // [] of sheet
        this.namedRanges = options.namedRanges;
    }
}

class NamedRanges
{
    constructor(options)
    {
        this.name = options.name;
        this.range = options.range;
    }
}

class GridRange
{
    constructor(options)
    {
        this.startRowIndex = options.startRowIndex;
        this.endRowIndex = options.endRowIndex;
        this.startColumnIndex = options.startColumnIndex;
        this.endColumnIndex = options.endColumnIndex;
    }
}
class SpreadsheetProperties
{
    constructor(options)
    {
        this.title = options.title;
    }
}


class Sheet
{
    constructor(options)
    {
        this.properties = options.properties;
        this.data = options.data; // array of Grid Data
    }
}

class SheetProperties
{
    constructor(options)
    {
        this.title = options.title;
        this.gridProperties = options.gridProperties;
    }
}

class GridProperties
{
    constructor(options)
    {
        this.rowCount = options.rowCount;
        this.columnCount = options.columnCount;
        this.frozenRowCount = options.frozenRowCount;
        this.frozenColumnCount = options.frozenColumnCount;
        this.hideGridlines = options.hideGridlines;
    }
}

class GridData
{
    constructor(options)
    {
        this.startRow = options.startRow;
        this.startColumn = options.startColumn;
        this.rowData = options.rowData;
    }
}

class RowData
{
    constructor(options)
    {
        this.values = options.values; // array of cell data
    }
}

class CellData
{
    constructor(options)
    {
        this.userEnteredValue = options.userEnteredValue;
        this.note = options.note;
    }
}

class ExtendedValue
{
    constructor(options)
    {
        this.numberValue = options.numberValue;
        this.stringValue = options.stringValue;
        this.boolValue = options.boolValue;
        this.formulaValue = options.formulaValue;
    }
}
module.exports = {Spreadsheet,NamedRanges,GridRange,SpreadsheetProperties,Sheet,SheetProperties,GridProperties,GridData,RowData,CellData,ExtendedValue}