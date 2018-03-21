const authorize = require('./authorize.js');
const SheetsHelper  = require('./sheets.js');
const {Spreadsheet,NamedRanges,GridRange,SpreadsheetProperties,Sheet,SheetProperties,GridProperties,GridData,RowData,CellData,ExtendedValue} = require('./models.js');

/**
 * 
 * @param {*} title 
 * @param {*} sheets : is an array of object which has the following format:
 * {
 *  title:string, 
 *  rowCount:number,
 *  columnCount:number,
 *  frozenRowCount:number,
 *  frozenColumnCount:number,
 *  hideGridlines:bool,
 *  data: array of objects which has the following format: 
 *      {
 *          startRow: number (0-based),
 *          startColumn: number (0-based),
 *          rowData: array of objects which has the following format:
 *              {
 *                  value: number/string/bool/formula
 *                  note: string
 *              }
 *      }
 * }
 * @param {*} callback 
 */
function CreateSpreadSheet(title,sheets,callback)
{
    
    sheets = sheets.map((sheet)=>{
        if(sheet.data)
            sheet.data = sheet.data.map((dt)=>{
                if(dt.rowData)
                    dt.rowData = dt.rowData.map((row)=>{
                        let result = {userEnteredValue:{}};
                        switch(typeof row.value)
                        {
                            case 'number': result.userEnteredValue.numberValue = row.value;break;
                            case 'string': {
                                if(row.value.startsWith('='))
                                    result.userEnteredValue.formulaValue = row.value;
                                else
                                    result.userEnteredValue.stringValue = row.value.toString();
                                break;
                            };
                            case 'boolean': result.userEnteredValue.boolValue = row.value;break;
                            default: result.userEnteredValue.stringValue = row.value.toString();
                        }
                        result.note = row.note;
                        return result;                
                    });
                return {
                    startRow:dt.startRow,
                    startColumn: dt.startColumn,
                    rowData:{
                        values:dt.rowData
                    }
                };
            });
        return {
            properties:{
                title:sheet.title,
                gridProperties: {
                    rowCount: sheet.rowCount,
                    columnCount: sheet.columnCount,
                    frozenRowCount: sheet.frozenRowCount,
                    frozenColumnCount: sheet.frozenColumnCount,
                    hideGridlines: sheet.hideGridlines
                }
            },
            data:sheet.data
        }
    })
    const resource = {
        properties: {title},
        sheets
    };
    authorize.start((auth)=>{
        const helper =new SheetsHelper(auth);
        helper.createSpreadsheet(resource,(error,response)=>{
            if(error)
                callback(error);
            else
                callback(null,response);
        });
    });
}

function GetSpreadSheet(spreadSheetId,callback)
{
    authorize.start((auth)=>{
        const helper =new SheetsHelper(auth);
        helper.getSpreadSheet(spreadSheetId,callback);
    });
}

function UpdateSpreadSheet(spreadSheetId,title,callback)
{
    authorize.start((auth)=>{
        const helper =new SheetsHelper(auth);
        helper.updateSpreadSheet(spreadSheetId,title,callback);
    });
}

function UpdateSheetProperties(spreadsheetId,sheetName,title,rowCount,columnCount,frozenRowCount,frozenColumnCount,hideGridlines,callback)
{
    authorize.start((auth)=>{
        const helper =new SheetsHelper(auth);
        helper.getSpreadSheet(spreadsheetId,(error,spreadSheet)=>{
            let sheet = spreadSheet.sheets.find(x=>x.properties.title==sheetName);
            console.log(sheet);
            if(sheet)
                helper.updateSheetProperties(spreadsheetId,sheet.properties.sheetId,title,sheet.properties.index,rowCount,columnCount,frozenRowCount,frozenColumnCount,hideGridlines,callback);
            else
                callback(new Error(`${sheetName} does not exist!`));
        });
        
    });
}


function AddSheet(spreadsheetId,title,rowCount,columnCount,frozenRowCount,frozenColumnCount,hideGridlines,callback)
{
    authorize.start((auth)=>{
        const helper =new SheetsHelper(auth);
        helper.addSheet(spreadsheetId,title,rowCount,columnCount,frozenRowCount,frozenColumnCount,hideGridlines,callback);

        
    });
}


function DeleteSheet(spreadsheetId,sheetName,callback)
{
    authorize.start((auth)=>{
        const helper =new SheetsHelper(auth);
        helper.getSpreadSheet(spreadsheetId,(error,spreadSheet)=>{
            let sheet = spreadSheet.sheets.find(x=>x.properties.title==sheetName);
            if(sheet)
                helper.deleteSheet(spreadsheetId,sheet.properties.sheetId,callback);
            else
                callback(new Error(`${sheetName} does not exist!`));
        });
    });
}

function UpdateRows(spreadsheetId,sheetName,rows,startRow,startColumn,callback)
{
    let formattedRows = rows.map((row)=>{
        row = row.map((data)=>{
            let result = {userEnteredValue:{}};
            switch(typeof data)
            {
                case 'number': result.userEnteredValue.numberValue = data;break;
                case 'string': {
                    if(data.startsWith('='))
                        result.userEnteredValue.formulaValue = data;
                    else
                        result.userEnteredValue.stringValue = data.toString();
                    break;
                };
                case 'boolean': result.userEnteredValue.boolValue = data;break;
                default: result.userEnteredValue.stringValue = data.toString();
            }
            return result; 
        });
        return {
            values:row
        }
    });
    authorize.start((auth)=>{
        const helper =new SheetsHelper(auth);
        helper.getSpreadSheet(spreadsheetId,(error,spreadSheet)=>{
            let sheet = spreadSheet.sheets.find(x=>x.properties.title==sheetName);
            if(sheet)
                helper.updateCells(spreadsheetId,sheet.properties.sheetId,formattedRows,startRow,startColumn,callback);
            else
                callback(new Error(`${sheetName} does not exist!`));
        });
    });
}

function AddRows(spreadsheetId,sheetName,rows,callback)
{
    let formattedRows = rows.map((row)=>{
        row = row.map((data)=>{
            let result = {userEnteredValue:{}};
            switch(typeof data)
            {
                case 'number': result.userEnteredValue.numberValue = data;
                    break;
                case 'string':
                    if (data.startsWith('='))
                        result.userEnteredValue.formulaValue = data;
                    else
                        result.userEnteredValue.stringValue = data.toString();
                    break;
                case 'boolean': result.userEnteredValue.boolValue = data;
                    break;
                case 'object': result.userEnteredValue.stringValue = '';
                    break;
                default: result.userEnteredValue.stringValue = data.toString();
            }
            return result; 
        });
        return {
            values:row
        }
    });
    authorize.start((auth)=>{
        const helper =new SheetsHelper(auth);
        helper.getSpreadSheet(spreadsheetId,(error,spreadSheet)=>{
            let sheet = spreadSheet.sheets.find(x=>x.properties.title==sheetName);
            if(sheet)
                helper.addCells(spreadsheetId,sheet.properties.sheetId,formattedRows,callback);
            else
                callback(new Error(`${sheetName} does not exist!`));
        });
    });
}

function DeleteRangedCells(spreadsheetId,sheetName,startRow,endRow,startColumn,endColumn,shiftDimension,callback)
{
    authorize.start((auth)=>{
        const helper =new SheetsHelper(auth);
        helper.getSpreadSheet(spreadsheetId,(error,spreadSheet)=>{
            let sheet = spreadSheet.sheets.find(x=>x.properties.title==sheetName);
            if(sheet)
                helper.deleteCells(spreadsheetId,sheet.properties.sheetId,startRow,endRow,startColumn,endColumn,shiftDimension,callback);
            else
                callback(new Error(`${sheetName} does not exist!`));
        });
    });
}

function GetRangedCells(spreadsheetId,ranges,majorDimension,callback)
{
    authorize.start((auth)=>{
        const helper =new SheetsHelper(auth);

        helper.getCells(spreadsheetId,ranges,majorDimension,null,null,callback);
    });
}
module.exports = {CreateSpreadSheet,GetSpreadSheet,UpdateSpreadSheet,UpdateSheetProperties,AddSheet,DeleteSheet,UpdateRows,AddRows,DeleteRangedCells,GetRangedCells};





