var google = require('googleapis');
var util = require('util');


var SheetsHelper = function(auth) {
  this.service = google.sheets({version: 'v4', auth: auth});
};

// create a spread sheet
SheetsHelper.prototype.createSpreadsheet = function(resource, callback) {

    var self = this;
    var request = { resource };
    self.service.spreadsheets.create(request, function(err, spreadsheet) {
      if (err) {
        return callback(err);
      }
      callback(null,spreadsheet);
    });
};

SheetsHelper.prototype.getSpreadSheet = function(spreadsheetId,callback)
{
  var self = this;
  var request = { 
    spreadsheetId:spreadsheetId
  };
  self.service.spreadsheets.get(request, function(err, spreadsheet) {
    if (err) {
      return callback(err);
    }
    callback(null,spreadsheet);
  });
}

SheetsHelper.prototype.updateSpreadSheet = function(spreadsheetId,title,callback)
{
  var self = this;
  var request = { 
    spreadsheetId:spreadsheetId,
    resource: {
      requests: [
        {
          updateSpreadsheetProperties:{
            properties:{title},
            fields:"*"
          }
        }
      ]
    }
  };
  self.service.spreadsheets.batchUpdate(request, function(err, spreadsheet) {
    if (err) {
      return callback(err);
    }
    callback(null,spreadsheet);
  });
}


SheetsHelper.prototype.updateSheetProperties = function(spreadsheetId,sheetId,title,rowCount,columnCount,frozenRowCount,frozenColumnCount,hideGridlines,callback)
{
  var self = this;
  var request = { 
    spreadsheetId:spreadsheetId,
    resource: {
      requests: [
        {
          updateSheetProperties:{
            properties:{
              sheetId,
              title,
              gridProperties:{
                rowCount,
                columnCount,
                frozenRowCount,
                frozenColumnCount,
                hideGridlines
              }
            },
            fields:"*"
          }
        }
      ]
    }
  };
  self.service.spreadsheets.batchUpdate(request, function(err, spreadsheet) {
    if (err) {
      return callback(err);
    }
    callback(null,spreadsheet);
  });
}

SheetsHelper.prototype.addSheet = function(spreadsheetId,title,rowCount,columnCount,frozenRowCount,frozenColumnCount,hideGridlines,callback)
{
  var self = this;
  var request = { 
    spreadsheetId:spreadsheetId,
    resource: {
      requests: [
        {
          addSheet:{
            properties:{
              title,
              gridProperties:{
                rowCount,
                columnCount,
                frozenRowCount,
                frozenColumnCount,
                hideGridlines
              }
            }
          }
        }
      ]
    }
  };
  self.service.spreadsheets.batchUpdate(request, function(err, spreadsheet) {
    if (err) {
      return callback(err);
    }
    callback(null,spreadsheet);
  });
}

SheetsHelper.prototype.deleteSheet = function(spreadsheetId,sheetId,callback)
{
  var self = this;
  var request = { 
    spreadsheetId:spreadsheetId,
    resource: {
      requests: [
        {
          deleteSheet:{
            sheetId
          }
        }
      ]
    }
  };
  self.service.spreadsheets.batchUpdate(request, function(err, spreadsheet) {
    if (err) {
      return callback(err);
    }
    callback(null,spreadsheet);
  });
}

SheetsHelper.prototype.updateCells = function(spreadsheetId,sheetId,rows,startRow,startColumn,callback)
{
  var self = this;
  var request = { 
    spreadsheetId:spreadsheetId,
    resource: {
      requests: [
        {
          updateCells:{
            rows,
            fields:'*',
            start:{
              sheetId,
              rowIndex:startRow,
              columnIndex:startColumn
            }
          }
        }
      ]
    }
  };
  self.service.spreadsheets.batchUpdate(request, function(err, spreadsheet) {
    if (err) {
      return callback(err);
    }
    callback(null,spreadsheet);
  });
}

SheetsHelper.prototype.addCells = function(spreadsheetId,sheetId,rows,callback)
{
  var self = this;
  var request = { 
    spreadsheetId:spreadsheetId,
    resource: {
      requests: [
        {
          appendCells:{
            sheetId,
            rows,
            fields:'*'
          }
        }
      ]
    }
  };
  self.service.spreadsheets.batchUpdate(request, function(err, spreadsheet) {
    if (err) {
      return callback(err);
    }
    callback(null,spreadsheet);
  });
}

SheetsHelper.prototype.deleteCells = function(spreadsheetId,sheetId,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,shiftDimension,callback)
{
  var self = this;
  var request = { 
    spreadsheetId:spreadsheetId,
    resource: {
      requests: [
        {
          deleteRange:{            
            range:{ sheetId, startRowIndex, startColumnIndex,endRowIndex,endColumnIndex },
            shiftDimension
          }
        }
      ]
    }
  };
  self.service.spreadsheets.batchUpdate(request, function(err, spreadsheet) {
    if (err) {
      return callback(err);
    }
    callback(null,spreadsheet);
  });
}


var COLUMNS = [
  { field: 'id', header: 'ID' },
  { field: 'customerName', header: 'Customer Name'},
  { field: 'productCode', header: 'Product Code' },
  { field: 'unitsOrdered', header: 'Units Ordered' },
  { field: 'unitPrice', header: 'Unit Price' },
  { field: 'status', header: 'Status'}
];
function buildHeaderRowRequest(sheetId) {
  var cells = COLUMNS.map(function(column) {
    return {
      userEnteredValue: {
        stringValue: column.header
      },
      userEnteredFormat: {
        textFormat: {
          bold: true
        }
      }
    }
  });
  return {
    updateCells: {
      start: {
        sheetId: sheetId,
        rowIndex: 0,
        columnIndex: 0
      },
      rows: [
        {
          values: cells
        }
      ],
      fields: 'userEnteredValue,userEnteredFormat.textFormat.bold'
    }
  };
}

SheetsHelper.prototype.sync = function(spreadsheetId, sheetId, orders, callback) {
    var requests = [];
    // Resize the sheet.
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId: sheetId,
          gridProperties: {
            rowCount: orders.length + 1,
            columnCount: COLUMNS.length
          }
        },
        fields: 'gridProperties(rowCount,columnCount)'
      }
    });
    // Set the cell values.
    requests.push({
      updateCells: {
        start: {
          sheetId: sheetId,
          rowIndex: 1,
          columnIndex: 0
        },
        rows: buildRowsForOrders(orders),
        fields: '*'
      }
    });
    // Send the batchUpdate request.
    var request = {
      spreadsheetId: spreadsheetId,
      resource: {
        requests: requests
      }
    };
    this.service.spreadsheets.batchUpdate(request, function(err) {
      if (err) {
        return callback(err);
      }
      return callback();
    });
};
function buildRowsForOrders(orders) {
  return orders.map(function(order) {
    var cells = COLUMNS.map(function(column) {
      switch (column.field) {
        case 'unitsOrdered':
          return {
            userEnteredValue: {
              numberValue: order.unitsOrdered
            },
            userEnteredFormat: {
              numberFormat: {
                type: 'NUMBER',
                pattern: '#,##0'
              }
            }
          };
          break;
        case 'unitPrice':
          return {
            userEnteredValue: {
              numberValue: order.unitPrice
            },
            userEnteredFormat: {
              numberFormat: {
                type: 'CURRENCY',
                pattern: '"$"#,##0.00'
              }
            }
          };
          break;
        case 'status':
          return {
            userEnteredValue: {
              stringValue: order.status
            },
            dataValidation: {
              condition: {
                type: 'ONE_OF_LIST',
                values: [
                  { userEnteredValue: 'PENDING' },
                  { userEnteredValue: 'SHIPPED' },
                  { userEnteredValue: 'DELIVERED' }
                ]
              },
              strict: true,
              showCustomUi: true
            }
          };
          break;
        default:
          return {
            userEnteredValue: {
              stringValue: order[column.field].toString()
            }
          };
      }
    });
    return {
      values: cells
    };
  });
}
function buildPivotTableRequest(sourceSheetId, targetSheetId) {
  return {
    updateCells: {
      start: { sheetId: targetSheetId, rowIndex: 0, columnIndex: 0 },
      rows: [
        {
          values: [
            {
              pivotTable: {
                source: {
                  sheetId: sourceSheetId,
                  startRowIndex: 0,
                  startColumnIndex: 0,
                  endColumnIndex: COLUMNS.length
                },
                rows: [
                  {
                    sourceColumnOffset: getColumnForField('productCode').index,
                    showTotals: false,
                    sortOrder: 'ASCENDING'
                  }
                ],
                values: [
                  {
                    summarizeFunction: 'SUM',
                    sourceColumnOffset: getColumnForField('unitsOrdered').index
                  },
                  {
                    summarizeFunction: 'SUM',
                    name: 'Revenue',
                    formula: util.format("='%s' * '%s'",
                        getColumnForField('unitsOrdered').header,
                        getColumnForField('unitPrice').header)
                  }
                ]
              }
            }
          ]
        }
      ],
      fields: '*'
    }
  };
}
  
function buildFormatPivotTableRequest(sheetId) {
  return {
    repeatCell: {
      range: { sheetId: sheetId, startRowIndex: 1, startColumnIndex: 2 },
      cell: {
        userEnteredFormat: {
          numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }
        }
      },
      fields: 'userEnteredFormat.numberFormat'
    }
  };
}
  
function buildAddChartRequest(sheetId) {
  return {
    addChart: {
      chart: {
        spec: {
          title: 'Revenue per Product',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'RIGHT_LEGEND',
            domains: [
              // Show a bar for each product code in the pivot table.
              {
                domain: { sourceRange: { sources: [{
                  sheetId: sheetId,
                  startRowIndex: 0,
                  startColumnIndex: 0,
                  endColumnIndex: 1
                }]}}
              }
            ],
            series: [
              // Set that bar's length based on the total revenue.
              {
                series: { sourceRange: { sources: [{
                  sheetId: sheetId,
                  startRowIndex: 0,
                  startColumnIndex: 2,
                  endColumnIndex: 3
                }]}}
              }
            ]
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: sheetId, rowIndex: 0, columnIndex: 3 },
            widthPixels: 600,
            heightPixels: 400
          }
        }
      }
    }
  };
}
  
function getColumnForField(field) {
  return COLUMNS.reduce(function(result, column, i) {
    if (column.field == field) {
      column.index = i;
      return column;
    }
    return result;
  });
}
module.exports = SheetsHelper;