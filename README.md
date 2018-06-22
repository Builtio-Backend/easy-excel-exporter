# About
This nodejs module can be used to create an excel spread-sheet with set of header row and data rows.

The headers can have a particular data-type and value.

Each column will accomodate the data of particular data-type specified in the header row.

# Initialize
```
var excelAdapter = require('excel_adapter');
var excel = new excelAdapter(options);
```
option is an object used to create instance of excel_adapter.

```
options : {
  sheetName : 'this-is-an-optional-field',
  fileName  : 'this-is-an-optional-field',
}
```

# Methods
The module has three methods.

## createColumns
To add header row
```
excel.createColumns(headerRow);
```
headerRow will be array of objects with name of every column and dataType for that column. eg:

```
var headerRow = [
  {columnName : Name, dataType : string},
  {columnName : Age, dataType : number},
  {columnName : Profile, dataType : Object},
]
```

It will return the row index of the header row i.e 1.

## addObjects
To add header row
```
excel.addObjects(dataRows, lastRowIndex);
```

dataRows will be array of objects to fill the data in the spreadsheet.

```
dataRows = [
  {Name : 'abc', Age : 22, Profile : {}},
  {Name : 'bcd', Age : 21},
]
```
It will return the last row index of the data added in the spreadsheet

## downloadFile
This will save the file on the system & returns the readStream for the file. Once the file is read it unlinks the file from the system.

```
excel.downloadFile();
```