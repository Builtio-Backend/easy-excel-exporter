# About
This NodeJS module can be used to create an Excel spreadsheets with set of header row and data rows.

These cells can have a particular data-type and value.

Each column will accomodate the data of particular data-type specified in the header row.

NOTE: ExcelExporter only supports these below mentioned datatypes:<br><b>object, string, number, boolean, date.</b>

# Initialize
```
var ExcelExporter = require('easy-excel-exporter');
var excelAdaptor  = ExcelExporter(options);
```
options is an object used to create instance of excel-exporter.

```
options : {
  sheetName : 'this-is-an-optional-field',
  fileName  : 'this-is-an-optional-field',
  autoCast  : 'this-is-an-optional-field'
}
```

# Methods
The module has three methods.

## createColumns
To add header row
```
excelAdaptor.createColumns(headerRow);
```
headerRow will be array of objects with name of every column and dataType for that column. <br>E.g.:

```
var headerRow = [
  {columnName : Name, dataType : string},
  {columnName : Age, dataType : number},
  {columnName : Profile, dataType : Object},
]
```

## addObjects
To add header row
```
excelAdaptor.addObjects(dataRows);
```

dataRows will be array of objects to fill the data in the spreadsheet.

```
dataRows = [
  {Name : "abc", Age : 22, Profile : {}},
  {Name : "bcd", Age : 21},
]
```
It will return the last row index of the data added in the spreadsheet

## downloadFile
This will save the file on the system & return a stream of the file. Once the file is read, this temporary file is erased from your system.

```
excelAdaptor.downloadFile();
```