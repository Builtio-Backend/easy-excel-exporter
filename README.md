# About
This NodeJS module can be used to create an Excel spreadsheets with set of header row and data rows.

These cells can have a particular data-type and value.

Each column will accomodate the data of particular data-type specified in the header row.

NOTE: ExcelExporter only supports these below mentioned datatypes:<br><b>object, string, number, boolean, date.</b>

# Initialize
```
var EasyExcelExporter = require('easy-excel-exporter');
var easyExcelExporter = EasyExcelExporter(options);
```
options is an object used to create instance of excel-exporter.

```
options : {
  sheetName : 'sheet1', //optional
  fileName  : 'file1', //optional
  autoCast  : true //optional
  path : '/home/work' //optional
}
```

# Methods
The module has three methods.

## createColumns
To add header row
```
easyExcelExporter.createColumns(headerRow);
```
HeaderRow will be array of objects with name of every column and dataType for that column. <br>E.g.:

```
var headerRow = [
  {columnName : 'Name', dataType : 'string'},
  {columnName : 'Age', dataType : 'number'},
  {columnName : 'Profile', dataType : 'object'},
]
```

createColumns method returns promise.

## addObjects
To add data rows
```
easyExcelExporter.addObjects(dataRows);
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
This will save the file on the system & returns a stream of the file. Once the file is read, this temporary file is erased from your system.

```
easyExcelExporter.downloadFile();
```

# Features

## autocast
<b>when autocast is set to false</b>, data is written directly in the spreadsheet. Typecasting of data is not involved in the process.

For any cell if data is object then it is converted to string and then saved in spreadsheet.

<b>when autocast is set to true</b>, data is typecasted as per the dataType of the column, mentioned in the header row.

If the data is possible to be typecasted, then value is typecasted and inserted into cell else cell is set to null.    