# About
This NodeJS module can be used to create Excel spreadsheets.
Every cell has its own datatype associated with it, and a value assigned for that cell.

# Initialization
```
var EasyExcelExporter = require('easy-excel-exporter');
var excelAdapter      = EasyExcelExporter(options);
```
options will be used to re-initialize default option values used to create Excel Spreadsheet.

```
options : {
  sheetName : 'sheet-name', // String value for assigning your own Sheet Name.
  fileName  : 'test-file', // String value for assigning a name for the Excel file created.
  autoCast  : true // Boolean value that will indicate whether to type cast values for cells or not(Default : false).
  path : '/<<file-path>>/' // String value to define your own storage path where the excel file will be saved.
}
```

# Methods
easy-excel-exporter provides three methods. All these return a promise.

## createColumns(excelHeaders)
This function creates a row in your Excel spreadsheet which contains the values of the column names as specified in excelHeaders.
```
excelAdapter.createColumns(excelHeaders);
```
excelHeaders is an array of objects with name of column specified as <b>"columnName"</b> and its associated dataType specified as <b>"dataType"</b>.<br>E.g.:


```
var excelHeaders = [{
  columnName: 'Name',
  dataType: EasyExcelExporter.dataType.String
}, {
  columnName: 'Age',
  dataType: EasyExcelExporter.dataType.Number
}, {
  columnName: 'Profile',
  dataType: EasyExcelExporter.dataType.Object
}]
```

### supported datatypes
easy-excel-exporter provides static method <b>EasyExcelExporter.dataType</b> and supports these below mentioned types:<br><b>object, string, number, boolean, date.</b>
```
EasyExcelExporter.dataType.String  // for string dataType
EasyExcelExporter.dataType.Object  // for object dataType
EasyExcelExporter.dataType.Boolean // for boolean dataType
EasyExcelExporter.dataType.Number  // for number dataType
EasyExcelExporter.dataType.Date    // for date dataType
```

Once your columns have been set in the Excel spreadsheet, it will return a Promise.

## addObjects(rows)
This function iterates through rows, which is an array of objects provided as an argument. Each object will be treated as a row in the Excel Spreadsheet.
```
rows = [{
  Name : "abc",
  Age : 22,
  Profile : {}
}]

excelAdapter.addObjects(rows);
```
This function will return an index of the last row created in the Excel Spreadsheet.

<b>NOTE:</b> Objects are always stringified when they are added to cells.

## downloadFile()
This function will return a downloadable stream of the Excel spreadsheet, created at the default storage path or, the path specified in options while creating an instance of easy-excel-exporter.

```
excelAdapter.downloadFile();
```

# Features

## autocast option
If you set autocast option as <b>true</b> while creating easy-excel-exporter instance, value for that cell will be typecasted to the dataType of the column mentioned while creating spreadsheet columns.<br><br>If typecast fails, then the cell will contain a <b>null</b> value.
<br><br>
Default value for autocast is <b>false</b> which means that the <b>dataType</b> provided while creating columns will be ignored when Excel cell is being added for that corresponding column. 

