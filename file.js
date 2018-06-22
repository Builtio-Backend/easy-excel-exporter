var localAdapter = require('./index.js');

var adapter = localAdapter({sheetName : 'okay', autocast : true});

var columnArray = [
  {
    columnName : 'name',
    dataType   : 'String'
  },
  {
    columnName : 'age',
    dataType   : 'Number'
  },
  {
    columnName : 'profile',
    dataType   : 'Object'
  },
  {
    columnName : 'active',
    dataType   : 'Boolean'
  },
]


var data = [
  {
    name : 'aman', age:22, profile:{image :''}, active:"adsasd"
  }
]

adapter.createColumns(columnArray)
  .then(function(){
    return adapter.addObjects(data);
  })
  .then(function(){
    adapter.downloadFile();
  })