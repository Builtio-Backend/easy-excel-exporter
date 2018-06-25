const chai       = require('chai');
const should     = chai.should();
const expect     = chai.expect();
const assert     = require("assert")
const fs         = require('fs');
var ExcelAdapter = require('./index');

describe('testing excel_adapter_module', function() {
  var excelAdapter;
  var columnArray = [];
  var excelSheetRows = [];

  beforeEach(function() { 
    excelAdapter = new ExcelAdapter({
      sheetName: 'my-new-sheet',
      autoCast  : false
    })
  })

  describe("Testing autoCast", function() {
    it("Testing exceladapter with autocast false", function() {
      var sheetName = "Test Sheet"
      var newExcelAdapter = new ExcelAdapter({
        sheetName : sheetName,
        autoCast  : false
      })
  
      var headers = [{
        "columnName" : "Name",
        "dataType"   : "string"
      }, {
        "columnName" : "Age",
        "dataType"   : "number"
      }, {
        "columnName" : "Active",
        "dataType"   : "Boolean"
      }, {
        "columnName" : "DOB",
        "dataType"   : "date"
      }]
  
      var objects = [{
        "Name" : "Test 1",
        "Age" : "25",
        "Active" : true,
        "DOB" : "1993-05-21T09:44:27.870Z"
      }, {
        "Name" : "Test 2",
        "Age" : 30,
        "Active" : "False",
        "DOB" : "1988-04-11T09:44:27.870Z"
      }, {
        "Name" : "Test 3",
        "Age" : "adkhdsdf",
        "Active" : "TRUE",
        "DOB" : "1989-06-10T09:44:27.870Z"
      }]
  
      return newExcelAdapter.createColumns(headers)
      .then(function() {
        return newExcelAdapter.addObjects(objects)
      })
      .then(function(lastRowIndex) {  
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["A2"].v, "string")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["B2"].v, "string")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["C2"].v, "boolean");
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["D2"].v, "string");
  
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["A3"].v, "string")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["B3"].v, "number")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["C3"].v, "string");
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["D3"].v, "string");
  
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["A4"].v, "string")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["B4"].v, "string")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["C4"].v, "string");
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["D4"].v, "string");
      })
    })
  
    it("Testing exceladapter with autocast true", function() {
      var sheetName = "Test Sheet"
      var newExcelAdapter = new ExcelAdapter({
        sheetName : sheetName,
        autoCast  : true
      })
  
      var headers = [{
        "columnName" : "Name",
        "dataType"   : "string"
      }, {
        "columnName" : "Age",
        "dataType"   : "number"
      }, {
        "columnName" : "Active",
        "dataType"   : "Boolean"
      }, {
        "columnName" : "DOB",
        "dataType"   : "date"
      }]
  
      var objects = [{
        "Name" : "Test 1",
        "Age" : "25",
        "Active" : true,
        "DOB" : "1993-05-21T09:44:27.870Z"
      }, {
        "Name" : "Test 2",
        "Age" : 30,
        "Active" : "False",
        "DOB" : "1988-04-11T09:44:27.870Z"
      }, {
        "Name" : "Test 3",
        "Age" : "adkhdsdf",
        "Active" : "TRUE",
        "DOB" : "1989-06-10T09:44:27.870Z"
      }]
  
      return newExcelAdapter.createColumns(headers)
      .then(function() {
        return newExcelAdapter.addObjects(objects)
      })
      .then(function(lastRowIndex) {  
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["A2"].v, "string")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["B2"].v, "number")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["C2"].v, "boolean");
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["D2"].v, "string");
  
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["A3"].v, "string")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["B3"].v, "number")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["C3"].v, "boolean");
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["D3"].v, "string");
  
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["A4"].v, "string")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["B4"].v, "number")
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["C4"].v, "boolean");
        assert.equal(typeof newExcelAdapter.workbook.Sheets[sheetName]["D4"].v, "string");
      })
    })

    it.skip("Should download file", function() {
      var sheetName = "Test Sheet"
      var newExcelAdapter = new ExcelAdapter({
        sheetName : sheetName,
        autoCast  : false
      })
  
      var headers = [{
        "columnName" : "Name",
        "dataType"   : "string"
      }, {
        "columnName" : "Age",
        "dataType"   : "number"
      }, {
        "columnName" : "Active",
        "dataType"   : "Boolean"
      }]
  
      var objects = [{
        "Name" : "Test 1",
        "Age" : "25",
        "Active" : true,
        "DOB" : "1993-05-21T09:44:27.870Z"
      }, {
        "Name" : "Test 2",
        "Age" : 30,
        "Active" : "False",
        "DOB" : "1988-04-11T09:44:27.870Z"
      }, {
        "Name" : "Test 3",
        "Age" : "adkhdsdf",
        "Active" : "TRUE"}]
  
      return newExcelAdapter.createColumns(headers)
      .then(function() {
        return newExcelAdapter.addObjects(objects)
      })
      .then(function(lastRowIndex) {  
        return newExcelAdapter.downloadFile()
      })
    })
  })

  it('should throw error if columnArray is blank', function(){
    return excelAdapter
      .createColumns(columnArray)
      .catch(function(err){
        assert.equal(err.errorKey, "noColumnsFound")
      });
  })

  it('should give error if columnArray is not an array', function(){
    columnArray = 'not-an-array';
    return excelAdapter
      .createColumns(columnArray)
      .then(function(result){
        console.log('error', result);
      })
      .catch(function(err){
        err.errorKey.should.be.eql('invalidColumnArray');
      })
  })

  it('should give error if in columnArray dataType & columnName is not found', function(){
    columnArray = [
      {columnName : 'Age', dataType : 'number'},
      {a : 'name', dataType:'object'}
    ];

    return excelAdapter
      .createColumns(columnArray)
      .then(function(result){
        console.log('error', result);
      })
      .catch(function(err){
        err.errorKey.should.be.eql('invalidColumnArray');
      })
  })

  it('should give error if in columnArray, columnName & dataType is not String', function(){
    columnArray = [
      {columnName : 'Name', dataType : 'number'},
      {columnName: 'Age', dataType: 'number'},
      {columnName: {name : 'das'}, dataType: 'string'}];

    return excelAdapter
      .createColumns(columnArray)
      .then(function(lastRowIndex){
        console.log('error', lastRowIndex);
      })
      .catch(function(err){
        err.errorKey.should.be.eql('invalidColumnArray')
      })
    
  })

  it('should add columnArray ', function(){
    columnArray = [
      {columnName:'Name', dataType : 'string'},
      {columnName:'Age', dataType : 'number'},
      {columnName:'Profile', dataType : 'object'}
    ]
    
    return excelAdapter
    .createColumns(columnArray)
    .then(function(){

    })
    .catch(function(err){
      throw err;
    })
  })

  it('should give error if excelsheet rows is not an array', function(){
    var excelSheetRows = {name : 'abc', age :21}
    var columnArray = [{columnName:'name', dataType:'string'}]
    
    return excelAdapter
      .createColumns(columnArray)
      .then(function(lastRowIndex){
        return excelAdapter.addObjects(excelSheetRows, lastRowIndex)
      })
      .then(function(result){
        console.log('error', result);
      })
      .catch(function(err){
        err.errorKey.should.be.eql('invalidExcelSheetRows');
      }) 
  })

  it('should give error on adding excel rows if null columnArray', function(){
    var excelSheetRows = [
      {name : 'abc', age :21},
      {name : 'bcd', age: 23}
    ];

    columnArray = [
      {columnName:'Name', dataType : 'string'},
      {columnName:'Age', dataType : 'number'},
      {columnName:'Profile', dataType : 'object'}
    ]
    
    return excelAdapter
    .createColumns(columnArray)
    .then(function(lastRowIndex){
      excelAdapter.columnArray = [];
      return excelAdapter.addObjects(excelSheetRows)
    })
    .then(function(result){
      console.log('error', result);
    })
    .catch(function(err){
      err.errorKey.should.be.eql('noColumnsFound');
    }) 
  })

  it('should give error on adding excel rows if invalid columnArray ', function(){
    var excelSheetRows = [
      {name : 'abc', age :21},
      {name : 'bcd', age: 23}
    ];

    // invalid type column Array
    excelAdapter
      .createColumns(columnArray)
      .then(function (lastRowIndex) {
        excelAdapter.columnArray = [{ keys: 'abcd' }]
        return excelAdapter.addObjects(excelSheetRows)
      })
      .then(function (result) {
        console.log('error', result);
      })
      .catch(function (err) {
        err.errorKey.should.be.eql('invalidColumnArray');
      }) 

  })

  it('should update the workbook object and return promise', function(){
    var columnArray    = [
      {columnName:"name", dataType:'String'},
      {columnName:"age", dataType:'number'}
    ];
    
    var excelSheetRows = [
      {name : 'abc', age :21},
      {name : 'bcd', age: 23}
    ];
    
    return excelAdapter
    .createColumns(columnArray)
      .then(function(){
        return excelAdapter.addObjects(excelSheetRows)
      })
      .then(function(lastRowIndex){
        lastRowIndex.should.be.a('number');
      })
      .catch(function(err){
        throw err;
      })

  })

  it('should return error on adding rows if workbook object is changed to null', function(){
    var columnArray    = [
      {columnName:"name", dataType:'String'},
      {columnName:"age", dataType:'number'}
    ];
    
    var excelSheetRows = [
      {name : 'abc', age :21},
      {name : 'bcd', age: 23}
    ];
    
    return excelAdapter
      .createColumns(columnArray)
      .then(function(){
        excelAdapter.workbook = null        
        return excelAdapter.addObjects(excelSheetRows)
      })
      .then(function(lastRowIndex){
        throw false;
      })
      .catch(function(err){
        err.errorKey.should.be.eql('noWorkbookObjectFound');
      })
  })

  it('should download file if valid workbook object', function(){
    columnArray    = [
      {columnName:"name", dataType:'String'},
      {columnName:"age", dataType:'number'}
    ];
    
    var excelSheetRows = [
      {name : 'abc', age :21},
      {name : 'bcd', age: 23}
    ];
    
    return excelAdapter
      .createColumns(columnArray)
      .then(function(){  
        return excelAdapter.addObjects(excelSheetRows)
      })
      .then(function(lastRowIndex){
        return excelAdapter.downloadFile();
      })
      .then(function(res){
        res.should.not.be.empty;

        setTimeout(function(){
          res.pipe(fs.createWriteStream("test.txt"))
        }, 3000);
      })
      .catch(function(err){
        console.log('error', err);
        throw err;
      })
  })

  it('should give error if null workbook object is used to download file', function(){
    var columnArray    = [
      {columnName:"name", dataType:'String'},
      {columnName:"age", dataType:'number'}
    ];
    
    var excelSheetRows = [
      {name : 'abc', age :21},
      {name : 'bcd', age: 23}
    ];

    return excelAdapter
      .createColumns(columnArray)
      .then(function(){    
        return excelAdapter.addObjects(excelSheetRows)
      })
      .then(function(lastRowIndex){
        excelAdapter.workbook = null
        return excelAdapter.downloadFile();
      })
      .catch(function(err){
        err.errorKey.should.be.eql('noWorkbookObjectFound');  
      })
  })

  it('should give error if invalid workbook object is used to download file', function(){
    var columnArray    = [
      {columnName:"name", dataType:'String'},
      {columnName:"age", dataType:'number'}
    ];
    
    var excelSheetRows = [
      {name : 'abc', age :21},
      {name : 'bcd', age: 23}
    ];

    return excelAdapter
      .createColumns(columnArray)
      .then(function(){    
        return excelAdapter.addObjects(excelSheetRows)
      })
      .then(function(lastRowIndex){
        excelAdapter.workbook = {bac : 'sdas'};
        return excelAdapter.downloadFile();
      })
      .catch(function(err){
        err.errorKey.should.be.eql('invalidWorkBook');  
      })
  })
})