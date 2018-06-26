const chai          = require('chai')
const should        = chai.should()
const expect        = chai.expect()
const assert        = require("assert")
const fs            = require('fs')
const ExcelAdapter  = require('./index')
const R             = require('ramda')
const errorMessages = require("./error-messages.json")

describe('Testing Excel Adaptor', function() {
  var excelAdapter
  var columnArray    = [];
  var excelSheetRows = [];

  beforeEach(function() {
    excelAdapter = ExcelAdapter({ 
      sheetName: 'my-new-sheet',
      autoCast  : false
    })
  })

  it("checks for options provided while creating Excel Adaptor instance", function() {
    assert.equal(excelAdapter.excelSheetName, "my-new-sheet")
    assert.equal(excelAdapter.autoCast, false)
  })

  describe("Testing createColumn function", function() {
    it('should give error if columnArray is not an array', function() {
      columnArray = null;
      return excelAdapter
        .createColumns(columnArray)
        .catch(function(err){
          assert.equal(err.message, getErrorMessage("invalidColumnArray"))
        })
    })
  
    it('should throw error if columnArray is blank', function(){
      columnArray = []
      return excelAdapter
        .createColumns(columnArray)
        .catch(function(err){
          assert.equal(err.message, getErrorMessage("noColumnsFound"))
        });
    })
  
    it('should give error if in columnArray dataType & columnName is not found', function(){
      columnArray = [
        {columnName : 'Age', dataType : 'number'},
        {a : 'name', dataType:'object'}
      ];
  
      return excelAdapter
        .createColumns(columnArray)
        .catch(function(err){
          assert.equal(err.message, getErrorMessage("invalidColumnArray"))
        })
    })
  
    it('should give error if in columnArray, columnName & dataType is not String', function(){
      columnArray = [
        {columnName : 'Name', dataType : 'number'},
        {columnName: 'Age', dataType: 'number'},
        {columnName: {name : 'das'}, dataType: 'string'}];
  
      return excelAdapter
        .createColumns(columnArray)
        .catch(function(err){
          assert.equal(err.message, getErrorMessage("invalidColumnArray"))
        })
    })
  
    it("throws error when invalid datatype is given in columnsArray", function() {
      columnArray = [{
        "columnName" : "Name",
        "dataType"   : "invalid-type"
      }]
  
      return excelAdapter.createColumns(columnArray)
      .catch(err => {
        assert.equal(err.message, getErrorMessage("unsupportedDataType"))
      })
    })
  
    it('adds array of columns to excel adapter object', function() {
      columnArray = [{
        columnName: 'Name',
        dataType: 'string'
      }, { 
        columnName: 'Age',
        dataType: 'number'
      }, { 
        columnName: 'Profile',
        dataType: 'object'
      }]
      
      return excelAdapter
      .createColumns(columnArray)
      .then(function() {
        var sheetName = excelAdapter.excelSheetName
        assert.equal(R.is(Array, excelAdapter.columnArray), true)
        assert.equal(excelAdapter.columnArray.length, 3)
        assert.equal(excelAdapter.workbook.Sheets[sheetName]["A1"].v, columnArray[0].columnName)
        assert.equal(excelAdapter.workbook.Sheets[sheetName]["B1"].v, columnArray[1].columnName)
        assert.equal(excelAdapter.workbook.Sheets[sheetName]["C1"].v, columnArray[2].columnName)
      })
    })
  })

  describe("Testing autoCast option in Excel Adapter", function() {
    it("Testing exceladapter with autocast false", function() {
      var sheetName = "Test Sheet"
      var newExcelAdapter = ExcelAdapter({
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
        assert.equal(R.is(String, newExcelAdapter.workbook.Sheets[sheetName]["A2"].v), true)
        assert.equal(R.is(Number, newExcelAdapter.workbook.Sheets[sheetName]["B2"].v), false)
        assert.equal(R.is(Boolean, newExcelAdapter.workbook.Sheets[sheetName]["C2"].v), true)
        assert.equal(R.is(Date, newExcelAdapter.workbook.Sheets[sheetName]["D2"].v), false)
  
        assert.equal(R.is(String, newExcelAdapter.workbook.Sheets[sheetName]["A3"].v), true)
        assert.equal(R.is(Number, newExcelAdapter.workbook.Sheets[sheetName]["B3"].v), true)
        assert.equal(R.is(Boolean, newExcelAdapter.workbook.Sheets[sheetName]["C3"].v), false)
        assert.equal(R.is(Date, newExcelAdapter.workbook.Sheets[sheetName]["D3"].v), false)

        assert.equal(R.is(String, newExcelAdapter.workbook.Sheets[sheetName]["A4"].v), true)
        assert.equal(R.is(Number, newExcelAdapter.workbook.Sheets[sheetName]["B4"].v), false)
        assert.equal(R.is(Boolean, newExcelAdapter.workbook.Sheets[sheetName]["C4"].v), false)
        assert.equal(R.is(Date, newExcelAdapter.workbook.Sheets[sheetName]["D4"].v), false)
      })
    })
  
    it("Testing exceladapter with autocast true", function() {
      var sheetName = "Test Sheet"
      var newExcelAdapter = ExcelAdapter({
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
        "DOB" : new Date("1993-05-21T09:44:27.870Z")
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
        assert.equal(R.is(String, newExcelAdapter.workbook.Sheets[sheetName]["A2"].v), true)
        assert.equal(R.is(Number, newExcelAdapter.workbook.Sheets[sheetName]["B2"].v), true)
        assert.equal(R.is(Boolean, newExcelAdapter.workbook.Sheets[sheetName]["C2"].v), true)
        assert.equal(R.is(Date, newExcelAdapter.workbook.Sheets[sheetName]["D2"].v), true)
  
        assert.equal(R.is(String, newExcelAdapter.workbook.Sheets[sheetName]["A3"].v), true)
        assert.equal(R.is(Number, newExcelAdapter.workbook.Sheets[sheetName]["B3"].v), true)
        assert.equal(R.is(Boolean, newExcelAdapter.workbook.Sheets[sheetName]["C3"].v), true)
        assert.equal(R.is(Date, newExcelAdapter.workbook.Sheets[sheetName]["D3"].v), true)

        assert.equal(R.is(String, newExcelAdapter.workbook.Sheets[sheetName]["A4"].v), true)
        assert.equal(R.is(Number, newExcelAdapter.workbook.Sheets[sheetName]["B4"].v), true)
        assert.equal(R.is(Boolean, newExcelAdapter.workbook.Sheets[sheetName]["C4"].v), true)
        assert.equal(R.is(Date, newExcelAdapter.workbook.Sheets[sheetName]["D4"].v), true)
      })
    })

    it.skip("Should download file", function() {
      var sheetName = "Test Sheet"
      var newExcelAdapter = ExcelAdapter({
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
      }, {
        "columnName" : "City",
        "dataType"   : "object"
      }]
  
      var objects = [{
        "Name" : "Test 1",
        "Age" : "25",
        "Active" : true,
        "DOB" : "1993-05-21T09:44:27.870Z",
        "City" : {
          "Address" : "Test Address"
        }
      }, {
        "Name" : "Test 2",
        "Age" : 30,
        "Active" : "False",
        "DOB" : "1988-04-11T09:44:27.870Z",
        "City" : {
          "Address" : "Test Address 2"
        }
      }, {
        "Name" : "Test 3",
        "Age" : "adkhdsdf",
        "Active" : "TRUE",
        "DOB" : "1990-04-11T09:44:27.870Z",
        "City" : "Test Address 3"
      }]
  
      return newExcelAdapter.createColumns(headers)
      .then(function() {
        return newExcelAdapter.addObjects(objects)
      })
      .then(function(lastRowIndex) {  
        return newExcelAdapter.downloadFile()
      })
    })
  })

  describe("Testing addObjects function", function() {
    it("throws error when lastRowIndex is invalid", function() {
      var columnArray    = [{
        columnName       : 'name',
        dataType         : 'string'
      }]
      
      return excelAdapter
        .createColumns(columnArray)
        .then(function(lastRowIndex){
          excelAdapter.lastRowIndex = null
          return excelAdapter.addObjects([], lastRowIndex)
        })
        .catch(function(err){
          assert.equal(err.message, getErrorMessage("noLastRowIndex"))
        }) 
    })
  
    it('should return error on adding rows if workbook object is modified', function(){
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
        .catch(function(err){
          assert.equal(err.message, getErrorMessage("noWorkbookObjectFound"))
        })
    })
  
    it('should give error if excelsheet rows is not an array', function(){
      var excelSheetRows = {}
      var columnArray    = [{
        columnName       : 'name',
        dataType         : 'string'
      }]
      
      return excelAdapter
        .createColumns(columnArray)
        .then(function(lastRowIndex){
          return excelAdapter.addObjects(excelSheetRows, lastRowIndex)
        })
        .catch(function(err){
          assert.equal(err.message, getErrorMessage("invalidExcelSheetRows"))
        }) 
    })
  
    it('should give error on adding excel rows if columnArray is null', function() {
      var excelSheetRows = [
        {name : 'abc', age :21},
        {name : 'bcd', age: 23}
      ]
  
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
      .catch(function(err){
        assert.equal(err.message, getErrorMessage("noColumnsFound"))
      }) 
    })
  
    it('should give error on adding excel rows if invalid columnArray ', function(){
      var excelSheetRows = [
        {name : 'abc', age :21},
        {name : 'def', age: 23}
      ];

      columnArray = []
  
      // invalid type column Array
      return excelAdapter
        .createColumns(columnArray)
        .then(function (lastRowIndex) {
          excelAdapter.columnArray = [{"sfsdfsdfs": "columnName"}]
          return excelAdapter.addObjects(excelSheetRows)
        })
        .catch(function (err) {
          assert.equal(err.message, getErrorMessage("noColumnsFound"))
        })
    })
  
    it('should adds objects to workbook and return last row index', function(){
      var columnArray = [{
        columnName: "name",
        dataType: 'String'
      }, {
        columnName: "age",
        dataType: 'number'
      }];
      
      var excelSheetRows = [{
        name: 'Name 1',
        age : 21
      }, {
        name: 'Name 2',
        age : 23
      }];
      
      return excelAdapter
      .createColumns(columnArray)
        .then(function(){
          return excelAdapter.addObjects(excelSheetRows)
        })
        .then(function(lastRowIndex){
          assert.equal(R.is(Number, lastRowIndex), true)
          assert.equal(lastRowIndex, 3)
          assert.equal(excelAdapter.workbook.Sheets[excelAdapter.excelSheetName]["A2"].v, excelSheetRows[0].name)
          assert.equal(excelAdapter.workbook.Sheets[excelAdapter.excelSheetName]["B2"].v, excelSheetRows[0].age)
  
          assert.equal(excelAdapter.workbook.Sheets[excelAdapter.excelSheetName]["A3"].v, excelSheetRows[1].name)
          assert.equal(excelAdapter.workbook.Sheets[excelAdapter.excelSheetName]["B3"].v, excelSheetRows[1].age)
        })
        .catch(function(err){
          throw err;
        })
    })
  })

  describe("Testing downloadFile function", function(){
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
          res.should.not.be.empty
          var newExcelFile = __dirname + "/test.xlsx"
          res.pipe(fs.createWriteStream(newExcelFile))
        })
    })
  
    it('should give error when workbook object is null and download file is called', function() {
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
          assert.equal(err.message, getErrorMessage("noWorkbookObjectFound"))
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
        .then(function() {    
          return excelAdapter.addObjects(excelSheetRows)
        })
        .then(function(lastRowIndex) {
          excelAdapter.workbook = {bac : 'sdas'};
          return excelAdapter.downloadFile();
        })
        .catch(function(err){
          assert.equal(err.message, getErrorMessage("invalidWorkBook"))
        })
    })
  })  
})

function getErrorMessage(errorKey) {
  return errorMessages[errorKey]
}