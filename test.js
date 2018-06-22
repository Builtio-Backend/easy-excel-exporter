const chai = require('chai');
const should = chai.should();
const expect = chai.expect();
const fs = require('fs');
var ExcelAdapter = require('./index');

describe('testing excel_adapter_module', function() {
  var excelAdapter;
  var columnArray = [];
  var excelSheetRows = [];

  beforeEach(function() {
    excelAdapter = new ExcelAdapter({sheetName: 'my-new-sheet'})
  })

  it('should give error if columnArray is blank', function(){

    return excelAdapter
      .createColumns(columnArray)
      .catch(function(err){
        err.errorKey.should.be.eql('noColumnsFound')
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

  // removed test for blank array, allowing excelsheet rows to be null
  // it('should give error if excelSheetRows is blank', function(){
  //   var excelSheetRows = []
  //   var columnArray = [{columnName:'name', dataType:'string'}]
    
  //   return excelAdapter
  //     .createColumns(columnArray)
  //     .then(function(lastRowIndex){
  //       return excelAdapter.addObjects(excelSheetRows, lastRowIndex)
  //     })
  //     .then(function(result){
  //       console.log('error', result);
  //     })
  //     .catch(function(err){
  //       console.log(err);
  //       err.errorKey.should.be.eql('invalidExcelSheetRows');
  //     })

  // })

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
      .then(function(lastRowIndex){
        excelAdapter.columnArray = [{keys:'abcd'}]
        return excelAdapter.addObjects(excelSheetRows)
      })
      .then(function(result){
        console.log('error', result);
      })
      .catch(function(err){
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