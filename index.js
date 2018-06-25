const R                = require('ramda')
const xlsx             = require('xlsx')
const fs               = require('fs')
const randomCharString = 'abcdefghijklmnopqrstuvwxyz0123456789'
const allowedDataTypes = ['object', 'boolean', 'number', 'string', 'date']
const columnsList      = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

var LocalExcelAdaptor = R.curry(function(options){
  this.workbook          = null
  this.columnArray       = []
  this.lastRowIndex      = 1
  
  this.excelSheetName    = R.isEmpty(sanitizeString(options.sheetName)) ? 'Sheet 1' : sanitizeString(options.sheetName)

  this.excelFileName     = sanitizeString(options.fileName)
  this.storagePath       = sanitizeString(options.path)

  // Check for boolean type in autoCast
  this.autoCast          = options.autoCast ? true : false
})

/**
 * @param {Array} columnArray - array of objects with columnName & dataType
 * @example [{columnName:email, dataType:String}]
 * @returns {number} rowIndex of column added
*/
LocalExcelAdaptor.prototype.createColumns = function(columnArray) {
  // Rename columnArray to excelColumns
  var that = this
  return new Promise(function(resolve, reject){
    // validate the columnArray.
    validateColumnsArray(columnArray)

    that.columnArray = columnArray

    // array of columns to be created inside excel spreadsheet
    var excelSheetColumns = []

    // Sets width for each cell inside an excel spreadsheet
    that.columnArray.forEach(function() {
      excelSheetColumns.push({
        wch : 25
      })
    })

    var lastCellRef = getCellName(that.columnArray.length - 1, 1)

    that.workbook = {
      SheetNames: [that.excelSheetName],
      Sheets : {
        [that.excelSheetName] : {
          '!ref' : 'A1: ' + lastCellRef,
          '!cols': excelSheetColumns  
        }
      }
    }

    for (var columnIndex = 0 ; columnIndex < that.columnArray.length ; columnIndex++) {
      that.workbook.Sheets[that.excelSheetName][getCellName(columnIndex, 1)] = {
        t: "s",
        v: that.columnArray[columnIndex].columnName
      }
    }

    resolve()
  })
}

/**
 * @param {Array} excelSheetRows - array of objects to be inserted into the excel spreadsheet. 
 * @returns {number} rowIndex of last inserted object.
*/
LocalExcelAdaptor.prototype.addObjects = function(excelSheetRows) {
  var that = this
  return new Promise(function(resolve, reject){
    // validate workbook instance
    validateWorkBook(that.workbook)

    // validate the columnArray.
    validateColumnsArray(that.columnArray)

    // validate the excelSheetRows.
    validateRowsArray(excelSheetRows)

    // validate the lastRowIndex
    validatelastRowIndex(that.lastRowIndex)

    // increase the ref rows in workbook sheet if the number of rows increase in excelSheetRows
    var lastCellRef = getCellName(that.columnArray.length - 1 , that.lastRowIndex + excelSheetRows.length)

    // update the workbook ref rows. 
    that.workbook.Sheets[that.excelSheetName]['!ref'] = 'A1:' + lastCellRef
    
    // creates excel rows in workbook from data in excelSheetRows 
    excelSheetRows.forEach(function(excelSheetRow) {
      var excelValue = null
      var dataType   = null

      that.lastRowIndex = that.lastRowIndex + 1 // MOVE ON TO NEXT CELL IN SPREADSHEET TO ADD DATA

      for(var columnIndex = 0 ; columnIndex < that.columnArray.length ; columnIndex++) {
        excelValue = excelSheetRow[that.columnArray[columnIndex].columnName]
        dataType   = that.columnArray[columnIndex].dataType.toLowerCase()

        if (R.isNil(excelValue)){
          continue
        }

        if (that.autoCast) {
          excelValue = autoCast(excelValue, dataType)
        }
        
        // if (typeof excelValue !== dataType){
        //   console.log(excelValue, typeof excelValue, dataType)
        //   continue
        // }

        if(R.is(Object, excelValue)){
          excelValue = JSON.stringify(excelValue)
        }

        that.workbook.Sheets[that.excelSheetName][getCellName(columnIndex, that.lastRowIndex )] = {
          t: checkExcelValueDatatype(excelValue),
          v: excelValue
        }
      }
    })

    resolve(that.lastRowIndex)
  })
}

/**
 * writes the file in directory, reject errors if any
*/
LocalExcelAdaptor.prototype.downloadFile = function() {
  var that = this
  return new Promise(function(resolve, reject){
    // validate workbook object
    validateWorkBook(that.workbook)

    var fileName      = R.isEmpty(that.excelFileName) ? generateRandomFileName(10, randomCharString): that.excelFileName
    var storagePath   = R.isEmpty(that.storagePath) ? __dirname : that.storagePath
    var excelFilePath = storagePath + '/' + fileName + '.xlsx'

    console.log("Excel file path : ", excelFilePath)

    try {
      xlsx.writeFile(that.workbook, excelFilePath)

      var excelFileStream = fs.createReadStream(excelFilePath)

      excelFileStream.on("end", function(data) {
				fs.unlinkSync(excelFilePath)
			})
	
			excelFileStream.on("error", function(err){
				fs.unlinkSync(excelFilePath)
      })
      
			resolve(excelFileStream)
    } catch(err) {
      reject(generateErrorMessage('downloadFailed'))
    }
  })
}

/**
 * autoCast the value according to dataType
 */
function autoCast(excelValue, dataType){
  switch(dataType) {
    case "string": 
      excelValue = stringValidator(excelValue)
      break;

    case "boolean":
      excelValue = booleanValidator(excelValue)
      break

    case "number":
      excelValue = numberValidator(excelValue)
      break

    // case "date":
    //   console.log("Date datatype found")
    //   excelValue = dateValidator(excelValue)
    //   break;
  }

  return excelValue
}

function stringValidator(value) {
  if(R.is(String, value)) {
    return value
  }

  try {
    value = value.toString()
  } catch(err) {
    value = null
  }

  return value
}

function numberValidator(value) {
  if(R.is(Number, value)) {
    return value
  }

  try {
    value = parseInt(value)
  } catch(err) {
    value = null
  }

  return value
}

function booleanValidator(value) {
  if(R.is(Boolean, value)) {
    return value
  }

  if (value.toLowerCase() === 'true') {
    value = true
  } else if (value.toLowerCase() === 'false') {
    value = false
  }

  return value
}

function dateValidator(value) {
  if(R.is(Date, value)) {
    return value
  }

  try {
    value = new Date(value)
  } catch(err) {
    value = null
  }

  return value
}

/** 
 @param {*} excelValue - excelValue to be inserted inside the cell
 @param {*} column - column object eg {dataType : string, columnName : name}
 @returns {String} - datatype for the particular cell of excelSheet
*/
function checkExcelValueDatatype (excelValue) {
  if(R.isNil(excelValue)) {
    return 'z'
  } else if (typeof excelValue === 'boolean') {
    return 'b'
  } else if (typeof excelValue === 'number') {
    return 'n'
  } else {
    return 's'
  }
}

/**
 * Generate appropriate error message incase of any errors.
 * @param {string} errorMessage - Message that needs to be shown to the user.
 * @return {object} - with errorKey, errorData and errors
*/
function generateErrorMessage(errorMessage) {
  return {
    errorKey : errorMessage,
    errorData: {},
    errors   : []
  }
}

/**
  @param {Number} columnIndex - column index in excel spreadsheet to save data.
	@param {Number} rowIndex   - Row number to generated a cell name to store data from klassData.
	@return {String} A proper cell name consisting of column name and row number
  @example : If column index is 0 and row number is 2, generated cellname would be : A2.
*/
function getCellName(columnIndex, rowIndex) {
	var cellName = ""
	if (columnIndex > 25) {
    // logic to increase the column name in case of number of columns > 25
    // eg : AA1, AA2
    var quotient = Math.floor(columnIndex/26) - 1
    var remainder = columnIndex % 26
    cellName = columnsList[quotient] + "" + columnsList[remainder] + "" + rowIndex
	} else {
		cellName = columnsList[columnIndex] + "" + rowIndex
  }
  
	return cellName
}

/**
	* validate columnsArray.
  * @param {Array} columnArray - columnArray to be validated.
  * throws error if any.
*/
function validateColumnsArray(columnArray) {
  if (!R.is(Array, columnArray)) {
    throw generateErrorMessage('invalidColumnArray')
  }

  if (columnArray.length === 0) {
    throw generateErrorMessage('noColumnsFound')
  }

  columnArray.forEach(function(column){
    if(!column.hasOwnProperty('columnName') || !column.hasOwnProperty('dataType')){
      throw generateErrorMessage('invalidColumnArray')
    }
    // columnName & dataType fields must be in Strinf format
    if (typeof column['columnName'] !== 'string' || typeof column['dataType'] !== 'string' ){
      throw generateErrorMessage('invalidColumnArray')
    }
    //  column['dataType'] must be from allowed dataTypes only
    if(allowedDataTypes.indexOf(column['dataType'].toLowerCase()) == -1){
      throw generateErrorMessage('unsupportedDataType')
    }
  })
}

/**
	* validate excelRowsArray.
  * @param {Array} excelRowsArray - excelRows to be validated.
  * throws errors if any.
*/
function validateRowsArray(excelRowsArray) {
  if (!R.is(Array, excelRowsArray)) {
    throw generateErrorMessage('invalidExcelSheetRows')
  }
}

/**
	* validate lastRowIndex.
  * @param {number} lastRowIndex - row index to be validated.
  * throws error if any.
*/
function validatelastRowIndex(lastRowIndex) {
  if (lastRowIndex == undefined || typeof lastRowIndex !== 'number'){
    throw generateErrorMessage('noLastRowIndex')
  }
}

/**
  * validate workbook object.
  * @param {object} workbook - workbook object to be validated.
  * throws error if any
*/
function validateWorkBook(workbook) {
  if (R.isNil(workbook)){
    throw generateErrorMessage('noWorkbookObjectFound')
  }
  if (!workbook.hasOwnProperty('Sheets')){
    throw generateErrorMessage('invalidWorkBook')
  }
}

/**
 * generates random string
 * @param {number} length - length of random string
 * @param {string} chars - character needed in random string
 * @returns {string} - random string.
 */

function generateRandomFileName(length, chars) {
  var result = ''
  for (var i = length; i > 0; --i) {
    result += chars[Math.round(Math.random() * (chars.length - 1))]
  }
  return result
}

function sanitizeString(value) {
  var str = ''
  if (R.is(String, value))
    var str = value.trim()
  return str
}

module.exports = function(options){
  return new LocalExcelAdaptor(options)
}