const R = require('ramda');

module.exports.stringValidator = function stringValidator(value) {
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

module.exports.numberValidator = function numberValidator(value) {
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

module.exports.booleanValidator = function booleanValidator(value) {
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

module.exports.dateValidator = function dateValidator(value) {
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