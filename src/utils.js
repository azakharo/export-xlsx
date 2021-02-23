'use strict';

function log(msg) {
  console.log(msg);
}

function stringify(data) {
  return JSON.stringify(data, null, 2);
}

function logData(data) {
  console.log(stringify(data));
}


module.exports = {
  log,
  stringify,
  logData,
};
