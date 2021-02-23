#!/usr/bin/env node

'use strict';

const exit = process.exit;
const moment = require('moment');
const chalk = require('chalk');

const {log, stringify} = require('./utils');
const createXlsx = require('./create_xlsx');

async function main() {
  let dtStart = moment().startOf('month').startOf('day');
  let dtEnd = moment().endOf('day');

  const outdir = process.cwd();

  // Write data to XLSX
  createXlsx(dtStart, dtEnd, outdir);

  log(chalk.green('OK'));
}

// Start the program
main();
