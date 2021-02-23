#!/usr/bin/env node

'use strict';

const {log} = require('./utils');
const createXlsx = require('./create_xlsx');

const main = () => {
  const outdir = process.cwd();

  const data = {
    employees: [
      {
        id: 'Alexey',
        totalWork: 1,
        name: 'Alexey',
        proj2work: {
          "1": 4,
          "2": 4,
        },
      },
      {
        id: 'Anton',
        totalWork: 2,
        name: 'Anton',
        proj2work: {
          "1": 6,
          "2": 6,
        },
      },
      {
        id: 'Vadim',
        totalWork: 3,
        name: 'Vadim',
        proj2work: {
          "1": 4,
          "2": 10,
        },
      },
    ],
    projects: [
      {
        id: '1',
        name: 'Проект №1',
        totalWork: 14,
        tasks: [
          {
            name: 'task 11',
            totalWork: 7,
            employee2work: {
              'Alexey': 2,
              'Anton': 3,
              'Vadim': 2,
            },
          },
          {
            name: 'task 12',
            totalWork: 7,
            employee2work: {
              'Alexey': 2,
              'Anton': 3,
              'Vadim': 2,
            },
          },
        ],
      },
      {
        id: '2',
        name: 'Проект №2',
        totalWork: 20,
        tasks: [
          {
            name: 'task 21',
            totalWork: 10,
            employee2work: {
              'Alexey': 2,
              'Anton': 3,
              'Vadim': 5,
            },
          },
          {
            name: 'task 22',
            totalWork: 10,
            employee2work: {
              'Alexey': 2,
              'Anton': 3,
              'Vadim': 5,
            },
          },
        ],
      },
    ],
  };

  // Write data to XLSX
  createXlsx(data, outdir);

  log('DONE');
}

// Start the program
main();
