'use strict';

const XLSX = require('xlsx');

const {log} = require('./utils');

module.exports = function createXlsx(data, outdir) {
  const {employees, projects} = data;

  log('Creating XLSX...');
  const SHEET_NAME = 'Лист 1';
  const book = XLSX.utils.book_new();

  // Document properties
  book.Props = {
    Title: "Отчёт из Мегаплана",
    Subject: "Стоимость работ / затраченное время, ресурсы",
    Author: "Alexey Zakharov",
    CreatedDate: new Date()
  };

  // Init variables
  let lineNum = 0;
  const employeeColProps = employees.map(e => ({title: e.name, width: 18}));
  let commonColProps = [
    {title: 'Проект', width: 38},
    {title: 'Задача', width: 44},
    {title: 'Затраченное время', width: 17},
  ];
  const COL_PROJ = 0;
  const COL_TASK = 1;
  const COL_WORK = 2;
  const employeeColStart = commonColProps.length;
  commonColProps = [...commonColProps, ...employeeColProps];
  const cols = commonColProps.map(c => ({wch: c.width}));

  // Create worksheet
  book.SheetNames.push(SHEET_NAME);
  book.Sheets[SHEET_NAME] = {
    '!ref': 'A1:',
    '!cols': cols
  };
  const sheet = book.Sheets[SHEET_NAME];

  // Draw the header
  commonColProps.forEach((col, colInd) => {
    const cell = {
      t: "s",
      v: col.title,
      s: {
        font: {
          bold: true
        }
      }
    };

    drawCell(cell, sheet, lineNum, colInd);
  });
  lineNum += 1;

  // Draw the data table's body
  const projLineStyle = {
    fill: {
      fgColor: {rgb: "FFFF75"} // Actually set's background
    },
    border: {
      top: {style: "thin", color: {auto: 1}},
      right: {style: "thin", color: {auto: 1}},
      bottom: {style: "thin", color: {auto: 1}},
      left: {style: "thin", color: {auto: 1}}
    }
  };
  projects.forEach(prj => {
    // Draw project line (totals)
    drawCell({t: 's', v: prj.name, s: projLineStyle}, sheet, lineNum, COL_PROJ);
    drawCell({t: 's', v: '', s: projLineStyle}, sheet, lineNum, COL_TASK);
    drawCell({t: 'n', z: '0', v: prj.totalWork, s: projLineStyle}, sheet, lineNum, COL_WORK);
    // Draw work hours per employee
    employees.forEach((empl, emplInd) => {
      drawCell({t: 'n', z: '0', v: empl.proj2work[prj.id] || 0, s: projLineStyle}, sheet,
        lineNum, employeeColStart + emplInd);
    });
    lineNum += 1;

    // Draw the project's task lines
    prj.tasks.forEach(task => {
      drawCell({t: 's', v: task.name}, sheet, lineNum, COL_TASK);
      drawCell({t: 'n', z: '0', v: task.totalWork}, sheet, lineNum, COL_WORK);
      // Draw work hours per employee
      employees.forEach((empl, emplInd) => {
        drawCell({t: 'n', z: '0', v: task.employee2work[empl.id] || 0}, sheet, lineNum, employeeColStart + emplInd);
      });
      lineNum += 1;
    });
  });

  // Finalize the document (finalize sheet)
  const lastColIndex = cols.length - 1;
  const endOfWsRange = XLSX.utils.encode_cell({c: lastColIndex, r: lineNum});
  sheet['!ref'] += endOfWsRange;

  // Write report to file
  const xlsPath = `${outdir}/report.xlsx`;
  XLSX.writeFile(book, xlsPath);
  log(`Saved report to '${xlsPath}'`);
};

function drawCell(cell, sheet, row, col) {
  const cellAddress = {c: col, r: row};
  const cellRef = XLSX.utils.encode_cell(cellAddress);
  sheet[cellRef] = cell;
}
