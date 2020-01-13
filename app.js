const pdf2table = require('pdf2table');
const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');

const rateSheet1 = './rate-sheets/para01.pdf';
const rateSheet2 = './rate-sheets/para02.pdf';
const rateSheet3 = './rate-sheets/para03.pdf';
const rateSheet5 = './rate-sheets/para05.pdf';
const rateSheet6 = './rate-sheets/para06.pdf';
const rateSheet7 = './rate-sheets/para07.pdf';
const rateSheet8 = './rate-sheets/para08.pdf';
const rateSheet9 = './rate-sheets/para09.pdf';

currentSheet = rateSheet1;

function parsePDF() {
    fs.readFile(currentSheet, function(err, buffer) {
        if (err) return console.log(err);

        pdf2table.parse(buffer, function (err, rows) {
            if (err) return console.log(err);
            getData(rows);
        })
    });
}

function getData(data) {
    if (!data) {
        return;
    }
    
    const plan = data.splice(0, 24);
    createRow(plan, data);
}

let excelRows = [];

function createRow(plan, data) {
    if (plan.length > 0) {
        const startDate = plan[0].toString().slice(28,37);
        const endDate = plan[0].toString().slice(41, 50);
        const planName = plan[2][3];
        const ratingArea = plan[1][1].slice(4);
        const state = plan[1][1].slice(0,2);
    
        let rateArray = [];
        
        const rateSection = plan.slice(4, 19);
    
        rateSection.forEach(row => {
            row.map(item => {
                if (item.length > 4) {
                    rateArray.push(Number(item));
                }
            })
        })
    
        const rates = rateArray.sort((a,b) => a - b);
        const row = [startDate, endDate, planName, state, ratingArea, rates[0], ...rates, rates[44]];
        
        excelRows.push(row);
        getData(data);     
    } else {
        exportToExcel(excelRows);
    }
}

function exportToExcel(data) {
    XlsxPopulate.fromFileAsync('./benefix-template.xlsx')
        .then(workbook => {
            workbook.sheet('Blank Upload Template').cell('A2').value(data);            
            return workbook.toFileAsync("./benefix-template.xlsx");
        })
}

parsePDF();