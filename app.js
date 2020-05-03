const Excel = require('exceljs');
const data = require('./data.json')
const express = require('express');
const app = express();

var img = "./logo.png";    


async function makeSheet(){
    const workbook = new Excel.Workbook();
    workbook.created = new Date();
    workbook.modified = new Date();
    const title = data[0].title;
    const worksheet = workbook.addWorksheet(title);
    var imageId1 = workbook.addImage({ 
        filename: img,
        extension: 'png',
     });
    worksheet.addImage(imageId1, {
        tl: {col: 0, row: 0},
        br: {col: 1, row:7}
    });
    worksheet.getCell('A11').value = title;
    worksheet.getCell('A11').font = {
        bold: true,
        size: 20
    };
    worksheet.mergeCells('A11:D11');
    worksheet.mergeCells('A1:D7');
    // worksheet.properties.showGridLines = false;
    worksheet.views = [{
        showGridLines: false
    }];
    // worksheet.getRow(1).hidden = true;
    worksheet.getCell('A11').alignment = {horizontal: 'center', vertical: 'middle'};
    worksheet.getRow(11).height = 40;
    worksheet.getRow(11).outlineLevel = 5;

    
    worksheet.columns = [{key:'Vehicle'}, {key:'Date and Time'}, {key:'Location'}, {key:'Speed'}]
    worksheet.addRow(12).values = ['', '', '', ''];
    worksheet.mergeCells('A12:D12');
    worksheet.addRow(13).values = ['Vehicle', 'Date and Time', 'Location', 'Speed'];
    worksheet.getRow(13).font = {bold: true};
    // worksheet.addRow(12).values = ['', '', '', ''];
    worksheet.columns.forEach(column => {
        console.log(column.key);
        if(column.key.length < 10) column.width = 14
        else column.width = column.key.length + 5
        // column.width = column.key.length < 12 ? 12 : column.key.length
      });
//     let rowIndex = 1;
//     for (rowIndex; rowIndex <= worksheet.rowCount; rowIndex++) {
//         worksheet.getRow(rowIndex).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
// }

    // thead = [{header:'Vehicle', key:'Vehicle', width: 30},{header:'Date and Time and Time', key:'Date and Time and Time', width: 30},{header:'Location', key:'Location', width:50},{header:'Speed', key:'Speed', width:12}];
    // worksheet.columns = thead;
    
    worksheet.getCell('A13').fill = {type: 'pattern',pattern: 'solid',fgColor:{argb:'FF0000'}};
    worksheet.getCell('B13').fill = {type: 'pattern',pattern: 'solid',fgColor:{argb:'FF0000'}};
    worksheet.getCell('C13').fill = {type: 'pattern',pattern: 'solid',fgColor:{argb:'FF0000'}};
    worksheet.getCell('D13').fill = {type: 'pattern',pattern: 'solid',fgColor:{argb:'FF0000'}};
    for(i=1;i<data.length;i++)
    {
        worksheet.addRow(data[i]);
    }
    var borderStyles = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
    };
      
    worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
        row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
          cell.border = borderStyles;
        });
    });
    await workbook.xlsx.writeFile('export.xlsx');
    console.log('File is written');  
};

app.get('/', function(req, res) {
    res.send('<h1> Hi, redirect to localhost:3040/excel to download the excel sheet.</h1>');
});

app.get('/excel', function (req, res, next) {

    const workbook = new Excel.Workbook();
    workbook.created = new Date();
    workbook.modified = new Date();
    const title = data[0].title;
    const worksheet = workbook.addWorksheet(title);
    var imageId1 = workbook.addImage({ 
        filename: img,
        extension: 'png',
     });
    worksheet.addImage(imageId1, {
        tl: {col: 0, row: 0},
        br: {col: 1, row:7}
    });
    worksheet.getCell('A11').value = title;
    worksheet.getCell('A11').font = {
        bold: true,
        size: 20
    };
    worksheet.mergeCells('A11:D11');
    worksheet.mergeCells('A1:D7');

    worksheet.views = [{
        showGridLines: false
    }];

    worksheet.getCell('A11').alignment = {horizontal: 'center', vertical: 'middle'};
    worksheet.getRow(11).height = 40;
    worksheet.getRow(11).outlineLevel = 5;

    worksheet.columns = [{key:'Vehicle'}, {key:'Date and Time'}, {key:'Location'}, {key:'Speed'}]
    worksheet.addRow(12).values = ['', '', '', ''];
    worksheet.mergeCells('A12:D12');
    worksheet.addRow(13).values = ['Vehicle', 'Date and Time', 'Location', 'Speed'];
    worksheet.getRow(13).font = {bold: true};
    
    worksheet.columns.forEach(column => {
        console.log(column.key);
        if(column.key.length < 10) column.width = 14
        else column.width = column.key.length + 5
      });
    
    worksheet.getCell('A13').fill = {type: 'pattern',pattern: 'solid',fgColor:{argb:'FF0000'}};
    worksheet.getCell('B13').fill = {type: 'pattern',pattern: 'solid',fgColor:{argb:'FF0000'}};
    worksheet.getCell('C13').fill = {type: 'pattern',pattern: 'solid',fgColor:{argb:'FF0000'}};
    worksheet.getCell('D13').fill = {type: 'pattern',pattern: 'solid',fgColor:{argb:'FF0000'}};
    
    for(i=1;i<data.length;i++) {
        worksheet.addRow(data[i]);
    };

    var borderStyles = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
    };
      
    worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
        row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
          cell.border = borderStyles;
        });
    });
    workbook.xlsx.writeFile('json-to-excel.xlsx').then(function() {
        console.log('file is written');
        res.download('json-to-excel.xlsx', function(err){
            console.log('----------');
        });
    });
});

makeSheet();

app.listen(3040, function () {
    console.log('Excel app listening on port 3040');
});

