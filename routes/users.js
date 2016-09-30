var express = require('express');
var Excel = require('exceljs');
var router = express.Router();

var workbook = new Excel.Workbook();

workbook.creator = 'Arun';
workbook.lastModifiedBy = 'someone';
workbook.created = new Date(2016, 8, 30);
workbook.modified = new Date();

var sheet = workbook.addWorksheet("My sheet");

//noinspection JSUnresolvedFunction
workbook.xlsx.writeFile("Something.xlsx").then(function () {
    console.log("Written");
});

//noinspection JSUnresolvedFunction
workbook.xlsx.readFile("marks2.xlsx").then(function () {
    console.log("File read");

    var worksheet = workbook.getWorksheet("Sheet1");
    worksheet.eachRow({includeEmpty : true}, function (row, rowNumber) {
        console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
    });
});

/* GET users listing. */
router.get('/', function(req, res, next) {
  res.send('respond with a resource');
});

module.exports = router;
