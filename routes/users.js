var express = require('express');
var Excel = require('exceljs');
var router = express.Router();

var workbook = new Excel.Workbook();

var options = {
    filename: './employee_new.xlsx',
    useStyles: true,
    useSharedStrings: true
};

var workbook_streamer = new Excel.stream.xlsx.WorkbookWriter(options);

var worksheet_streamer = workbook_streamer.addWorksheet("new emp sheet");

worksheet_streamer.columns = [
    {header: "Id", key: "id", width: 10},
    {header: "Name", key: "name", width: 32},
    {header: "Roll no", key: "rno", width: 10}
];

worksheet_streamer.addRow({
    id: '1',
    name: 'Arun',
    rno: '34'
});

worksheet_streamer.addRow({
    id: '2',
    name: 'Aditya',
    rno: '27'
});

worksheet_streamer.addRow({
    id: '3',
    name: 'Suyog',
    rno: '12'
});

worksheet_streamer.getCell('F5').value = 'B5';

worksheet_streamer.getRow(3).commit();

worksheet_streamer.commit();
workbook_streamer.commit().then(function () {
    console.log("Workbook committed");
});


/*workbook.creator = 'Arun';
workbook.lastModifiedBy = 'someone';
workbook.created = new Date(2016, 8, 30);
workbook.modified = new Date();

var sheet = workbook.addWorksheet("My sheet");

//noinspection JSUnresolvedFunction
workbook.xlsx.writeFile("Something.xlsx").then(function () {
    console.log("Written");
});*/

//noinspection JSUnresolvedFunction

workbook.xlsx.readFile("employee.xlsx").then(function () {
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
