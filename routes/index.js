var express = require('express');
var excel = require('./excel');
var XLSX = require('xlsx');
var Busboy = require('busboy');
var router = express.Router();

var db = [];

/* GET home page. */
router.get('/', function (req, res, next) {
    res.render('index', {title: 'Express'});
});
router.get('/test', function (req, res, next) {
    var _headers = [
        {caption: '学号', type: 'string'},
        {caption: '密码', type: 'string'},
        {caption: '姓名', type: 'string'}];
    var rows = [];
    if (db.length > 0) {
        rows = db;
    } else {
        rows = [
            ['13111041', '13111041', '孙颢'],
            ['14111037', '14111037', '不知道']
        ];
    }

    var result = excel.exportExcel(_headers, rows);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats');
    res.setHeader("Content-Disposition", "attachment; filename=" + "test.xlsx");
    return res.end(result, 'binary');
});

router.post('/post', function (req, res, next) {
    var busboy = new Busboy({
        headers: req.headers,
        limits: {
            files: 1,
            fileSize: 50000000
        }
    });
    busboy.on('file', function (fieldname, file, filename, encoding, mimetype) {
        file.on('limit', function () {
            return res.json(Result.FAIL('To large'));
        });
        file.on('data', function (data) {
            db = [];
            console.log('File [' + fieldname + '] got ' + data.length + ' bytes');

            var workbook = XLSX.read(data);
            var sheetNames = workbook.SheetNames; // 返回 ['sheet1', 'sheet2',……]
            var worksheet = workbook.Sheets[sheetNames[0]];// 获取excel的第一个表格
            var ref = worksheet['!ref']; //获取excel的有效范围,比如A1:F20
            var reg = /[a-zA-Z]/g;
            ref = ref.replace(reg, "");
            var line = parseInt(ref.split(':')[1]); // 获取excel的有效行数

            console.log("line====>", line);

            // header ['姓名','邮箱','身份','部门','手机号']

            //循环读出每一行，然后处理
            for (var i = 2; i <= line; i++) {
                if (!worksheet['A' + i] && !worksheet['B' + i] && !worksheet['C' + i] && i !== 2) {   //如果大于2的某行为空,则下面的行不算了
                    break;
                }

                var number = worksheet['A' + i].v || '';
                var psd = worksheet['B' + i].v || '';
                var name = worksheet['C' + i].v || '';

                var step = [];
                step.push(number);
                step.push(psd);
                step.push(name);
                db.push(step);
            }

            res.send({status: 200, model: '导入成功'});
        });
    });
    return req.pipe(busboy);
});
module.exports = router;
