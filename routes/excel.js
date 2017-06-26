var nodeExcel = require('excel-export');
/**
 *
 * @param _headers example  [
 {caption:'用户状态',type:'string'},
 {caption:'部门',type:'string'},
 {caption:'姓名',type:'string'},
 {caption:'邮箱',type:'string'},
 {caption:'有效期',type:'string'},
 {caption:'身份',type:'string'}];
 * @param rows example
 [['未激活','信息部','testname','123@qq.com','2019-11-09','管理员'],
 ['未激活','信息部','testname2','12345@qq.com','2019-11-09','普通成员']]
 */
module.exports.exportExcel = function (_headers, rows) {
    var conf = {};
    conf.name = "mysheet";
    conf.cols = [];
    for (var i = 0; i < _headers.length; i++) {
        var col = {};
        col.caption = _headers[i].caption;
        col.type = _headers[i].type;
        conf.cols.push(col);
    }
    conf.rows = rows;
    return nodeExcel.execute(conf);
};