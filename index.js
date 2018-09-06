
var fs = require('fs');
var express = require('express');
var multer  = require('multer');

var XLSX = require('xlsx');

// var workbook = XLSX.readFile('upload/深圳市铱云云计算有限公司.xlsx');
// var sheet_name_list = workbook.SheetNames;
// console.log(XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]))

var app = express();

var createFolder = function(folder){
    try{
        fs.accessSync(folder);
    }catch(e){
        fs.mkdirSync(folder);
    }
};

/*
def read_xls_to_dict_list_attendence(rdbook_obj, sheetx=0, sheet_name=None, key_rows=0):
    if sheet_name:
        sheet = rdbook_obj.sheet_by_name(sheet_name)
    else:
        sheet = rdbook_obj.sheet_by_index(sheetx)

    # read header values into the list
    keys_common = [sheet.cell(key_rows, col_index).value for col_index in range(7)]

    dict_list = []
    for row_index in range(1, sheet.nrows):
        d = {keys_common[col_index]: sheet.cell(row_index, col_index).value
             for col_index in range(7)}
        d['打卡时间1'] = sheet.cell(row_index, 7).value
        d['打卡结果1'] = sheet.cell(row_index, 8).value
        d['打卡时间2'] = sheet.cell(row_index, 9).value
        d['打卡结果2'] = sheet.cell(row_index, 10).value
        dict_list.append(d)

    return dict_list

 */


function read_xls_to_dict_list_attendence(sheetx=0, sheet_name= null, key_rows=0) {

    var rdbook_obj = XLSX.readFile('upload/深圳市铱云云计算有限公司.xlsx');
    var sheet;
    var sheet_name;

    if(sheet_name) {
        sheet = rdbook_obj.Sheets[sheet_name];
    }
    else {
        sheet_name = rdbook_obj.SheetNames[sheetx];
        console.log(sheet_name)
        sheet = rdbook_obj.Sheets[sheet_name];
    }

    //  read header values into the list
    //keys_common = [sheet.cell(key_rows, col_index).value for col_index in range(7)]

    //得到当前页内数据范围
    var range = XLSX.utils.decode_range(sheet['!ref']);
    console.log(range)
    //保存数据范围数据
    var row_start = range.s.r;
    var row_end = range.e.r;
    var col_start = range.s.c;
    var col_end = range.e.c;
    var rows = [];
    var result = {}

    //按行对 sheet 内的数据循环
    for(;row_start<=row_end;row_start++) {
        row_data = [];
        //读取当前行里面各个列的数据
        for(i=col_start;i<=col_end;i++) {
            addr = XLSX.utils.encode_col(i) + XLSX.utils.encode_row(row_start);
            console.log(addr);
            cell = sheet[addr];
            console.log("" + cell);
            //如果是链接，保存为对象，其它格式直接保存原始值
            var desired_value = (cell ? cell.v : undefined);
            row_data.push(desired_value);
        }
        rows.push(row_data);
    }
    //保存当前页内的数据
    result[sheet_name] = rows;

    console.log("" + rows);

    return rows


}

read_xls_to_dict_list_attendence()



var uploadFolder = './upload/';

createFolder(uploadFolder);

// 通过 filename 属性定制
var storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadFolder);    // 保存的路径，备注：需要自己创建
    },
    filename: function (req, file, cb) {
        // 将保存文件名设置为 字段名 + 时间戳，比如 logo-1478521468943
        cb(null, file.fieldname + '-' + Date.now());
    }
});

// 通过 storage 选项来对 上传行为 进行定制化
var upload = multer({ storage: storage });

// 单图上传
app.post('/upload', upload.single('logo'), function(req, res, next){
    var file = req.file;
    res.send({ret_code: '0'});
});

app.get('/form', function(req, res, next){
    var form = fs.readFileSync('./form.html', {encoding: 'utf8'});
    res.send(form);
});

app.get('/xlsx', function(req, res, next){
    var form = read_xls_to_dict_list_attendence();
    res.send(form);
});


app.get('/show', function(req, res, next){
    console.log("Request handler 'show' was called.");
    fs.readFile("./upload/logo-1536246310455", "binary", function(error, file) {
        if(error) {
            res.writeHead(500, {"Content-Type": "text/plain"});
            res.write(error + "\n");
            res.end();
        } else {
            res.writeHead(200, {"Content-Type": "image/png"});
            res.write(file, "binary");
            res.end();
        }
    });
});



app.listen(3000);