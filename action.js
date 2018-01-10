'use strict';
const xlsx = require('xlsx');
const fs = require('fs');
const process = require('process');
const cp = require('child_process');

const json = JSON.parse(fs.readFileSync("./config/config.json"));

for (let o in json){
    if (o === "keyword") continue;
    action(json, o);
}

function action(json, o) {
    const path = './' + json[o]["file"];
    const workbook = xlsx.readFile(path);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    deleteBlankRow(worksheet);
    processData(worksheet, json[o]["data"]);
    xlsx.writeFile(workbook, path);
}

function svnCommit(path, message) {
    process.chdir(path);
    execute("svn commit -m " + message);
}

function turnData(path) {
    process.chdir(path);
    execute("转部分数据.bat");
}

function execute(command) {
    cp.exec(command, function (err, stdout, stderr) {
        if(err){
            console.log('stderr: ' + stderr);
        } else {
            console.log('stdout: ' + stdout);
        }
    });
}

function getZGDate() {
    const nowDate = new Date();
    const week = nowDate.getDay();
    const nowDateShort = new Date(nowDate.getFullYear(), nowDate.getMonth(), nowDate.getDate());
    if (4 - week >= 0){
        return getDelayDate(nowDateShort, 4-week);
    } else {
        return getDelayDate(nowDateShort, 11-week);
    }
}

function  getDelayDate(date, delayDate, delayHour=0, delayMinute=0, delaySecond=0) {
    date.setDate(date.getDate() + delayDate);
    date.setHours(date.getHours() + delayHour,
        date.getMinutes() + delayMinute,
        date.getSeconds() + delaySecond);
    return date;
}

function getFormatDate(date) {
    return date.getFullYear() + "-" + to2digit(date.getMonth() + 1) + "-" + to2digit(date.getDate()) + " "
        + to2digit(date.getHours()) + ":" + to2digit(date.getMinutes()) + ":" + to2digit(date.getSeconds());
}

function to2digit(num) {
    if (num < 10) {
        return "0" + num;
    } else {
        return num;
    }
}

function getResultDate(date, time) {
    const timeArr = time.split(':');
    const hour = parseInt(timeArr[0], 10);
    const minute = parseInt(timeArr[1], 10);
    const second = parseInt(timeArr[2], 10);
    return getFormatDate(getDelayDate(getZGDate(), date, hour, minute, second));
}

function processData(worksheet, data) {
    for (let o in data) {
        if (typeof data[o]["pos"] === "string") {
            processDataType(worksheet, o, data[o]["pos"], 0);
        } else {
            for (let i = 0; i < data[o]["pos"].length; i++) {
                processDataType(worksheet, o, data[o]["pos"][i], i);
            }
        }
    }
}

function processDataType(worksheet, o, pos, i) {
    let value;
    switch (o) {
        case "string":
            if (data[o]["type"] === "change") {
                valueToExcel(worksheet, pos, data[o]["value"][i]);
            }
            if (data[o]["type"] === "add") {
                valueToExcel(worksheet, getAddPos(pos), data[o]["value"][i])
            }
            break;
        case "date":
            if (data[o]["type"] === "change") {
                value = getResultDate(data[o]["delayDate"][i], data[o]["delayTime"][i]);
                valueToExcel(worksheet, pos, value);
            }
            if (data[o]["type"] === "add") {
                value = getResultDate(data[o]["delayDate"][i], data[o]["delayTime"][i]);
                valueToExcel(worksheet, getAddPos(worksheet, pos), value)
            }
            break;
        case "order":
            orderValue(worksheet, pos);
            break;
        case "input":
            //TODO
            break;
        default:
            console.log("未知的类型：" + o["type"]);
    }
}

function deleteBlankRow(worksheet) {
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    let cellRef;
    let count;
    let deleteRowNum = 0;
    for (let row = range.e.r; row > range.s.r; row--) {
        count = 0;
        for (let col = range.e.c; col > range.s.c; col--) {
            cellRef = xlsx.utils.encode_cell({c: col, r: row});
            if (worksheet[cellRef] == null) count++;
        }
        if (count === range.e.c - range.s.c) {
            deleteRowNum++;
        } else {
            break;
        }
    }
    const range2 = {
        s: {c: range.s.c, r:range.s.r},
        e: {c: range.e.c, r:range.e.r - deleteRowNum}
    };
    worksheet['!ref'] = xlsx.utils.encode_range(range2);
}

function orderValue(worksheet, pos) {
    let originalRow = xlsx.utils.decode_range(worksheet['!ref']).e.r;
    let cellAbove;
    let cellBelow;
    const address = xlsx.utils.decode_range(pos);
    for (let i = address.s.c; i <= address.e.c; ++i) {
        cellAbove = xlsx.utils.encode_cell({c: i, r: originalRow});
        while (worksheet[cellAbove] == null) {
            originalRow--;
            if (originalRow < 0) return;
            cellAbove = xlsx.utils.encode_cell({c: i, r: originalRow});
        }
        cellBelow = xlsx.utils.encode_cell({c: i, r: originalRow + 1});
        valueToExcel(worksheet, cellBelow, worksheet[cellAbove].v + 1);
    }
}

function valueToExcel(worksheet, pos, value) {
    //扩充range
    const address = xlsx.utils.decode_range(pos);
    if (worksheet['!ref'] == null) {
        worksheet['!ref'] = xlsx.utils.encode_range(address);
    } else {
        const range = xlsx.utils.decode_range(worksheet['!ref']);
        const range2 = {
            s: {c: (range.s.c > address.s.c) ? address.s.c : range.s.c,
                r: (range.s.r > address.s.r) ? address.s.r : range.s.r},
            e: {c: (range.e.c < address.e.c) ? address.e.c : range.e.c,
                r: (range.e.r < address.e.r) ? address.e.r : range.e.r}
        };
        worksheet['!ref'] = xlsx.utils.encode_range(range2);
    }
    //赋值
    let cellRef;
    for (let i = address.s.c; i <= address.e.c; ++i) {
        for (let j = address.s.r; j <= address.e.r; ++j) {
            cellRef = xlsx.utils.encode_cell({c: i, r: j});
            if (typeof value === "number") {
                worksheet[pos] = {t: 'n', v: value};
            } else {
                worksheet[pos] = {t: 's', v: value};
            }
        }
    }
}

function getAddPos(worksheet, pos) {
    const originalRow = xlsx.utils.decode_range(worksheet['!ref']).e.r;
    const address = xlsx.utils.decode_range(pos);
    const address2 = {
        s: {c: address.s.c, r: originalRow + address.s.r},
        e: {c: address.e.c, r: originalRow + address.e.r}
    };
    return xlsx.utils.encode_range(address2);
}