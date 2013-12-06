/* Copyright (C) 2013  SheetJS */

var XLSX = require('xlsx');
var XLS = require('xlsjs');
var fs = require('fs');
var Buffer = require('buffer').Buffer;

var error = function(message) {
    return Error(message);
}

var readFile = function(filename, options, callback) {
    fs.open(filename, 'r', function(status, fd) {
        if (status) {
            console.error('fd status exception', status.message);
            return callback(error(status.message));
        }
        var buffer = new Buffer(1);
        fs.read(fd, buffer, 0, 1, 0, function(err, num) {
            if (err) {
                return callback(err);
            }
            switch(buffer[0]) {
                /* CFB container */
                case 0xd0: return callback(null, [XLS,   XLS.readFile(filename)]);
                /* Zip container */
                case 0x50: return callback(null, [XLSX, XLSX.readFile(filename)]);
            }
            return callback(error('unrecognized file type'), [undefined, buffer]);
        });
    });
};

var readFileSync = function(filename, options) {
    var f = fs.readFileSync(filename);
    switch(f[0]) {
        /* CFB container */
        case 0xd0: return [XLS,   XLS.readFile(filename)];
        /* Zip container */
        case 0x50: return [XLSX, XLSX.readFile(filename)];
    }
    return [undefined, f];
};

// ⚑
// TODO: Maintaining backward merge compat with the naming here (but it's weird)
module.exports = {
    XLS: XLS,
    XLSX: XLSX,
    readFile: readFileSync,
    readFileAsync: readFile
};
