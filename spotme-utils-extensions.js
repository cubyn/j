var xls_utils  = require('xlsjs').utils
  , xlsx_utils = require('xlsx').utils
  ;

function format_cell(cell, v) {
    if(!cell) return "";
    if(typeof cell.w !== 'undefined') return cell.w;
    if(typeof v === 'undefined') v = cell.v;
    if(!cell.XF) return v;
    try { cell.w = SSF.format(cell.XF.ifmt||0, v); } catch(e) { return v; }
    return cell.w;
}

var XLS = { utils: {
    // originally https://github.com/SheetJS/js-xls/blob/4076850087785b76e6814213877eea99a8390be5/xls.js#L5135
    sheet_to_row_object_array_with_column_index_props:
        function(sheet, opts) {
            var val, row, r, hdr = {}, isempty, R, C, v;
            var out = [];
            opts = opts || {};
            if(!sheet || !sheet["!ref"]) return out;
            r = xls_utils.decode_range(sheet["!ref"]);
            for(R=r.s.r, C = r.s.c; C <= r.e.c; ++C) {
                val = sheet[xls_utils.encode_cell({c: C, r:R})];
                if(!val) continue;
                hdr[C] = C;
            }

            for (R = r.s.r; R <= r.e.r; ++R) {
                isempty = true;
                /* row index available as __rowNum__ */
                row = Object.create({ __rowNum__ : R });
                for (C = r.s.c; C <= r.e.c; ++C) {
                    val = sheet[xls_utils.encode_cell({c: C, r: R})];
                    if(!val || !val.t) continue;
                    v = (val || {}).v;
                    switch(val.t){
                        case 'e': continue; /* TODO: emit error text? */
                        case 's': case 'str': break;
                        case 'b': case 'n': break;
                        default: throw 'unrecognized type ' + val.t;
                    }
                    if(typeof v !== 'undefined') {
                        row[hdr[C]] = opts.raw ? v || val.v : xls_utils.format_cell(val, v);
                        isempty = false;
                    }
                }
                if(!isempty) out.push(row);
            }
            return out;
        }
} };

var XLSX = { utils: {
    // originally https://github.com/SheetJS/js-xlsx/blob/37cc0006f2aca061df8d859b28f7a982397fcfc7/xlsx.js#L2972
    sheet_to_row_object_array_with_column_index_props:
        function(sheet, opts) {
            var val, row, r, hdr = {}, isempty, R, C;
            var out = [];
            opts = opts || {};
            if(!sheet || !sheet["!ref"]) return out;
            r = xlsx_utils.decode_range(sheet["!ref"]);
            for(R=r.s.r, C = r.s.c; C <= r.e.c; ++C) {
                val = sheet[xlsx_utils.encode_cell({c: C, r: R})];
                if(!val) continue;
                hdr[C] = C;
            }

            for (R = r.s.r; R <= r.e.r; ++R) {
                isempty = true;
                /* row index available as __rowNum__ */
                row = Object.create({ __rowNum__ : R });
                for (C = r.s.c; C <= r.e.c; ++C) {
                    val = sheet[xlsx_utils.encode_cell({c: C, r: R})];
                    if(!val || !val.t) continue;
                    if(typeof val.w !== 'undefined' && !opts.raw) {
                        row[hdr[C]] = val.w; isempty = false;
                    }
                    else switch(val.t){
                        case 's': case 'str': case 'b': case 'n':
                            if(typeof val.v !== 'undefined') {
                                row[hdr[C]] = val.v;
                                isempty = false;
                            }
                            break;
                        case 'e': break; /* throw */
                        default: throw 'unrecognized type ' + val.t;
                    }
                }
                if(!isempty) out.push(row);
            }
            return out;
        }
} };

exports.extendXLS = function(_XLS) {
    _XLS.utils.sheet_to_row_object_array_with_column_index_props
        = XLS.utils.sheet_to_row_object_array_with_column_index_props;
    return XLS;
};

exports.extendXLSX = function(_XLSX) {
    _XLSX.utils.sheet_to_row_object_array_with_column_index_props
        = XLSX.utils.sheet_to_row_object_array_with_column_index_props;
    return XLSX;
};
