/*::
type MJRObject = {
	row: any;
	isempty: boolean;
};
*/
function make_json_row(sheet/*:Worksheet*/, r/*:Range*/, R/*:number*/, cols/*:Array<string>*/, header/*:number*/, hdr/*:Array<any>*/, dense/*:boolean*/, o/*:Sheet2JSONOpts*/)/*:MJRObject*/ {
	var rr = encode_row(R);
	var defval = o.defval, raw = o.raw || !Object.prototype.hasOwnProperty.call(o, "raw");
	var isempty = true;
	var row/*:any*/ = (header === 1) ? [] : {};
	if(header !== 1) {
		if(Object.defineProperty) try { Object.defineProperty(row, '__rowNum__', {value:R, enumerable:false}); } catch(e) { row.__rowNum__ = R; }
		else row.__rowNum__ = R;
	}
	if(!dense || sheet[R]) for (var C = r.s.c; C <= r.e.c; ++C) {
		var val = dense ? sheet[R][C] : sheet[cols[C] + rr];
		if(val === undefined || val.t === undefined) {
			if(defval === undefined) continue;
			if(hdr[C] != null) { row[hdr[C]] = defval; }
			continue;
		}
		var v = val.v;
		switch(val.t){
			case 'z': if(v == null) break; continue;
			case 'e': v = void 0; break;
			case 's': case 'd': case 'b': case 'n': break;
			default: throw new Error('unrecognized type ' + val.t);
		}
		if(hdr[C] != null) {
			if(v == null) {
				if(defval !== undefined) row[hdr[C]] = defval;
				else if(raw && v === null) row[hdr[C]] = null;
				else continue;
			} else {
				row[hdr[C]] = raw || (o.rawNumbers && val.t == "n") ? v : format_cell(val,v,o);
			}
			if(v != null) isempty = false;
		}
	}
	return { row: row, isempty: isempty };
}


function sheet_to_json(sheet/*:Worksheet*/, opts/*:?Sheet2JSONOpts*/) {
	if(sheet == null || sheet["!ref"] == null) return [];
	var val = {t:'n',v:0}, header = 0, offset = 1, hdr/*:Array<any>*/ = [], v=0, vv="";
	var r = {s:{r:0,c:0},e:{r:0,c:0}};
	var o = opts || {};
	var range = o.range != null ? o.range : sheet["!ref"];
	if(o.header === 1) header = 1;
	else if(o.header === "A") header = 2;
	else if(Array.isArray(o.header)) header = 3;
	else if(o.header == null) header = 0;
	switch(typeof range) {
		case 'string': r = safe_decode_range(range); break;
		case 'number': r = safe_decode_range(sheet["!ref"]); r.s.r = range; break;
		default: r = range;
	}
	if(header > 0) offset = 0;
	var rr = encode_row(r.s.r);
	var cols/*:Array<string>*/ = [];
	var out/*:Array<any>*/ = [];
	var outi = 0, counter = 0;
	var dense = Array.isArray(sheet);
	var R = r.s.r, C = 0, CC = 0;
	if(dense && !sheet[R]) sheet[R] = [];
	for(C = r.s.c; C <= r.e.c; ++C) {
		cols[C] = encode_col(C);
		val = dense ? sheet[R][C] : sheet[cols[C] + rr];
		switch(header) {
			case 1: hdr[C] = C - r.s.c; break;
			case 2: hdr[C] = cols[C]; break;
			case 3: hdr[C] = o.header[C - r.s.c]; break;
			default:
				if(val == null) val = {w: "__EMPTY", t: "s"};
				vv = v = format_cell(val, null, o);
				counter = 0;
				for(CC = 0; CC < hdr.length; ++CC) if(hdr[CC] == vv) vv = v + "_" + (++counter);
				hdr[C] = vv;
		}
	}
	for (R = r.s.r + offset; R <= r.e.r; ++R) {
		var row = make_json_row(sheet, r, R, cols, header, hdr, dense, o);
		if((row.isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) out[outi++] = row.row;
	}
	out.length = outi;
	return out;
}

var qreg = /"/g;
function make_csv_row(sheet/*:Worksheet*/, r/*:Range*/, R/*:number*/, cols/*:Array<string>*/, fs/*:number*/, rs/*:number*/, FS/*:string*/, o/*:Sheet2CSVOpts*/)/*:?string*/ {
	var isempty = true;
	var row/*:Array<string>*/ = [], txt = "", rr = encode_row(R);
	for(var C = r.s.c; C <= r.e.c; ++C) {
		if (!cols[C]) continue;
		var val = o.dense ? (sheet[R]||[])[C]: sheet[cols[C] + rr];
		if(val == null) txt = "";
		else if(val.v != null) {
			isempty = false;
			txt = ''+(o.rawNumbers && val.t == "n" ? val.v : format_cell(val, null, o));
			for(var i = 0, cc = 0; i !== txt.length; ++i) if((cc = txt.charCodeAt(i)) === fs || cc === rs || cc === 34 || o.forceQuotes) {txt = "\"" + txt.replace(qreg, '""') + "\""; break; }
			if(txt == "ID") txt = '"ID"';
		} else if(val.f != null && !val.F) {
			isempty = false;
			txt = '=' + val.f; if(txt.indexOf(",") >= 0) txt = '"' + txt.replace(qreg, '""') + '"';
		} else txt = "";
		/* NOTE: Excel CSV does not support array formulae */
		row.push(txt);
	}
	if(o.blankrows === false && isempty) return null;
	return row.join(FS);
}

function sheet_to_csv(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/)/*:string*/ {
	var out/*:Array<string>*/ = [];
	var o = opts == null ? {} : opts;
	if(sheet == null || sheet["!ref"] == null) return "";
	var r = safe_decode_range(sheet["!ref"]);
	var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
	var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
	var endregex = new RegExp((FS=="|" ? "\\|" : FS)+"+$");
	var row = "", cols/*:Array<string>*/ = [];
	o.dense = Array.isArray(sheet);
	var colinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!cols"] || [];
	var rowinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!rows"] || [];
	for(var C = r.s.c; C <= r.e.c; ++C) if (!((colinfo[C]||{}).hidden)) cols[C] = encode_col(C);
	for(var R = r.s.r; R <= r.e.r; ++R) {
		if ((rowinfo[R]||{}).hidden) continue;
		row = make_csv_row(sheet, r, R, cols, fs, rs, FS, o);
		if(row == null) { continue; }
		if(o.strip) row = row.replace(endregex,"");
		out.push(row + RS);
	}
	delete o.dense;
	return out.join("");
}

function sheet_to_txt(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
	if(!opts) opts = {}; opts.FS = "\t"; opts.RS = "\n";
	var s = sheet_to_csv(sheet, opts);
	if(typeof cptable == 'undefined' || opts.type == 'string') return s;
	var o = cptable.utils.encode(1200, s, 'str');
	return String.fromCharCode(255) + String.fromCharCode(254) + o;
}

function sheet_to_formulae(sheet/*:Worksheet*/)/*:Array<string>*/ {
	var y = "", x, val="";
	if(sheet == null || sheet["!ref"] == null) return [];
	var r = safe_decode_range(sheet['!ref']), rr = "", cols/*:Array<string>*/ = [], C;
	var cmds/*:Array<string>*/ = [];
	var dense = Array.isArray(sheet);
	for(C = r.s.c; C <= r.e.c; ++C) cols[C] = encode_col(C);
	for(var R = r.s.r; R <= r.e.r; ++R) {
		rr = encode_row(R);
		for(C = r.s.c; C <= r.e.c; ++C) {
			y = cols[C] + rr;
			x = dense ? (sheet[R]||[])[C] : sheet[y];
			val = "";
			if(x === undefined) continue;
			else if(x.F != null) {
				y = x.F;
				if(!x.f) continue;
				val = x.f;
				if(y.indexOf(":") == -1) y = y + ":" + y;
			}
			if(x.f != null) val = x.f;
			else if(x.t == 'z') continue;
			else if(x.t == 'n' && x.v != null) val = "" + x.v;
			else if(x.t == 'b') val = x.v ? "TRUE" : "FALSE";
			else if(x.w !== undefined) val = "'" + x.w;
			else if(x.v === undefined) continue;
			else if(x.t == 's') val = "'" + x.v;
			else val = ""+x.v;
			cmds[cmds.length] = y + "=" + val;
		}
	}
	return cmds;
}

function sheet_add_json(_ws/*:?Worksheet*/, js/*:Array<any>*/, opts)/*:Worksheet*/ {
	if(!js.length) return _ws;
	var o = opts || {};
	var offset = +!o.skipHeader;
	var ws/*:Worksheet*/ = _ws || ({}/*:any*/);
	var _R = 0, _C = 0;
	if(ws && o.origin != null) {
		if(typeof o.origin == 'number') _R = o.origin;
		else {
			var _origin/*:CellAddress*/ = typeof o.origin == "string" ? decode_cell(o.origin) : o.origin;
			_R = _origin.r; _C = _origin.c;
		}
	}
	var cell/*:Cell*/;
	var range/*:Range*/ = ({s: {c:0, r:0}, e: {c:_C, r:_R + js.length - 1 + offset}}/*:any*/);
	if(ws['!ref']) {
		var _range = safe_decode_range(ws['!ref']);
		range.e.c = Math.max(range.e.c, _range.e.c);
		range.e.r = Math.max(range.e.r, _range.e.r);
		if(_R == -1) { _R = _range.e.r + 1; range.e.r = _R + js.length - 1 + offset; }
	} else {
		if(_R == -1) { _R = 0; range.e.r = js.length - 1 + offset; }
	}
	var hdr/*:Array<string>*/ = o.header || [], C = 0;

	js.forEach(function (JS, R/*:number*/) {
		keys(JS).forEach(function(k) {
			if((C=hdr.indexOf(k)) == -1) hdr[C=hdr.length] = k;
			var v = JS[k];
			var t = 'z';
			var z = "";
			var ref = encode_cell({c:_C + C,r:_R + R + offset});
			cell = utils.sheet_get_cell(ws, ref);
			if(v && typeof v === 'object' && !(v instanceof Date)){
				ws[ref] = v;
			} else {
				if(typeof v == 'number') t = 'n';
				else if(typeof v == 'boolean') t = 'b';
				else if(typeof v == 'string') t = 's';
				else if(v instanceof Date) {
					t = 'd';
					if(!o.cellDates) { t = 'n'; v = datenum(v); }
					z = (o.dateNF || SSF._table[14]);
				}
				if(!cell) ws[ref] = cell = ({t:t, v:v}/*:any*/);
				else {
					cell.t = t; cell.v = v;
					delete cell.w; delete cell.R;
					if(z) cell.z = z;
				}
				if(z) cell.z = z;
			}
		});
	});
	range.e.c = Math.max(range.e.c, _C + hdr.length - 1);
	var __R = encode_row(_R);
	if(offset) for(C = 0; C < hdr.length; ++C) ws[encode_col(C + _C) + __R] = {t:'s', v:hdr[C]};
	ws['!ref'] = encode_range(range);
	return ws;
}
function json_to_sheet(js/*:Array<any>*/, opts)/*:Worksheet*/ { return sheet_add_json(null, js, opts); }

var utils/*:any*/ = {
	encode_col: encode_col,
	encode_row: encode_row,
	encode_cell: encode_cell,
	encode_range: encode_range,
	decode_col: decode_col,
	decode_row: decode_row,
	split_cell: split_cell,
	decode_cell: decode_cell,
	decode_range: decode_range,
	format_cell: format_cell,
	get_formulae: sheet_to_formulae,
	make_csv: sheet_to_csv,
	make_json: sheet_to_json,
	make_formulae: sheet_to_formulae,
	sheet_add_aoa: sheet_add_aoa,
	sheet_add_json: sheet_add_json,
	sheet_add_dom: sheet_add_dom,
	aoa_to_sheet: aoa_to_sheet,
	json_to_sheet: json_to_sheet,
	table_to_sheet: parse_dom_table,
	table_to_book: table_to_book,
	sheet_to_csv: sheet_to_csv,
	sheet_to_txt: sheet_to_txt,
	sheet_to_json: sheet_to_json,
	sheet_to_html: HTML_.from_sheet,
	sheet_to_formulae: sheet_to_formulae,
	sheet_to_row_object_array: sheet_to_json
};

if ((typeof 'module' != 'undefined'  && typeof require != 'undefined') || (typeof $ != 'undefined')) {
  var StyleBuilder = function (options) {

    if(typeof module !== "undefined" && typeof require !== 'undefined' ) {
      var cheerio = require('cheerio');
      createElement = function(str) { return cheerio(cheerio(str, null, null, {xmlMode: true})); };
    }
    else if (typeof jQuery !== 'undefined' || typeof $ !== 'undefined') {
      createElement = function(str) { return $(str); }
    }
    else {
      createElement = function() { }
    }


    var customNumFmtId = 164;


    var table_fmt = {
      0:  'General',
      1:  '0',
      2:  '0.00',
      3:  '#,##0',
      4:  '#,##0.00',
      9:  '0%',
      10: '0.00%',
      11: '0.00E+00',
      12: '# ?/?',
      13: '# ??/??',
      14: 'm/d/yy',
      15: 'd-mmm-yy',
      16: 'd-mmm',
      17: 'mmm-yy',
      18: 'h:mm AM/PM',
      19: 'h:mm:ss AM/PM',
      20: 'h:mm',
      21: 'h:mm:ss',
      22: 'm/d/yy h:mm',
      37: '#,##0 ;(#,##0)',
      38: '#,##0 ;[Red](#,##0)',
      39: '#,##0.00;(#,##0.00)',
      40: '#,##0.00;[Red](#,##0.00)',
      45: 'mm:ss',
      46: '[h]:mm:ss',
      47: 'mmss.0',
      48: '##0.0E+0',
      49: '@',
      56: '"上午/下午 "hh"時"mm"分"ss"秒 "',
      65535: 'General'
    };
    var fmt_table = {};

    for (var idx in table_fmt) {
      fmt_table[table_fmt[idx]] = idx;

    }


    var baseXmlprefix = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    var baseXml =
        '<styleSheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"\
        xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" mc:Ignorable="x14ac">\
            <numFmts count="1">\
              <numFmt numFmtId="164" formatCode="0.00%"/>\
            </numFmts>\
            <fonts count="0" x14ac:knownFonts="1"></fonts>\
            <fills count="0"></fills>\
            <borders count="0"></borders>\
            <cellStyleXfs count="1">\
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>\
            </cellStyleXfs>\
            <cellXfs count="0"></cellXfs>\
            <cellStyles count="1">\
              <cellStyle name="Normal" xfId="0" builtinId="0"/>\
            </cellStyles>\
            <dxfs count="0"/>\
            <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>\
        </styleSheet>';

    _hashIndex = {};
    _listIndex = [];

    return {

      initialize: function (options) {
        if (typeof cheerio !== 'undefined') {
          this.$styles = cheerio.load(baseXml, {xmlMode: true});
          this.$styles.find = function(q) { return this(q)}
        }
        else {
          this.$styles = $(baseXml);
        }


        // need to specify styles at index 0 and 1.
        // the second style MUST be gray125 for some reason

        var defaultStyle = options.defaultCellStyle;
        if (!defaultStyle) defaultStyle = {
          font: {name: 'Calibri', sz: '11'},
          fill: { fgColor: { patternType: "none"}},
          border: {},
          numFmt: null
        };
        if (!defaultStyle.border) { defaultStyle.border = {}}

        var gray125Style = JSON.parse(JSON.stringify(defaultStyle));
        gray125Style.fill = { fgColor: { patternType: "gray125"}}

        this.addStyles([defaultStyle, gray125Style]);
        return this;
      },

      // create a style entry and returns an integer index that can be used in the cell .s property
      // these format of this object follows the emerging Common Spreadsheet Format
      addStyle: function (attributes) {
        var attributes = this._duckTypeStyle(attributes);
        var hashKey = JSON.stringify(attributes);
        var index = _hashIndex[hashKey];
        if (index == undefined) {
          index = this._addXf(attributes || {});
          _hashIndex[hashKey] = index;

        }
        else {
          index = _hashIndex[hashKey];
        }
        return index;
      },

      // create style entries and returns array of integer indexes that can be used in cell .s property
      addStyles: function (styles) {
        var self = this;
        return styles.map(function (style) {
          return self.addStyle(style);
        })
      },

      _duckTypeStyle: function(attributes) {

        if (typeof attributes == 'object' && (attributes.patternFill || attributes.fgColor)) {
          return {fill: attributes }; // this must be read via XLSX.parseFile(...)
        }
        else if (attributes.font || attributes.numFmt || attributes.border || attributes.fill) {
          return attributes;
        }
        else {
          return this._getStyleCSS(attributes)
        }
      },

      _getStyleCSS: function(css) {
        return css; //TODO
      },

      // Create an <xf> record for the style as well as corresponding <font>, <fill>, <border>, <numfmts>
      // Right now this is simple and creates a <font>, <fill>, <border>, <numfmts> for every <xf>
      // We could perhaps get fancier and avoid duplicating  auxiliary entries as Excel presumably intended, but bother.
      _addXf: function (attributes) {


        var fontId = this._addFont(attributes.font);
        var fillId = this._addFill(attributes.fill);
        var borderId = this._addBorder(attributes.border);
        var numFmtId = this._addNumFmt(attributes.numFmt);

        var $xf = createElement('<xf></xf>')
            .attr("numFmtId", numFmtId)
            .attr("fontId", fontId)
            .attr("fillId", fillId)
            .attr("borderId", 0)
            .attr("xfId", "0");

        if (fontId > 0) {
          $xf.attr('applyFont', "1");
        }
        if (fillId > 0) {
          $xf.attr('applyFill', "1");
        }
        if (borderId > 0) {
          $xf.attr('applyBorder', "1");
        }
        if (numFmtId > 0) {
          $xf.attr('applyNumberFormat', "1");
        }


        var $cellXfs = this.$styles.find('cellXfs');

        $cellXfs.append($xf);
        var count = +$cellXfs.attr('count') + 1;

        $cellXfs.attr('count', count);
        return count - 1;
      },

      _addFont: function (attributes) {
        if (!attributes) {
          return 0;
        }

        var $font = createElement('<font/>', null, null, {xmlMode: true});

        $font.append(createElement('<sz/>').attr('val', attributes.sz))
            .append(createElement('<color/>').attr('theme', '1'))
            .append(createElement('<name/>').attr('val', attributes.name))
//              .append(createElement('<family/>').attr('val', '2'))
//              .append(createElement('<scheme/>').attr('val', 'minor'));

        if (attributes.bold) $font.append('<b/>');
        if (attributes.underline) $font.append('<u/>');
        if (attributes.italic) $font.append('<i/>');

        if (attributes.color) {
          if (attributes.color.theme) {
            $font.append(createElement('<color/>').attr('theme', attributes.color.theme));
          } else if (attributes.color.rgb) {
            $font.append(createElement('<color/>').attr('rgb', attributes.color.rgb));
          }
        }


        var $fonts = this.$styles.find('fonts');
        $fonts.append($font);

        var count = $fonts.children().length;
        $fonts.attr('count', count);
        return count - 1;
      },

      _addNumFmt: function (numFmt) {
        if (!numFmt) {
          return 0;
        }

        if (typeof numFmt == 'string') {
          var numFmtIdx = fmt_table[numFmt];
          if (numFmtIdx >= 0) {
            return numFmtIdx; // we found a match against built in formats
          }
        }

        if (numFmt == +numFmt) {
          return numFmt; // we're matching an integer against some known code
        }

        var $numFmt = createElement('<numFmt/>', null, null, {xmlMode: true})
            .attr("numFmtId", ++customNumFmtId )
            .attr("formatCode", numFmt);

        var $numFmts = this.$styles.find('numFmts');
        $numFmts.append($numFmt);

        var count = $numFmts.children().length;
        $numFmts.attr('count', count);
        return customNumFmtId;
      },

      _addFill: function (attributes) {

        if (!attributes) {
          return 0;
        }
        var $patternFill = createElement('<patternFill></patternFill>', null, null, {xmlMode: true})
            .attr('patternType', attributes.patternType || 'solid');

        if (attributes.fgColor) {
          //Excel doesn't like it when we set both rgb and theme+tint, but xlsx.parseFile() sets both
          //var $fgColor = createElement('<fgColor/>', null, null, {xmlMode: true}).attr(attributes.fgColor)
          if (attributes.fgColor.rgb) {

            if (attributes.fgColor.rgb.length == 6) {
              attributes.fgColor.rgb = "FF" + attributes.fgColor.rgb /// add alpha to an RGB as Excel expects aRGB
            }
            var $fgColor = createElement('<fgColor/>', null, null, {xmlMode: true}).
                attr('rgb', attributes.fgColor.rgb);
            $patternFill.append($fgColor);
          }
          else if (attributes.fgColor.theme) {
            var $fgColor = createElement('<fgColor/>', null, null, {xmlMode: true});
            $fgColor.attr('theme', attributes.fgColor.theme);
            if (attributes.fgColor.tint) {
              $fgColor.attr('tint', attributes.fgColor.tint);
            }
            $patternFill.append($fgColor);
          }

          if (!attributes.bgColor) {
            attributes.bgColor = { "indexed": "64"}
          }
        }

        if (attributes.bgColor) {
          var $bgColor = createElement('<bgColor/>', null, null, {xmlMode: true}).attr(attributes.bgColor);
          $patternFill.append($bgColor);
        }

        var $fill = createElement('<fill></fill>')
            .append($patternFill);

        this.$styles.find('fills').append($fill);
        var $fills = this.$styles.find('fills')
        $fills.append($fill);

        var count = $fills.children().length;
        $fills.attr('count', count);
        return count - 1;
      },

      _addBorder: function (attributes) {
        if (!attributes) {
          return 0;
        }
        var $border = createElement('<border></border>')
            .append('<left></left>')
            .append('<right></right>')
            .append('<top></top>')
            .append('<bottom></bottom>')
            .append('<diagonal></diagonal>');

        var $borders = this.$styles.find('borders');
        $borders.append($border);

        var count = $borders.children().length;
        $borders.attr('count', count);
        return count;
      },

      toXml: function () {
        if (this.$styles.find('numFmts').children().length == 0) {
          this.$styles.find('numFmts').remove();
        }
        if (this.$styles.xml) { return this.$styles.xml(); }
        else { return baseXmlprefix + this.$styles.html(); }
      }
    }.initialize(options||{});
  }
}
