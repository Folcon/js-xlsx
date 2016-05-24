function parse_ws_xml_dim(ws/*:Worksheet*/, s/*:string*/) {
	var d = safe_decode_range(s);
	if(d.s.r<=d.e.r && d.s.c<=d.e.c && d.s.r>=0 && d.s.c>=0) ws["!ref"] = encode_range(d);
}
var mergecregex = /<(?:\w:)?mergeCell ref="[A-Z0-9:]+"\s*[\/]?>/g;
var sheetdataregex = /<(?:\w+:)?sheetData[^>]*>([\s\S]*)<\/(?:\w+:)?sheetData>/;
var hlinkregex = /<(?:\w:)?hyperlink [^>]*>/mg;
var dimregex = /"(\w*:\w*)"/;
var colregex = /<(?:\w:)?col\b[^>]*[\/]?>/g;
var afregex = /<(?:\w:)?autoFilter[^>]*([\/]|>([\s\S]*)<\/(?:\w:)?autoFilter)>/g;
var marginregex= /<(?:\w:)?pageMargins[^>]*\/>/g;
var sheetprregex = /<(?:\w:)?sheetPr\b(?:[^>a-z][^>]*)?\/>/;
var svsregex = /<(?:\w:)?sheetViews[^>]*(?:[\/]|>([\s\S]*)<\/(?:\w:)?sheetViews)>/;

/* 18.3 Worksheets */
function parse_ws_xml(data/*:?string*/, opts, idx/*:number*/, rels, wb/*:WBWBProps*/, themes, styles)/*:Worksheet*/ {
	if(!data) return data;
	if(!rels) rels = {'!id':{}};
	if(DENSE != null && opts.dense == null) opts.dense = DENSE;

	/* 18.3.1.99 worksheet CT_Worksheet */
	var s = opts.dense ? ([]/*:any*/) : ({}/*:any*/);
	var refguess/*:Range*/ = ({s: {r:2000000, c:2000000}, e: {r:0, c:0} }/*:any*/);

	var data1 = "", data2 = "";
	var mtch/*:?any*/ = data.match(sheetdataregex);
	if(mtch) {
		data1 = data.slice(0, mtch.index);
		data2 = data.slice(mtch.index + mtch[0].length);
	} else data1 = data2 = data;

	/* 18.3.1.82 sheetPr CT_SheetPr */
	var sheetPr = data1.match(sheetprregex);
	if(sheetPr) parse_ws_xml_sheetpr(sheetPr[0], s, wb, idx);

	/* 18.3.1.35 dimension CT_SheetDimension */
	var ridx = (data1.match(/<(?:\w*:)?dimension/)||{index:-1}).index;
	if(ridx > 0) {
		var ref = data1.slice(ridx,ridx+50).match(dimregex);
		if(ref) parse_ws_xml_dim(s, ref[1]);
	}

	/* 18.3.1.88 sheetViews CT_SheetViews */
	var svs = data1.match(svsregex);
	if(svs && svs[1]) parse_ws_xml_sheetviews(svs[1], wb);

	/* 18.3.1.17 cols CT_Cols */
	var columns/*:Array<ColInfo>*/ = [];
	if(opts.cellStyles) {
		/* 18.3.1.13 col CT_Col */
		var cols = data1.match(colregex);
		if(cols) parse_ws_xml_cols(columns, cols);
	}

	/* 18.3.1.80 sheetData CT_SheetData ? */
	if(mtch) parse_ws_xml_data(mtch[1], s, opts, refguess, themes, styles);

	/* 18.3.1.2  autoFilter CT_AutoFilter */
	var afilter = data2.match(afregex);
	if(afilter) s['!autofilter'] = parse_ws_xml_autofilter(afilter[0]);

	/* 18.3.1.55 mergeCells CT_MergeCells */
	var merges/*:Array<Range>*/ = [];
	var _merge = data2.match(mergecregex);
	if(_merge) for(ridx = 0; ridx != _merge.length; ++ridx)
		merges[ridx] = safe_decode_range(_merge[ridx].slice(_merge[ridx].indexOf("\"")+1));

	/* 18.3.1.48 hyperlinks CT_Hyperlinks */
	var hlink = data2.match(hlinkregex);
	if(hlink) parse_ws_xml_hlinks(s, hlink, rels);

	/* 18.3.1.62 pageMargins CT_PageMargins */
	var margins = data2.match(marginregex);
	if(margins) s['!margins'] = parse_ws_xml_margins(parsexmltag(margins[0]));

	if(!s["!ref"] && refguess.e.c >= refguess.s.c && refguess.e.r >= refguess.s.r) s["!ref"] = encode_range(refguess);
	if(opts.sheetRows > 0 && s["!ref"]) {
		var tmpref = safe_decode_range(s["!ref"]);
		if(opts.sheetRows <= +tmpref.e.r) {
			tmpref.e.r = opts.sheetRows - 1;
			if(tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
			if(tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
			if(tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
			if(tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
			s["!fullref"] = s["!ref"];
			s["!ref"] = encode_range(tmpref);
		}
	}
	if(columns.length > 0) s["!cols"] = columns;
	if(merges.length > 0) s["!merges"] = merges;
	return s;
}

function write_ws_xml_merges(merges/*:Array<Range>*/)/*:string*/ {
	if(merges.length === 0) return "";
	var o = '<mergeCells count="' + merges.length + '">';
	for(var i = 0; i != merges.length; ++i) o += '<mergeCell ref="' + encode_range(merges[i]) + '"/>';
	return o + '</mergeCells>';
}

function write_ws_xml_pagesetup(setup) {
  var pageSetup =  writextag('pageSetup', null, {
    scale: setup.scale || '100',
    orientation: setup.orientation || 'portrait',
    horizontalDpi : setup.horizontalDpi || '4294967292',
    verticalDpi : setup.verticalDpi || '4294967292'
  })
  return pageSetup;
}


function parse_ws_xml_hlinks(s, data, rels) {
	for(var i = 0; i != data.length; ++i) {
		var val = parsexmltag(utf8read(data[i]), true);
		if(!val.ref) return;
		var rel = ((rels || {})['!id']||[])[val.id];
		if(rel) {
			val.Target = rel.Target;
			if(val.location) val.Target += "#"+val.location;
		} else {
			val.Target = "#" + val.location;
			rel = {Target: val.Target, TargetMode: 'Internal'};
		}
		val.Rel = rel;
		if(val.tooltip) { val.Tooltip = val.tooltip; delete val.tooltip; }
		var rng = safe_decode_range(val.ref);
		for(var R=rng.s.r;R<=rng.e.r;++R) for(var C=rng.s.c;C<=rng.e.c;++C) {
			var addr = encode_cell({c:C,r:R});
			if(dense) {
				if(!s[R]) s[R] = [];
				if(!s[R][C]) s[R][C] = {t:"z",v:undefined};
				s[R][C].l = val;
			} else {
				if(!s[addr]) s[addr] = {t:"z",v:undefined};
				s[addr].l = val;
			}
		}
	}
}

function parse_ws_xml_margins(margin) {
	var o = {};
	["left", "right", "top", "bottom", "header", "footer"].forEach(function(k) {
		if(margin[k]) o[k] = parseFloat(margin[k]);
	});
	return o;
}
function write_ws_xml_margins(margin)/*:string*/ {
	default_margins(margin);
	return writextag('pageMargins', null, margin);
}

function parse_ws_xml_cols(columns, cols) {
	var seencol = false;
	for(var coli = 0; coli != cols.length; ++coli) {
		var coll = parsexmltag(cols[coli], true);
		if(coll.hidden) coll.hidden = parsexmlbool(coll.hidden);
		var colm=parseInt(coll.min, 10)-1, colM=parseInt(coll.max,10)-1;
		delete coll.min; delete coll.max; coll.width = +coll.width;
		if(!seencol && coll.width) { seencol = true; find_mdw_colw(coll.width); }
		process_col(coll);
		while(colm <= colM) columns[colm++] = dup(coll);
	}
}
function write_ws_xml_cols(ws, cols)/*:string*/ {
	var o = ["<cols>"], col;
	for(var i = 0; i != cols.length; ++i) {
		if(!(col = cols[i])) continue;
		o[o.length] = (writextag('col', null, col_obj_w(i, col)));
	}
	o[o.length] = "</cols>";
	return o.join("");
}

function write_ws_xml_cell(cell, ref, ws, opts, idx, wb) {
	if(cell.v === undefined && cell.s === undefined) return "";
	var vv = "";
	var oldt = cell.t, oldv = cell.v;
	if(cell.t !== "z") switch(cell.t) {
		case 'b': vv = cell.v ? "1" : "0"; break;
		case 'n': vv = ''+cell.v; break;
		case 'e': vv = BErr[cell.v]; break;
		case 'd':
			if(opts && opts.cellDates) vv = parseDate(cell.v, -1).toISOString();
			else {
				cell = dup(cell);
				cell.t = 'n';
				vv = ''+(cell.v = datenum(parseDate(cell.v)));
			}
			if(typeof cell.z === 'undefined') cell.z = SSF._table[14];
			break;
		default: vv = cell.v; break;
	}
	var v = writetag('v', escapexml(vv)), o = ({r:ref}/*:any*/);
	/* TODO: cell style */
	var os = get_cell_style(opts.cellXfs, cell, opts);
	if(os !== 0) o.s = os;
	switch(cell.t) {
		case 'n': break;
		case 'd': o.t = "d"; break;
		case 'b': o.t = "b"; break;
		case 'e': o.t = "e"; break;
		case 'z': break;
		default: if(cell.v == null) { delete cell.t; break; }
			if(opts && opts.bookSST) {
				v = writetag('v', ''+get_sst_id(opts.Strings, cell.v, opts.revStrings));
				o.t = "s"; break;
			}
			o.t = "str"; break;
	}
	if(cell.t != oldt) { cell.t = oldt; cell.v = oldv; }
	if(cell.f) {
		var ff = cell.F && cell.F.slice(0, ref.length) == ref ? {t:"array", ref:cell.F} : null;
		v = writextag('f', escapexml(cell.f), ff) + (cell.v != null ? v : "");
	}
	if(cell.l) ws['!links'].push([ref, cell.l]);
	if(cell.c) ws['!comments'].push([ref, cell.c]);
	return writextag('c', v, o);
}

var parse_ws_xml_data = (function() {
	var cellregex = /<(?:\w+:)?c[ >]/, rowregex = /<\/(?:\w+:)?row>/;
	var rregex = /r=["']([^"']*)["']/, isregex = /<(?:\w+:)?is>([\S\s]*?)<\/(?:\w+:)?is>/;
	var refregex = /ref=["']([^"']*)["']/;
	var match_v = matchtag("v"), match_f = matchtag("f");

return function parse_ws_xml_data(sdata/*:string*/, s, opts, guess/*:Range*/, themes, styles) {
	var ri = 0, x = "", cells/*:Array<string>*/ = [], cref/*:?Array<string>*/ = [], idx=0, i=0, cc=0, d="", p/*:any*/;
	var tag, tagr = 0, tagc = 0;
	var sstr, ftag;
	var fmtid = 0, fillid = 0;
	var do_format = Array.isArray(styles.CellXf), cf;
	var arrayf/*:Array<[Range, string]>*/ = [];
	var sharedf = [];
	var dense = Array.isArray(s);
	var rows/*:Array<RowInfo>*/ = [], rowobj = {}, rowrite = false;
	for(var marr = sdata.split(rowregex), mt = 0, marrlen = marr.length; mt != marrlen; ++mt) {
		x = marr[mt].trim();
		var xlen = x.length;
		if(xlen === 0) continue;

		/* 18.3.1.73 row CT_Row */
		for(ri = 0; ri < xlen; ++ri) if(x.charCodeAt(ri) === 62) break; ++ri;
		tag = parsexmltag(x.slice(0,ri), true);
		tagr = tag.r != null ? parseInt(tag.r, 10) : tagr+1; tagc = -1;
		if(opts.sheetRows && opts.sheetRows < tagr) continue;
		if(guess.s.r > tagr - 1) guess.s.r = tagr - 1;
		if(guess.e.r < tagr - 1) guess.e.r = tagr - 1;

		if(opts && opts.cellStyles) {
			rowobj = {}; rowrite = false;
			if(tag.ht) { rowrite = true; rowobj.hpt = parseFloat(tag.ht); rowobj.hpx = pt2px(rowobj.hpt); }
			if(tag.hidden == "1") { rowrite = true; rowobj.hidden = true; }
			if(tag.outlineLevel != null) { rowrite = true; rowobj.level = +tag.outlineLevel; }
			if(rowrite) rows[tagr-1] = rowobj;
		}

		/* 18.3.1.4 c CT_Cell */
		cells = x.slice(ri).split(cellregex);
		for(var rslice = 0; rslice != cells.length; ++rslice) if(cells[rslice].trim().charAt(0) != "<") break;
		cells = cells.slice(rslice);
		for(ri = 0; ri != cells.length; ++ri) {
			x = cells[ri].trim();
			if(x.length === 0) continue;
			cref = x.match(rregex); idx = ri; i=0; cc=0;
			x = "<c " + (x.slice(0,1)=="<"?">":"") + x;
			if(cref != null && cref.length === 2) {
				idx = 0; d=cref[1];
				for(i=0; i != d.length; ++i) {
					if((cc=d.charCodeAt(i)-64) < 1 || cc > 26) break;
					idx = 26*idx + cc;
				}
				--idx;
				tagc = idx;
			} else ++tagc;
			for(i = 0; i != x.length; ++i) if(x.charCodeAt(i) === 62) break; ++i;
			tag = parsexmltag(x.substr(0,i), true);
			if(!tag.r) tag.r = utils.encode_cell({r:tagr-1, c:tagc});
			d = x.substr(i);
			p = {t:""};

			if((cref=d.match(match_v))!== null && cref[1] !== '') p.v=unescapexml(cref[1]);
			if(opts.cellFormula && (cref=d.match(match_f))!== null) p.f=unescapexml(cref[1]);

			/* SCHEMA IS ACTUALLY INCORRECT HERE.  IF A CELL HAS NO T, EMIT "" */
			if(tag.t === undefined && tag.s === undefined && p.v === undefined) {
				if(!opts.sheetStubs) continue;
				p.t = "stub";
			}
			else p.t = tag.t || "n";
			if(guess.s.c > tagc) guess.s.c = tagc;
			if(guess.e.c < tagc) guess.e.c = tagc;
			/* 18.18.11 t ST_CellType */
			switch(p.t) {
				case 'n':
          p.v = parseFloat(p.v);
          if(isNaN(p.v)) p.v = "" // we don't want NaN if p.v is null
          break;
				case 's':
					if(typeof p.v == 'undefined') {
						if(!opts.sheetStubs) continue;
						p.t = 'z';
					} else {
						sstr = strs[parseInt(p.v, 10)];
						p.v = sstr.t;
						p.r = sstr.r;
						if(opts.cellHTML) p.h = sstr.h;
					}
					break;
				case 'str':
					p.t = "s";
					p.v = (p.v!=null) ? utf8read(p.v) : '';
					if(opts.cellHTML) p.h = escapehtml(p.v);
					break;
				case 'inlineStr':
					cref = d.match(isregex);
					p.t = 's';
					if(cref != null && (sstr = parse_si(cref[1]))) {
						p.v = sstr.t;
						if(opts.cellHTML) p.h = sstr.h;
					} else p.v = "";
					break;
				case 'b': p.v = parsexmlbool(p.v); break;
				case 'd':
					if(opts.cellDates) p.v = parseDate(p.v, 1);
					else { p.v = datenum(parseDate(p.v, 1)); p.t = 'n'; }
					break;
				/* error string in .w, number in .v */
				case 'e':
					if(!opts || opts.cellText !== false) p.w = p.v;
					p.v = RBErr[p.v]; break;
			}
            /* formatting */
            fmtid = fillid = 0;
            if(do_format && tag.s !== undefined) {
              cf = styles.CellXf[tag.s];
              if (opts.cellStyles) {
                p.s = get_cell_style_csf(cf)
              }
              if(cf != null) {
                if(cf.numFmtId != null) fmtid = cf.numFmtId;
                if(opts.cellStyles && cf.fillId != null) fillid = cf.fillId;
              }
            }
            safe_format(p, fmtid, fillid, opts);
            s[tag.r] = p;
      }
	}
	if(rows.length > 0) s['!rows'] = rows;
}; })();

function write_ws_xml_data(ws/*:Worksheet*/, opts, idx/*:number*/, wb/*:Workbook*//*::, rels*/)/*:string*/ {
	var o/*:Array<string>*/ = [], r/*:Array<string>*/ = [], range = safe_decode_range(ws['!ref']), cell="", ref, rr = "", cols/*:Array<string>*/ = [], R=0, C=0, rows = ws['!rows'];
	var dense = Array.isArray(ws);
	var params = ({r:rr}/*:any*/), row/*:RowInfo*/, height = -1;
	for(C = range.s.c; C <= range.e.c; ++C) cols[C] = encode_col(C);
	for(R = range.s.r; R <= range.e.r; ++R) {
		r = [];
		rr = encode_row(R);
		for(C = range.s.c; C <= range.e.c; ++C) {
			ref = cols[C] + rr;
			var _cell = dense ? (ws[R]||[])[C]: ws[ref];
			if(_cell === undefined) continue;
			if((cell = write_ws_xml_cell(_cell, ref, ws, opts, idx, wb)) != null) r.push(cell);
		}
		if(r.length > 0 || (rows && rows[R])) {
			params = ({r:rr}/*:any*/);
			if(rows && rows[R]) {
				row = rows[R];
				if(row.hidden) params.hidden = 1;
				height = -1;
				if(row.hpx) height = px2pt(row.hpx);
				else if(row.hpt) height = row.hpt;
				if(height > -1) { params.ht = height; params.customHeight = 1; }
				if(row.level) { params.outlineLevel = row.level; }
			}
			o[o.length] = (writextag('row', r.join(""), params));
		}
	}
	if(rows) for(; R < rows.length; ++R) {
		if(rows && rows[R]) {
			params = ({r:R+1}/*:any*/);
			row = rows[R];
			if(row.hidden) params.hidden = 1;
			height = -1;
			if (row.hpx) height = px2pt(row.hpx);
			else if (row.hpt) height = row.hpt;
			if (height > -1) { params.ht = height; params.customHeight = 1; }
			if (row.level) { params.outlineLevel = row.level; }
			o[o.length] = (writextag('row', "", params));
		}
	}
	return o.join("");
}

var WS_XML_ROOT = writextag('worksheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:r': XMLNS.r
});

function write_ws_xml(idx/*:number*/, opts, wb/*:Workbook*/, rels)/*:string*/ {
	var o = [XML_HEADER, WS_XML_ROOT];
	var s = wb.SheetNames[idx], sidx = 0, rdata = "";
	var ws = wb.Sheets[s];
	if(ws == null) ws = {};
	var ref = ws['!ref'] || 'A1';
	var range = safe_decode_range(ref);
	if(range.e.c > 0x3FFF || range.e.r > 0xFFFFF) {
		if(opts.WTF) throw new Error("Range " + ref + " exceeds format limit A1:XFD1048576");
		range.e.c = Math.min(range.e.c, 0x3FFF);
		range.e.r = Math.min(range.e.c, 0xFFFFF);
		ref = encode_range(range);
	}
	if(!rels) rels = {};
	ws['!comments'] = [];
	var _drawing = [];

	write_ws_xml_sheetpr(ws, wb, idx, opts, o);

	o[o.length] = (writextag('dimension', null, {'ref': ref}));

  var sheetView = writextag('sheetView', null,  {
    showGridLines: opts.showGridLines == false ? '0' : '1',
    tabSelected: opts.tabSelected === undefined ? '0' :  opts.tabSelected,
    workbookViewId: opts.workbookViewId === undefined ? '0' : opts.workbookViewId
  });
  o[o.length] = writextag('sheetViews', sheetView);

	if(ws['!cols'] !== undefined && ws['!cols'].length > 0) o[o.length] = (write_ws_xml_cols(ws, ws['!cols']));
	o[sidx = o.length] = '<sheetData/>';
	ws['!links'] = [];
	if(ws['!ref'] != null) {
		rdata = write_ws_xml_data(ws, opts, idx, wb, rels);
		if(rdata.length > 0) o[o.length] = (rdata);
	}
	if(o.length>sidx+1) { o[o.length] = ('</sheetData>'); o[sidx]=o[sidx].replace("/>",">"); }

	/* sheetCalcPr */

	if(ws['!protect'] != null) o[o.length] = write_ws_xml_protection(ws['!protect']);

	/* protectedRanges */
	/* scenarios */

	if(ws['!autofilter'] != null) o[o.length] = write_ws_xml_autofilter(ws['!autofilter'], ws, wb, idx);

	/* sortState */
	/* dataConsolidate */
	/* customSheetViews */

	if(ws['!merges'] != null && ws['!merges'].length > 0) o[o.length] = (write_ws_xml_merges(ws['!merges']));

	/* phoneticPr */
	/* conditionalFormatting */
	/* dataValidations */

	var relc = -1, rel, rId = -1;
	if(/*::(*/ws['!links']/*::||[])*/.length > 0) {
		o[o.length] = "<hyperlinks>";
		/*::(*/ws['!links']/*::||[])*/.forEach(function(l) {
			if(!l[1].Target) return;
			rel = ({"ref":l[0]}/*:any*/);
			if(l[1].Target.charAt(0) != "#") {
				rId = add_rels(rels, -1, escapexml(l[1].Target).replace(/#.*$/, ""), RELS.HLINK);
				rel["r:id"] = "rId"+rId;
			}
			if((relc = l[1].Target.indexOf("#")) > -1) rel.location = escapexml(l[1].Target.slice(relc+1));
			if(l[1].Tooltip) rel.tooltip = escapexml(l[1].Tooltip);
			o[o.length] = writextag("hyperlink",null,rel);
		});
		o[o.length] = "</hyperlinks>";
	}
	delete ws['!links'];

	/* printOptions */

	if(ws['!margins'] != null) o[o.length] =  write_ws_xml_margins(ws['!margins']);

	/* pageSetup */
	/* headerFooter */
	/* rowBreaks */
	/* colBreaks */
	/* customProperties */
	/* cellWatches */

	if(!opts || opts.ignoreEC || (opts.ignoreEC == (void 0))) o[o.length] = writetag("ignoredErrors", writextag("ignoredError", null, {numberStoredAsText:1, sqref:ref}));

	/* smartTags */

	if(_drawing.length > 0) {
		rId = add_rels(rels, -1, "../drawings/drawing" + (idx+1) + ".xml", RELS.DRAW);
		o[o.length] = writextag("drawing", null, {"r:id":"rId" + rId});
		ws['!drawing'] = _drawing;
	}

	if(ws['!comments'].length > 0) {
		rId = add_rels(rels, -1, "../drawings/vmlDrawing" + (idx+1) + ".vml", RELS.VML);
		o[o.length] = writextag("legacyDrawing", null, {"r:id":"rId" + rId});
		ws['!legacy'] = rId;
	}

	/* legacyDrawingHF */
	/* picture */
	/* oleObjects */
	/* controls */
	/* webPublishItems */
	/* tableParts */
	/* extLst */

  if (ws['!pageSetup'] !== undefined) o[o.length] =  write_ws_xml_pagesetup(ws['!pageSetup']);
  if (ws['!rowBreaks'] !== undefined) o[o.length] =  write_ws_xml_row_breaks(ws['!rowBreaks']);
  if (ws['!colBreaks'] !== undefined) o[o.length] =  write_ws_xml_col_breaks(ws['!colBreaks']);


  if (ws['!rowBreaks'] !== undefined) o[o.length] =  write_ws_xml_row_breaks(ws['!rowBreaks'])
  if (ws['!colBreaks'] !== undefined) o[o.length] =  write_ws_xml_col_breaks(ws['!colBreaks'])

	if(o.length>2) { o[o.length] = ('</worksheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}

function write_ws_xml_row_breaks(breaks) {
  console.log("Writing breaks")
  var brk = [];
  for (var i=0; i<breaks.length; i++) {
    var thisBreak = ''+ (breaks[i]);
    var nextBreak = '' + (breaks[i+1] || '16383');
    brk.push(writextag('brk', null, {id: thisBreak, max: nextBreak, man: '1'}))
  }
  return writextag('rowBreaks', brk.join(' '), {count: brk.length, manualBreakCount: brk.length})
}
function write_ws_xml_col_breaks(breaks) {
  console.log("Writing breaks");
  var brk = [];
  for (var i=0; i<breaks.length; i++) {
    var thisBreak = ''+ (breaks[i]);
    var nextBreak = '' + (breaks[i+1] || '16383');
    brk.push(writextag('brk', null, {id: thisBreak, max: nextBreak, man: '1'}))
  }
  return writextag('colBreaks', brk.join(' '), {count: brk.length, manualBreakCount: brk.length})
}