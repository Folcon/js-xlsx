var strs = {}; // shared strings
var _ssfopts = {}; // spreadsheet formatting options

RELS.WS = [
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	"http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
];

function get_sst_id(sst/*:SST*/, str/*:string*/)/*:number*/ {
	for(var i = 0, len = sst.length; i < len; ++i) if(sst[i].t === str) { sst.Count ++; return i; }
	sst[len] = {t:str}; sst.Count ++; sst.Unique ++; return len;
}

function col_obj_w(C/*:number*/, col) {
	var p = ({min:C+1,max:C+1}/*:any*/);
	/* wch (chars), wpx (pixels) */
	var wch = -1;
	if(col.MDW) MDW = col.MDW;
	if(col.width != null) p.customWidth = 1;
	else if(col.wpx != null) wch = px2char(col.wpx);
	else if(col.wch != null) wch = col.wch;
	if(wch > -1) { p.width = char2width(wch); p.customWidth = 1; }
	else if(col.width != null) p.width = col.width;
	if(col.hidden) p.hidden = true;
	return p;
}

function default_margins(margins, mode) {
	if(!margins) return;
	var defs = [0.7, 0.7, 0.75, 0.75, 0.3, 0.3];
	if(mode == 'xlml') defs = [1, 1, 1, 1, 0.5, 0.5];
	if(margins.left   == null) margins.left   = defs[0];
	if(margins.right  == null) margins.right  = defs[1];
	if(margins.top    == null) margins.top    = defs[2];
	if(margins.bottom == null) margins.bottom = defs[3];
	if(margins.header == null) margins.header = defs[4];
	if(margins.footer == null) margins.footer = defs[5];
}

function get_cell_style(styles, cell, opts) {
  if (typeof style_builder != 'undefined') {
    if (/^\d+$/.exec(cell.s)) { return cell.s}  // if its already an integer index, let it be
    if (cell.s && (cell.s == +cell.s)) { return cell.s}  // if its already an integer index, let it be
    var s = cell.s || {};
    if (cell.z) s.numFmt = cell.z;
    return style_builder.addStyle(s);
  }
  else {
    var z = opts.revssf[cell.z != null ? cell.z : "General"];
    for(var i = 0, len = styles.length; i != len; ++i) if(styles[i].numFmtId === z) return i;
    styles[len] = {
      numFmtId:z,
      fontId:0,
      fillId:0,
      borderId:0,
      xfId:0,
      applyNumberFormat:1
    };
    return len;
  }
}

function get_cell_style_csf(cellXf) {

  if (cellXf) {

    var s = {}

    if (typeof cellXf.numFmtId != undefined)  {
      s.numFmt = SSF._table[cellXf.numFmtId];
    }

    if(cellXf.fillId)  {
      s.fill =  styles.Fills[cellXf.fillId];
    }

    if (cellXf.fontId) {
      s.font = styles.Fonts[cellXf.fontId];
    }
    if (cellXf.borderId) {
      s.border = styles.Borders[cellXf.borderId];
    }
    if (cellXf.applyAlignment==1) {
      s.alignment = cellXf.alignment;
    }


    return JSON.parse(JSON.stringify(s));
  }
  return null;
}

function safe_format(p, fmtid, fillid, opts, themes, styles) {
	if(p.t === 'z') return;
	if(p.t === 'd' && typeof p.v === 'string') p.v = parseDate(p.v);
	try {
		if(opts.cellNF) p.z = SSF._table[fmtid];
	} catch(e) { if(opts.WTF) throw e; }
	if(!opts || opts.cellText !== false) try {
		if(p.t === 'e') p.w = p.w || BErr[p.v];
		else if(fmtid === 0) {
			if(p.t === 'n') {
				if((p.v|0) === p.v) p.w = SSF._general_int(p.v,_ssfopts);
				else p.w = SSF._general_num(p.v,_ssfopts);
			}
			else if(p.t === 'd') {
				var dd = datenum(p.v);
				if((dd|0) === dd) p.w = SSF._general_int(dd,_ssfopts);
				else p.w = SSF._general_num(dd,_ssfopts);
			}
			else if(p.v === undefined) return "";
			else p.w = SSF._general(p.v,_ssfopts);
		}
		else if(p.t === 'd') p.w = SSF.format(fmtid,datenum(p.v),_ssfopts);
		else p.w = SSF.format(fmtid,p.v,_ssfopts);
	} catch(e) { if(opts.WTF) throw e; }
}
