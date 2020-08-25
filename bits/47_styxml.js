function parse_fills(t, opts) {
	styles.Fills = [];
	var fill = {};
	t[0].match(tagregex).forEach(function (x) {
		var y = parsexmltag(x);
		switch (y[0]) {
			case '<fills':
			case '<fills>':
			case '</fills>':
				break;

			/* 18.8.20 fill CT_Fill */
			case '<fill>':
				break;
			case '</fill>':
				styles.Fills.push(fill);
				fill = {};
				break;

			/* 18.8.32 patternFill CT_PatternFill */
			case '<patternFill':
				if (y.patternType) fill.patternType = y.patternType;
				break;
			case '<patternFill/>':
			case '</patternFill>':
				break;

			/* 18.8.3 bgColor CT_Color */
			case '<bgColor':
				if (!fill.bgColor) fill.bgColor = {};
				if (y.indexed) fill.bgColor.indexed = parseInt(y.indexed, 10);
				if (y.theme) fill.bgColor.theme = parseInt(y.theme, 10);
				if (y.tint) fill.bgColor.tint = parseFloat(y.tint);


				if (y.theme && themes.themeElements && themes.themeElements.clrScheme) {
					fill.bgColor.rgb = rgb_tint(themes.themeElements.clrScheme[fill.bgColor.theme].rgb, fill.bgColor.tint || 0);
					if (opts.WTF) fill.bgColor.raw_rgb = rgb_tint(themes.themeElements.clrScheme[fill.bgColor.theme].rgb, 0);
				}
				/* Excel uses ARGB strings */
				if (y.rgb) fill.bgColor.rgb = y.rgb;//.substring(y.rgb.length - 6);
				break;
			case '<bgColor/>':
			case '</bgColor>':
				break;

			/* 18.8.19 fgColor CT_Color */
			case '<fgColor':
				if (!fill.fgColor) fill.fgColor = {};
				if (y.theme) fill.fgColor.theme = parseInt(y.theme, 10);
				if (y.tint) fill.fgColor.tint = parseFloat(y.tint);

				if (y.theme && themes.themeElements && themes.themeElements.clrScheme) {
					fill.fgColor.rgb = rgb_tint(themes.themeElements.clrScheme[fill.fgColor.theme].rgb, fill.fgColor.tint || 0);
					if (opts.WTF) fill.fgColor.raw_rgb = rgb_tint(themes.themeElements.clrScheme[fill.fgColor.theme].rgb, 0);
				}

				/* Excel uses ARGB strings */
				if (y.rgb) fill.fgColor.rgb = y.rgb;//.substring(y.rgb.length - 6);
				break;
			case '<fgColor/>':
			case '</fgColor>':
				break;

			default:
				if (opts.WTF) throw 'unrecognized ' + y[0] + ' in fills';
		}
	});
}

function parse_fonts(t, opts) {
	styles.Fonts = [];
	var font = {};
	t[0].match(tagregex).forEach(function (x) {
		var y = parsexmltag(x);
		switch (y[0]) {

			case '<fonts':
			case '<fonts>':
			case '</fonts>':
				break;
			case '<font':
				break;
			case '</font>':
				styles.Fonts.push(font);
				;
				font = {};
				break;

			case '<name':
				if (y.val) font.name = y.val;
				break;
			case '<name/>':
			case '</name>':
				break;


			case '<b/>':
				font.bold = true;
				break;
			case '<u/>':
				font.underline = true;
				break;
			case '<i/>':
				font.italic = true;
				break;
			case '<strike/>':
				font.strike = true;
				break;
			case '<outline/>':
				font.outline = true;
				break;
			case '<shadow/>':
				font.shadow = true;
				break;


			case '<sz':
				if (y.val) font.sz = y.val;
				break;
			case '<sz/>':
			case '</sz>':
				break;

			case '<vertAlign':
				if (y.val) font.vertAlign = y.val;
				break;
			case '<vertAlign/>':
			case '</vertAlign>':
				break;


			case '<color':
				if (!font.color) font.color = {};
				if (y.theme) font.color.theme = y.theme;
				if (y.tint) font.color.tint = y.tint;
				if (y.theme && themes.themeElements && themes.themeElements.clrScheme) {
					font.color.rgb = rgb_tint(themes.themeElements.clrScheme[font.color.theme].rgb, font.color.tint || 0);
				}
				if (y.rgb) font.color.rgb = y.rgb;
				break;
			case '<color/>':
			case '</color>':
				break;
		}
	});
}

function parse_borders(t, opts) {
	styles.Borders = [];
	var border = {}, sub_border = {};
	t[0].match(tagregex).forEach(function (x) {
		var y = parsexmltag(x);
		switch (y[0]) {
			case '<borders':
			case '<borders>':
			case '</borders>':
				break;
			case '<border':
			case '<border>':
				border = {};
				if (y.diagonalUp) { border.diagonalUp = y.diagonalUp; }
				if (y.diagonalDown) { border.diagonalDown = y.diagonalDown; }
				styles.Borders.push(border);

				break;
				break;
			case '</border>':
				break;

			case '<left':
				sub_border = border.left = {};
				if (y.style) {
					sub_border.style = y.style;
				}
				break;
			case '<right':
				sub_border = border.right = {};
				if (y.style) {
					sub_border.style = y.style;
				}
				break;
			case '<top':
				sub_border = border.top = {};
				if (y.style) {
					sub_border.style = y.style;
				}
				break;
			case '<bottom':
				sub_border = border.bottom = {};
				if (y.style) {
					sub_border.style = y.style;
				}
				break;
			case '<diagonal':
				sub_border = border.diagonal = {};
				if (y.style) {
					sub_border.style = y.style;
				}
				break;

			case '<color':
				sub_border.color = {};
				if (y.theme) sub_border.color.theme = y.theme;
				if (y.theme && themes.themeElements && themes.themeElements.clrScheme) {
					sub_border.color.rgb = rgb_tint(themes.themeElements.clrScheme[sub_border.color.theme].rgb, sub_border.color.tint || 0);
				}

				if (y.tint) sub_border.color.tint = y.tint;
				if (y.rgb) sub_border.color.rgb = y.rgb;
				if (y.auto) sub_border.color.auto = y.auto;
				break;
			case '<name/>':
			case '</name>':
				break;
			default:
				break;
		}
	});

}

/* 18.8.31 numFmts CT_NumFmts */
function parse_numFmts(t, opts) {
	styles.NumberFmt = [];
	var k = keys(SSF._table);
	for (var i = 0; i < k.length; ++i) styles.NumberFmt[k[i]] = SSF._table[k[i]];
	var m = t[0].match(tagregex);
	for (i = 0; i < m.length; ++i) {
		var y = parsexmltag(m[i]);
		switch (y[0]) {
			case '<numFmts':
			case '</numFmts>':
			case '<numFmts/>':
			case '<numFmts>':
				break;
			case '<numFmt':
				{
					var f = unescapexml(utf8read(y.formatCode)), j = parseInt(y.numFmtId, 10);
					styles.NumberFmt[j] = f;
					if (j > 0) SSF.load(f, j);
				}
				break;
			default:
				if (opts.WTF) throw 'unrecognized ' + y[0] + ' in numFmts';
		}
	}
}

function write_numFmts(NF, opts) {
	var o = ["<numFmts>"];
	[
		[5, 8],
		[23, 26],
		[41, 44],
		[63, 66],
		[164, 392]
	].forEach(function (r) {
		for (var i = r[0]; i <= r[1]; ++i) if (NF[i] !== undefined) o[o.length] = (writextag('numFmt', null, { numFmtId: i, formatCode: escapexml(NF[i]) }));
	});
	if (o.length === 1) return "";
	o[o.length] = ("</numFmts>");
	o[0] = writextag('numFmts', null, { count: o.length - 2 }).replace("/>", ">");
	return o.join("");
}

/* 18.8.10 cellXfs CT_CellXfs */
function parse_cellXfs(t, opts) {
	styles.CellXf = [];
	var xf;
	t[0].match(tagregex).forEach(function (x) {
		var y = parsexmltag(x);
		switch (y[0]) {
			case '<cellXfs':
			case '<cellXfs>':
			case '<cellXfs/>':
			case '</cellXfs>':
				break;

			/* 18.8.45 xf CT_Xf */
			case '<xf':
				xf = y;
				delete xf[0];
				delete y[0];
				if (xf.numFmtId) xf.numFmtId = parseInt(xf.numFmtId, 10);
				if (xf.fillId) xf.fillId = parseInt(xf.fillId, 10);
				styles.CellXf.push(xf);
				break;
			case '</xf>':
				break;

			/* 18.8.1 alignment CT_CellAlignment */
			case '<alignment':
			case '<alignment/>':
				var alignment = {}
				if (y.vertical) { alignment.vertical = y.vertical; }
				if (y.horizontal) { alignment.horizontal = y.horizontal; }
				if (y.textRotation != undefined) { alignment.textRotation = y.textRotation; }
				if (y.indent) { alignment.indent = y.indent; }
				if (y.wrapText) { alignment.wrapText = y.wrapText; }
				xf.alignment = alignment;

				break;

			/* 18.8.33 protection CT_CellProtection */
			case '<protection':
			case '</protection>':
			case '<protection/>':
				break;

			case '<extLst':
			case '</extLst>':
				break;
			case '<ext':
				break;
			default:
				if (opts.WTF) throw 'unrecognized ' + y[0] + ' in cellXfs';
		}
	});
}

function write_cellXfs(cellXfs) {
	var o = [];
	o[o.length] = (writextag('cellXfs', null));
	cellXfs.forEach(function (c) {
		o[o.length] = (writextag('xf', null, c));
	});
	o[o.length] = ("</cellXfs>");
	if (o.length === 2) return "";
	o[0] = writextag('cellXfs', null, { count: o.length - 2 }).replace("/>", ">");
	return o.join("");
}

/* 18.8 Styles CT_Stylesheet*/
var parse_sty_xml = (function make_pstyx() {
	var numFmtRegex = /<numFmts([^>]*)>.*<\/numFmts>/;
	var cellXfRegex = /<cellXfs([^>]*)>.*<\/cellXfs>/;
	var fillsRegex = /<fills([^>]*)>.*<\/fills>/;
	var bordersRegex = /<borders([^>]*)>.*<\/borders>/;

	return function parse_sty_xml(data, opts) {
		/* 18.8.39 styleSheet CT_Stylesheet */
		var t;

		/* numFmts CT_NumFmts ? */
		if ((t = data.match(numFmtRegex))) parse_numFmts(t, opts);

		/* fonts CT_Fonts ? */
		if ((t = data.match(/<fonts([^>]*)>.*<\/fonts>/))) parse_fonts(t, opts)

		/* fills CT_Fills */
		if ((t = data.match(fillsRegex))) parse_fills(t, opts);

		/* borders CT_Borders ? */
		if ((t = data.match(bordersRegex))) parse_borders(t, opts);
		/* cellStyleXfs CT_CellStyleXfs ? */

		/* cellXfs CT_CellXfs ? */
		if ((t = data.match(cellXfRegex))) parse_cellXfs(t, opts);

		/* dxfs CT_Dxfs ? */
		/* tableStyles CT_TableStyles ? */
		/* colors CT_Colors ? */
		/* extLst CT_ExtensionList ? */

		return styles;
	};
})();

var STYLES_XML_ROOT = writextag('styleSheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:vt': XMLNS.vt
});

RELS.STY = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

function write_sty_xml(wb, opts) {

	if (typeof style_builder != 'undefined' && typeof 'require' != 'undefined') {
		return style_builder.toXml();
	}

	var o = [XML_HEADER, STYLES_XML_ROOT], w;
	if ((w = write_numFmts(wb.SSF)) != null) o[o.length] = w;
	o[o.length] = ('<fonts count="1"><font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>');
	o[o.length] = ('<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>');
	o[o.length] = ('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');
	o[o.length] = ('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
	if ((w = write_cellXfs(opts.cellXfs))) o[o.length] = (w);
	o[o.length] = ('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
	o[o.length] = ('<dxfs count="0"/>');
	o[o.length] = ('<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>');

	if (o.length > 2) {
		o[o.length] = ('</styleSheet>');
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}