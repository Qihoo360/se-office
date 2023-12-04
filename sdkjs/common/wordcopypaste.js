/*
 * (c) Copyright Ascensio System SIA 2010-2023
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(/**
* @param {Window} window
* @param {undefined} undefined
*/
function(window, undefined) {

// Import
var prot;
var g_fontApplication = AscFonts.g_fontApplication;
var CFont = AscFonts.CFont;

var AscBrowser = AscCommon.AscBrowser;
var align_Right = AscCommon.align_Right;
var align_Left = AscCommon.align_Left;
var align_Center = AscCommon.align_Center;
var align_Justify = AscCommon.align_Justify;
var g_oDocumentUrls = AscCommon.g_oDocumentUrls;
var History = AscCommon.History;
var pptx_content_loader = AscCommon.pptx_content_loader;
var pptx_content_writer = AscCommon.pptx_content_writer;
var g_dKoef_pix_to_mm = AscCommon.g_dKoef_pix_to_mm;
var g_dKoef_mm_to_pix = AscCommon.g_dKoef_mm_to_pix;

var CShape = AscFormat.CShape;
var CGraphicFrame = AscFormat.CGraphicFrame;

var c_oAscError = Asc.c_oAscError;
var c_oAscShdClear = Asc.c_oAscShdClear;
var c_oAscShdNil = Asc.c_oAscShdNil;
var c_oAscXAlign = Asc.c_oAscXAlign;

function CDocumentReaderMode()
{
    this.DefaultFontSize = 12;

    this.CorrectDefaultFontSize = function(size)
    {
        if (size < 6)
            return;

        this.DefaultFontSize = size;
    };

    this.CorrectFontSize = function(size)
    {
        var dRes = size / this.DefaultFontSize;
        dRes = (1 + dRes) / 2;
        dRes = (100 * dRes) >> 0;
        dRes /= 100;

        return "" + dRes + "em";
    };
}

function GetObjectsForImageDownload(aBuilderImages, bSameDoc)
{
    var oMapImages = {}, aBuilderImagesByUrl = [], aUrls =[];
    for(var i = 0; i < aBuilderImages.length; ++i)
    {
		let oBuilderImg = aBuilderImages[i];
		let sUrl = oBuilderImg.Url;
        if(!g_oDocumentUrls.getImageLocal(sUrl) && !g_oDocumentUrls.isThemeUrl(sUrl))
        {
            if(!Array.isArray(oMapImages[sUrl]))
            {
                oMapImages[sUrl] = [];
            }
            oMapImages[sUrl].push(oBuilderImg);
        }
    }
    for(var key in oMapImages)
    {
        if(oMapImages.hasOwnProperty(key))
        {
            aUrls.push(key);
            aBuilderImagesByUrl.push(oMapImages[key]);
        }
    }
    if(bSameDoc !== true){
        //в конце добавляем ссылки на wmf, ole
        for(var i = 0; i < aBuilderImages.length; ++i)
        {
            var oBuilderImage = aBuilderImages[i];
            if (!g_oDocumentUrls.getImageLocal(oBuilderImage.Url))
            {
                if (oBuilderImage.AdditionalUrls) {
                    for (var j = 0; j < oBuilderImage.AdditionalUrls.length; ++j) {
                        aUrls.push(oBuilderImage.AdditionalUrls[j]);
                    }
                }
            }
        }

    }
    return {aUrls: aUrls, aBuilderImagesByUrl: aBuilderImagesByUrl};
}

function ResetNewUrls(data, aUrls, aBuilderImagesByUrl, oImageMap)
{
    for (var i = 0, length = Math.min(data.length, aBuilderImagesByUrl.length); i < length; ++i)
    {
        var elem = data[i];
        if (null != elem.url)
        {
            var name = g_oDocumentUrls.imagePath2Local(elem.path);
            var aImageElem = aBuilderImagesByUrl[i];
            if(Array.isArray(aImageElem))
            {
                for(var j = 0; j < aImageElem.length; ++j)
                {
                    var imageElem = aImageElem[j];
                    if (null != imageElem)
                    {
                        imageElem.SetUrl(name);
                    }
                }
            }
            oImageMap[i] = name;
        }
        else
        {
            oImageMap[i] = aUrls[i];
        }
    }
}

//TODO на счёт коэффициэнта не нахожу подходящего преобразования. пересмотреть.
var koef_mm_to_indent = 3.88;

var PasteElementsId = {
  copyPasteUseBinary : true,
  g_bIsDocumentCopyPaste : true
};
function CopyElement(sName, bText){
	this.sName = sName;
	this.oAttributes = {};
	this.aChildren = [];
	this.bText = bText;
}
CopyElement.prototype.addChild = function(child){
	if(child.bText && this.aChildren.length > 0 && this.aChildren[this.aChildren.length - 1].bText)
		this.aChildren[this.aChildren.length - 1].sName += child.sName;//обьединяем текст, потому что есть места где мы определяем количество child и будет неправильное значение, потому на getOuterHtml тест обьединится в один
	else
		this.aChildren.push(child);
};
CopyElement.prototype.wrapChild = function(child){
	for(var i = 0; i < this.aChildren.length; ++i)
		child.addChild(this.aChildren[i]);
	this.aChildren = [child];
};
CopyElement.prototype.isEmptyChild = function(){
	return 0 === this.aChildren.length;
};
CopyElement.prototype.getInnerText = function(){
	if(this.bText)
		return this.sName;
	else{
		var sRes = "";
		for(var i = 0; i < this.aChildren.length; ++i)
			sRes += this.aChildren[i].getInnerText();
		return sRes;
	}
};
CopyElement.prototype.getInnerHtml = function(){
	if(this.bText)
		return this.sName;
	else{
		var sRes = "";
		for(var i = 0; i < this.aChildren.length; ++i)
			sRes += this.aChildren[i].getOuterHtml();
		return sRes;
	}
};
CopyElement.prototype.getOuterHtml = function(){
	if(this.bText)
		return this.sName;
	else{
		var sRes = "<" + this.sName;
		for(var i in this.oAttributes)
			sRes += " " + i + "=\"" + this.oAttributes[i] + "\"";
		var sInner = this.getInnerHtml();
		if(sInner.length > 0)
			sRes += ">" + sInner + "</" + this.sName + ">";
		else
			sRes += "/>";
		return sRes;
	}
};
CopyElement.prototype.moveChildTo = function (container) {
	for (let i = 0; i < this.aChildren.length; i++) {
		container.addChild && container.addChild(this.aChildren[i]);
	}
};
function CopyProcessor(api, onlyBinaryCopy)
{
	this.api = api;
    this.oDocument = api.WordControl.m_oLogicDocument;
	this.onlyBinaryCopy = onlyBinaryCopy;

	this.oBinaryFileWriter = new AscCommonWord.BinaryFileWriter(this.oDocument);
	this.oPresentationWriter = new AscCommon.CBinaryFileWriter();
    this.oPresentationWriter.Start_UseFullUrl();
    if (this.api.ThemeLoader) {
        this.oPresentationWriter.Start_UseDocumentOrigin(this.api.ThemeLoader.ThemesUrlAbs);
    }

    this.aFootnoteReference = [];
	this.oRoot = new CopyElement("root");
    this.listNextNumMap = [];
    this.instructionHyperlinkStart = null;
}
CopyProcessor.prototype =
{
    getInnerHtml : function()
    {
        return this.oRoot.getInnerHtml();
    },
    getInnerText : function()
    {
        return this.oRoot.getInnerText();
    },
    RGBToCSS : function(rgb, unifill)
    {
        if (null == rgb && null != unifill) {
            unifill.check(this.oDocument.Get_Theme(), this.oDocument.Get_ColorMap());
            var RGBA = unifill.getRGBAColor();
            rgb = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B);
        }
        var sResult = "#";
        var sR = rgb.r.toString(16);
        if(sR.length === 1)
            sR = "0" + sR;
        var sG = rgb.g.toString(16);
        if(sG.length === 1)
            sG = "0" + sG;
        var sB = rgb.b.toString(16);
        if(sB.length === 1)
            sB = "0" + sB;
        return "#" + sR + sG + sB;
    },
    Commit_pPr : function(Item, Para, nextElem)
    {
        //pPr
        var apPr = [];
        var Def_pPr = this.oDocument.Styles ? this.oDocument.Styles.Default.ParaPr : null;
        var Item_pPr = Item.CompiledPr && Item.CompiledPr.Pr && Item.CompiledPr.Pr.ParaPr ? Item.CompiledPr.Pr.ParaPr : Item.Pr;
        if(Item_pPr && Def_pPr)
        {
            //Ind
            if(Def_pPr.Ind.Left !== Item_pPr.Ind.Left)
                apPr.push("margin-left:" + (Item_pPr.Ind.Left * g_dKoef_mm_to_pt) + "pt");
            if(Def_pPr.Ind.Right !== Item_pPr.Ind.Right)
                apPr.push("margin-right:" + ( Item_pPr.Ind.Right * g_dKoef_mm_to_pt) + "pt");
            if(Def_pPr.Ind.FirstLine !== Item_pPr.Ind.FirstLine)
                apPr.push("text-indent:" + (Item_pPr.Ind.FirstLine * g_dKoef_mm_to_pt) + "pt");
            //Jc
            if(Def_pPr.Jc !== Item_pPr.Jc){
                switch(Item_pPr.Jc)
                {
                    case align_Left: apPr.push("text-align:left");break;
                    case align_Center: apPr.push("text-align:center");break;
                    case align_Right: apPr.push("text-align:right");break;
                    case align_Justify: apPr.push("text-align:justify");break;
                }
            }
            //KeepLines , WidowControl
            if(Def_pPr.KeepLines !== Item_pPr.KeepLines || Def_pPr.WidowControl !== Item_pPr.WidowControl)
            {
                if(Def_pPr.KeepLines !== Item_pPr.KeepLines && Def_pPr.WidowControl !== Item_pPr.WidowControl)
                    apPr.push("mso-pagination:none lines-together");
                else if(Def_pPr.KeepLines !== Item_pPr.KeepLines)
                    apPr.push("mso-pagination:widow-orphan lines-together");
                else if(Def_pPr.WidowControl !== Item_pPr.WidowControl)
                    apPr.push("mso-pagination:none");
            }
            //KeepNext
            if(Def_pPr.KeepNext !== Item_pPr.KeepNext)
                apPr.push("page-break-after:avoid");
            //PageBreakBefore
            if(Def_pPr.PageBreakBefore !== Item_pPr.PageBreakBefore)
                apPr.push("page-break-before:always");
            //Spacing
            if(Def_pPr.Spacing.Line !== Item_pPr.Spacing.Line)
            {
                if(Asc.linerule_AtLeast === Item_pPr.Spacing.LineRule)
                    apPr.push("line-height:"+(Item_pPr.Spacing.Line * g_dKoef_mm_to_pt)+"pt");
                else if( Asc.linerule_Auto === Item_pPr.Spacing.LineRule)
                {
                    if(1 === Item_pPr.Spacing.Line)
                        apPr.push("line-height:normal");
                    else
                        apPr.push("line-height:"+parseInt(Item_pPr.Spacing.Line * 100)+"%");
                }
            }
            if(Def_pPr.Spacing.LineRule !== Item_pPr.Spacing.LineRule)
            {
                if(Asc.linerule_Exact === Item_pPr.Spacing.LineRule)
                    apPr.push("mso-line-height-rule:exactly");
            }
			//TODO при вставке в EXCEL(внутрь ячейки) появляются лишние пустые строки из-за того, что в HTML пишутся отступы - BUG #14663
			//При вставке в word лучше чтобы эти значения выставлялись всегда
            //if(Def_pPr.Spacing.Before != Item_pPr.Spacing.Before)
            apPr.push("margin-top:" + (Item_pPr.Spacing.Before * g_dKoef_mm_to_pt) + "pt");
            //if(Def_pPr.Spacing.After != Item_pPr.Spacing.After)
            apPr.push("margin-bottom:" + (Item_pPr.Spacing.After * g_dKoef_mm_to_pt) + "pt");
            //Shd
            if (null != Item_pPr.Shd && c_oAscShdNil !== Item_pPr.Shd.Value && (null != Item_pPr.Shd.Color || null != Item_pPr.Shd.Unifill)){
				var _shdColor = Item_pPr.Shd.GetSimpleColor && Item_pPr.Shd.GetSimpleColor(this.oDocument.Get_Theme(), this.oDocument.Get_ColorMap());
				//TODO проверить сохранение в epub
				//todo проверить и убрать else, всегда использовать GetSimpleColor
				if (_shdColor) {
					_shdColor = this.RGBToCSS(_shdColor);
				} else {
					_shdColor = this.RGBToCSS(Item_pPr.Shd.Color, Item_pPr.Shd.Unifill);
				}
            	apPr.push("background-color:" + _shdColor);
			}
            //Tabs
            if(Item_pPr.Tabs.Get_Count() > 0)
            {
                var sTabs = "";
                //tab-stops:1.0cm 3.0cm 5.0cm
                for(var i = 0, length = Item_pPr.Tabs.Get_Count(); i < length; i++)
                {
                    if(0 !== i)
                        sTabs += " ";
                    sTabs += Item_pPr.Tabs.Get(i).Pos / 10 + "cm";
                }
                apPr.push("tab-stops:" + sTabs);
            }
            //Border
            if(null != Item_pPr.Brd)
            {
                apPr.push("border:none");

                //сравниваю бордеры со следующим параграфом
                var isNeedPrefix = true;
                if (Item && type_Paragraph === Item.GetType() && Item.IsTableCellContent && !Item.IsTableCellContent()) {
					isNeedPrefix = false;
                	if (nextElem && type_Paragraph === nextElem.GetType()) {
						isNeedPrefix = true;
						var Item_pPrCompile = Item.CompiledPr && Item.CompiledPr.Pr && Item.CompiledPr.Pr.ParaPr;
						var Next_pPrCompile = nextElem.Get_CompiledPr2 && nextElem.Get_CompiledPr2(false);
						Next_pPrCompile = Next_pPrCompile && Next_pPrCompile.ParaPr;
						if (Next_pPrCompile && Item_pPrCompile && !nextElem.private_CompareBorderSettings(Next_pPrCompile, Item_pPrCompile)) {
							isNeedPrefix = false;
						}
					}
				}

                var borderStyle = this._BordersToStyle(Item_pPr.Brd, false, true, isNeedPrefix ? "mso-" : null, isNeedPrefix ? "-alt" : null);
                if(null != borderStyle)
                {
                    var nborderStyleLength = borderStyle.length;
                    if(nborderStyleLength> 0)
                        borderStyle = borderStyle.substring(0, nborderStyleLength - 1);
                    apPr.push(borderStyle);
                }
            }
        }
        if(apPr.length > 0)
            Para.oAttributes["style"] = apPr.join(';');
    },
    parse_para_TextPr : function(Value, oTarget)
    {
        var aProp = [];
        if (null != Value.RFonts) {
            var sFontName = null;
            if (null != Value.RFonts.Ascii)
                sFontName = Value.RFonts.Ascii.Name;
            else if (null != Value.RFonts.HAnsi)
                sFontName = Value.RFonts.HAnsi.Name;
            else if (null != Value.RFonts.EastAsia)
                sFontName = Value.RFonts.EastAsia.Name;
            else if (null != Value.RFonts.CS)
                sFontName = Value.RFonts.CS.Name;
            if (null != sFontName){
                var oTheme = this.oDocument && this.oDocument.Get_Theme && this.oDocument.Get_Theme();
                if(oTheme && oTheme.themeElements && oTheme.themeElements.fontScheme){
                    sFontName = oTheme.themeElements.fontScheme.checkFont(sFontName)
                }
                aProp.push("font-family:" + "'" + CopyPasteCorrectString(sFontName) + "'");
            }

        }
        if (null != Value.FontSize) {
            if (!this.api.DocumentReaderMode)
                aProp.push("font-size:" + Value.FontSize + "pt");//font-size в pt все остальные метрики в mm
            else
                aProp.push("font-size:" + this.api.DocumentReaderMode.CorrectFontSize(Value.FontSize));
        }
        if (true == Value.Bold)
			oTarget.wrapChild(new CopyElement("b"));
        if (true == Value.Italic)
			oTarget.wrapChild(new CopyElement("i"));
		if (true == Value.Underline)
			oTarget.wrapChild(new CopyElement("u"));
        if (true == Value.Strikeout)
			oTarget.wrapChild(new CopyElement("s"));
		 if (true == Value.DStrikeout)
			 oTarget.wrapChild(new CopyElement("s"));
        if (null != Value.Shd && c_oAscShdNil !== Value.Shd.Value && (null != Value.Shd.Color || null != Value.Shd.Unifill))
            aProp.push("background-color:" + this.RGBToCSS(Value.Shd.Color, Value.Shd.Unifill));
        else if (null != Value.HighLight && highlight_None !== Value.HighLight)
            aProp.push("background-color:" + this.RGBToCSS(Value.HighLight, null));
        if (null != Value.Color || null != Value.Unifill) {
			var color;
			//TODO правка того, что в полученной html цвет текста всегда чёрный. стоит пересмотреть.
			if(null != Value.Unifill)
				color = this.RGBToCSS(null, Value.Unifill);
			else
				color = this.RGBToCSS(Value.Color, Value.Unifill);

            aProp.push("color:" + color);
            aProp.push("mso-style-textfill-fill-color:" + color);
        }
        if (null != Value.VertAlign) {
            if(AscCommon.vertalign_SuperScript === Value.VertAlign)
                aProp.push("vertical-align:super");
            else if(AscCommon.vertalign_SubScript === Value.VertAlign)
                aProp.push("vertical-align:sub");
        }
		if(aProp.length > 0)
			oTarget.oAttributes["style"] = aProp.join(';');
    },
    ParseItem : function(ParaItem, oTarget, nextParaItem, lengthContent)
    {
        let oSpan;
    	switch ( ParaItem.Type )
        {
            case para_Text:
				//экранируем спецсимволы
                let sValue = AscCommon.encodeSurrogateChar(ParaItem.Value);
                if(sValue)
					oTarget.addChild(new CopyElement(CopyPasteCorrectString(sValue), true));
                break;
            case para_Space:
				//TODO пересмотреть обработку пробелов - возможно стоит всегда копировать неразрывный пробел!!!!!
				//в случае нескольких пробелов друг за другом добавляю неразрывный пробел, иначе добавиться только один
				//lengthContent - если в элемент добавляется только один пробел, этот элемент не записывается в буфер, поэтому добавляем неразрывный пробел
				if((nextParaItem && nextParaItem.Type === para_Space) || lengthContent === 1)
					oTarget.addChild(new CopyElement("&nbsp;", true));
				else
					oTarget.addChild(new CopyElement(" ", true));
				break;
			case para_Tab:
				oSpan = new CopyElement("span");
				oSpan.oAttributes["style"] = "white-space:pre;";
				oSpan.oAttributes["style"] = "mso-tab-count:1;";
				oSpan.addChild(new CopyElement(String.fromCharCode(0x09), true));
				oTarget.addChild(oSpan);
				break;
			case para_NewLine:
				//page break -> <br clear=all style='mso-special-character:line-break;page-break-before:always'>
				//column break -> <br clear=all style='mso-column-break-before:always'>
				//section break(next page) -> <br clear=all style='page-break-before:always;mso-break-type:section-break'>
				//section break(even page) -> <br clear=all style='page-break-before:left;mso-break-type:section-break'>
				//section break(odd page)  -> <br clear=all style='page-break-before:right;mso-break-type:section-break'>

				let oBr = new CopyElement("br");

				if (ParaItem.BreakType === window['AscWord'].break_Page) {
					oBr.oAttributes["clear"] = "all";
					oBr.oAttributes["style"] = "mso-special-character:line-break;page-break-before:always;";
				} else if (ParaItem.BreakType === window['AscWord'].break_Column) {
					oBr.oAttributes["clear"] = "all";
					oBr.oAttributes["style"] = "mso-column-break-before:always;";
				} /*else if (ParaItem.BreakType === window['AscWord'].break_Line) {
					if (ParaItem.Clear === AscWord.break_Clear_All) {
						oBr.oAttributes["clear"] = "all";
						oBr.oAttributes["style"] = "page-break-before:always;mso-break-type:section-break;";
					} else if (ParaItem.Clear === AscWord.break_Clear_Left) {
						oBr.oAttributes["clear"] = "all";
						oBr.oAttributes["style"] = "page-break-before:left;mso-break-type:section-break;";
					} else if (ParaItem.Clear === AscWord.break_Clear_Right) {
						oBr.oAttributes["clear"] = "all";
						oBr.oAttributes["style"] = "page-break-before:right;mso-break-type:section-break;";
					}
				}*/ else {
					oBr.oAttributes["style"] = "mso-special-character:line-break;";
				}


				oTarget.addChild(oBr);
				//todo закончить этот параграф и начать новый
				//добавил неразрвной пробел для того, чтобы информация попадала в буфер обмена
				oSpan = new CopyElement("span");
				oSpan.addChild(new CopyElement("&nbsp;", true));
				oTarget.addChild(oSpan);
				break;
            case para_Drawing:
                let oGraphicObj = ParaItem.GraphicObj;
                let sSrc = oGraphicObj.getBase64Img();
                if(sSrc.length > 0)
                {
					let _h, _w;
					if(oGraphicObj.cachedPixH)
						_h = oGraphicObj.cachedPixH;
					else
						_h = ParaItem.Extent.H * g_dKoef_mm_to_pix;

					if(oGraphicObj.cachedPixW)
						_w = oGraphicObj.cachedPixW;
					else
						_w = ParaItem.Extent.W * g_dKoef_mm_to_pix;

					let oImg = new CopyElement("img");
					oImg.oAttributes["style"] = "max-width:100%;";
					oImg.oAttributes["width"] = Math.round(_w);
					oImg.oAttributes["height"] = Math.round(_h);
					oImg.oAttributes["src"] = sSrc;
					oTarget.addChild(oImg);
                    break;
                }
                break;
			case para_PageNum:
				if(null != ParaItem.String && "string" === typeof(ParaItem.String))
					oTarget.addChild(new CopyElement(CopyPasteCorrectString(ParaItem.String), true));
				break;
			case para_FootnoteReference:
				let oLink = new CopyElement("a");
				let index = this.aFootnoteReference.length + 1;
				let prefix = "ftn";
				oLink.oAttributes["style"] = "mso-footnote-id:" + prefix + index;
				oLink.oAttributes["href"] = "#_" + prefix + index;
				oLink.oAttributes["name"] = "_" + prefix + "ref" + index;
				oLink.oAttributes["title"] = "";

				oSpan = new CopyElement("span");
				oSpan.oAttributes["class"] = "MsoFootnoteReference";


				let _oSpan2 = new CopyElement("span");
				_oSpan2.addChild(new CopyElement(CopyPasteCorrectString("[" + index + "]"), true));
				if (_oSpan2.oAttributes["style"]) {
					_oSpan2.oAttributes["style"] += ";"
				} else {
					_oSpan2.oAttributes["style"] = "";
				}
				_oSpan2.oAttributes["style"] += "mso-special-character:footnote";

				oSpan.addChild(_oSpan2);
				this.parse_para_TextPr(ParaItem.Run.Get_CompiledTextPr(), oSpan);
				oLink.addChild(oSpan);
				oTarget.addChild(oLink);
				this.aFootnoteReference.push(ParaItem.Footnote);
				break;
			case para_FieldChar:
				if (ParaItem.ComplexField && ParaItem.ComplexField.Instruction && ParaItem.ComplexField.Instruction instanceof CFieldInstructionHYPERLINK) {
					if (fldchartype_Begin === ParaItem.CharType && ParaItem.ComplexField.Instruction.BookmarkName) {
						this.instructionHyperlinkStart = "#" + ParaItem.ComplexField.Instruction.BookmarkName;
					} else if (fldchartype_End === ParaItem.CharType) {
						this.instructionHyperlinkStart = null;
					}
				}
				break;
        }
    },
    CopyRun: function (Item, oTarget) {
		for (var i = 0; i < Item.Content.length; i++) {
			this.ParseItem(Item.Content[i], oTarget, Item.Content[i + 1], Item.Content.length);
		}
    },
    CopyRunContent: function (Container, oTarget, bOmitHyperlink) {
		var bookmarksStartMap = {};
		var bookmarkPrviousTargetMap = {};
		var bookmarkLevel = 0;

		var closeBookmarks = function (_level) {
			var tempTarget = bookmarkPrviousTargetMap[_level];
			if (tempTarget) {
				tempTarget.addChild(oTarget);
				oTarget = tempTarget;
			}
		};


		var realTarget;
		var oHyperlink;
    	for (var i = 0; i < Container.Content.length; i++) {
			var item = Container.Content[i];
			if (para_Run === item.Type) {
				//отдельная обработка для сносок, добавляем внутри данные
				if (item.Content && item.Content.length === 1 && item.Content[0] && item.Content[0].Type === para_FootnoteReference) {
					this.CopyRun(item, oTarget);
				} else {
					var oSpan = new CopyElement("span");
					this.CopyRun(item, oSpan);

					if (this.instructionHyperlinkStart && !realTarget) {
						oHyperlink = new CopyElement("a");
						oHyperlink.oAttributes["href"] = CopyPasteCorrectString(this.instructionHyperlinkStart);
						oTarget.addChild(oHyperlink);
						realTarget = oTarget;
						oTarget = oHyperlink;
						bOmitHyperlink = true;
					} else if (realTarget && !this.instructionHyperlinkStart) {
						//TODO close all bookmarks before change content. need to reconsider
						if (bookmarkLevel > 0) {
							while(bookmarkLevel > 0) {
								closeBookmarks(bookmarkLevel);
								bookmarkLevel--;
							}
						}
						oTarget = realTarget;
						bOmitHyperlink = false;
						realTarget = null;
					}


					if (!oSpan.isEmptyChild()) {
						this.parse_para_TextPr(item.Get_CompiledTextPr(), oSpan);
						oTarget.addChild(oSpan);
					}
				}
			} else if (para_Hyperlink === item.Type) {
				if (!bOmitHyperlink) {
					oHyperlink = new CopyElement("a");
					var sValue = item.IsAnchor() ? "#" + item.Anchor : item.GetValue();
					var sToolTip = item.GetToolTip();
					oHyperlink.oAttributes["href"] = CopyPasteCorrectString(sValue);
					oHyperlink.oAttributes["title"] = CopyPasteCorrectString(sToolTip);
					//вложенные ссылки в html запрещены.
					this.CopyRunContent(item, oHyperlink, true);
					oTarget.addChild(oHyperlink);
				} else {
					this.CopyRunContent(item, oTarget, true);
				}
			} else if (para_Math === item.Type) {
				var sSrc = item.MathToImageConverter();
				if (null != sSrc && null != sSrc.ImageUrl) {
					var oImg = new CopyElement("img");
					if (sSrc.w_px > 0) {
						oImg.oAttributes["width"] = sSrc.w_px;
					}
					if (sSrc.h_px > 0) {
						oImg.oAttributes["height"] = sSrc.h_px;
					}
					oImg.oAttributes["src"] = sSrc.ImageUrl;
					oTarget.addChild(oImg);
				}
			} else if (para_InlineLevelSdt === item.Type) {
				this.CopyRunContent(item, oTarget);
			} else if (para_Field === item.Type) {

				this.CopyRunContent(item, oTarget);
			} else if (para_Bookmark === item.Type) {
				//для внутренних ссылок
				//если конец ссылки находится в тепкущем параграфе, то закрываем тэг ссылки здесь
				//если он находится в следующем параграфе, то закрываем после того, как прошлись по всему содержимому данного параграфа
				//ms в данном случае берёт только первый элемент
				//чтобы заранее не проходиться по всему контенту параграфа в поисках закрытия bookmark - закрываю его после всего цикла
				//на следующий параграф не переносим
				if (item.Start) {
					bookmarkLevel++;
					bookmarksStartMap[item.BookmarkId] = 1;
					var oBookmark = new CopyElement("a");
					var name = item.GetBookmarkName();
					oBookmark.oAttributes["name"] = CopyPasteCorrectString(name);
					bookmarkPrviousTargetMap[bookmarkLevel] = oTarget;
					oTarget = oBookmark;
				} else if (bookmarksStartMap[item.BookmarkId]) {
					bookmarksStartMap[item.BookmarkId] = 0;
					closeBookmarks(bookmarkLevel);
					bookmarkLevel--;
				}
			}
		}

		if (bookmarkLevel > 0) {
    		while(bookmarkLevel > 0) {
				closeBookmarks(bookmarkLevel);
				bookmarkLevel--;
			}
		}
		if (this.instructionHyperlinkStart) {
			this.instructionHyperlinkStart = null;
		}
    },
    CopyParagraph : function(oDomTarget, Item, selectedAll, nextElem)
    {
        var oDocument = this.oDocument;
		var Para = null;
		//Для heading пишем в h1
        var styleId = Item.Style_Get();
        if(styleId)
        {
            var styleName = oDocument.Styles.Get_Name( styleId ).toLowerCase();
			//шаблон "heading n" (n=1:6)
            if(0 === styleName.indexOf("heading"))
            {
                var nLevel = parseInt(styleName.substring("heading".length));
                if(1 <= nLevel && nLevel <= 6)
                    Para = new CopyElement("h" + nLevel);
            }
        }
        if(null == Para)
            Para = new CopyElement("p");

        var oNumPr;
        var bIsNullNumPr = false;
        if(PasteElementsId.g_bIsDocumentCopyPaste)
        {
            oNumPr = Item.GetNumPr();
            bIsNullNumPr = (null == oNumPr || 0 == oNumPr.NumId);
        }
        else
        {
            oNumPr = Item.PresentationPr.Bullet;
            bIsNullNumPr = (0 == oNumPr.m_nType);
        }
		var bBullet = false;
        var sListStyle = "";
        if(!bIsNullNumPr)
        {
            if(PasteElementsId.g_bIsDocumentCopyPaste)
			{
				var oNum = this.oDocument.GetNumbering().GetNum(oNumPr.NumId);
				if (oNum)
				{
					var oNumberingLvl = oNum.GetLvl(oNumPr.Lvl);
					if (oNumberingLvl)
					{
						switch (oNumberingLvl.GetFormat())
						{
							case Asc.c_oAscNumberingFormat.Decimal:
								sListStyle = "decimal";
								break;
							case Asc.c_oAscNumberingFormat.LowerRoman:
								sListStyle = "lower-roman";
								break;
							case Asc.c_oAscNumberingFormat.UpperRoman:
								sListStyle = "upper-roman";
								break;
							case Asc.c_oAscNumberingFormat.LowerLetter:
								sListStyle = "lower-alpha";
								break;
							case Asc.c_oAscNumberingFormat.UpperLetter:
								sListStyle = "upper-alpha";
								break;
							default:
								sListStyle = "disc";
								bBullet    = true;
								break;
						}
					}
				}
			}
            else
            {
                var _presentation_bullet = Item.PresentationPr.Bullet;
                switch(_presentation_bullet.m_nType)
                {
                    case AscFormat.numbering_presentationnumfrmt_ArabicParenBoth:
                    case AscFormat.numbering_presentationnumfrmt_ArabicParenR:
                    case AscFormat.numbering_presentationnumfrmt_ArabicPeriod:
                    case AscFormat.numbering_presentationnumfrmt_ArabicPlain:
                    {
                        sListStyle = "decimal";
                        break;
                    }
                    case AscFormat.numbering_presentationnumfrmt_RomanLcParenBoth:
                    case AscFormat.numbering_presentationnumfrmt_RomanLcParenR:
                    case AscFormat.numbering_presentationnumfrmt_RomanLcPeriod:
                    {
                        sListStyle = "lower-roman";
                        break;
                    }
                    case AscFormat.numbering_presentationnumfrmt_RomanUcParenBoth:
                    case AscFormat.numbering_presentationnumfrmt_RomanUcParenR:
                    case AscFormat.numbering_presentationnumfrmt_RomanUcPeriod:
                    {
                        sListStyle = "upper-roman";
                        break;
                    }

                    case AscFormat.numbering_presentationnumfrmt_AlphaLcParenBoth:
                    case AscFormat.numbering_presentationnumfrmt_AlphaLcParenR:
                    case AscFormat.numbering_presentationnumfrmt_AlphaLcPeriod:
                    {
                        sListStyle = "lower-alpha";
                        break;
                    }
                    case AscFormat.numbering_presentationnumfrmt_AlphaUcParenR:
                    case AscFormat.numbering_presentationnumfrmt_AlphaUcPeriod:
                    case AscFormat.numbering_presentationnumfrmt_AlphaUcParenBoth:
                    {
                        sListStyle = "upper-alpha";
                        break;
                    }
                    default:
                        sListStyle = "disc";
                        bBullet = true;
                        break;
                }
            }
        }
        //pPr
        this.Commit_pPr(Item, Para, nextElem);

        if(false === selectedAll)
        {
			//если последний элемент в выделении неполностью выделенный параграф, то он копируется как обычный текст без настроек параграфа и списков
			this.CopyRunContent(Item, oDomTarget, false);
        }
        else
        {
			this.CopyRunContent(Item, Para, false);
			//добавляем &nbsp; потому что параграфы без содержимого не копируются
            if(Para.isEmptyChild())
                Para.addChild(new CopyElement("&nbsp;", true));
            if(bIsNullNumPr)
                oDomTarget.addChild( Para );
			else{
				var Li = new CopyElement( "li" );
				Li.oAttributes["style"] = "list-style-type: " + sListStyle;
				Li.addChild( Para );
				//пробуем добавить в предыдущий список
				var oTargetList = null;
				if(oDomTarget.aChildren.length > 0){
					var oPrevElem = oDomTarget.aChildren[oDomTarget.aChildren.length - 1];
					if((bBullet && "ul" === oPrevElem.sName) || (!bBullet && "ol" === oPrevElem.sName))
						oTargetList = oPrevElem;
				}

				if (!bBullet) {
					if (!this.listNextNumMap[oNumPr.NumId]) {
						this.listNextNumMap[oNumPr.NumId] = 1;
					} else {
						this.listNextNumMap[oNumPr.NumId]++;
					}
				}
				if (null == oTargetList) {
					if (bBullet) {
						oTargetList = new CopyElement("ul");
					} else {
						oTargetList = new CopyElement("ol");
					}
					oTargetList.oAttributes["style"] = "padding-left:40px";
					//если список идёт с промежуточными элементами, добавляем аттрибут start
					if (!bBullet && this.listNextNumMap[oNumPr.NumId] > 1) {
						oTargetList.oAttributes["start"] = this.listNextNumMap[oNumPr.NumId];
					}
					oDomTarget.addChild(oTargetList);
				}
				oTargetList.addChild(Li);
			}
        }
    },
    _BorderToStyle : function(border, name)
    {
        var res = "";
        if(border_None === border.Value)
            res += name + ":none;";
        else
        {
            //TODO получение цвета рассмотреть аналогично получению фону у ячейки с ипользованием функции GetSimpleColor
        	var size = 0.5;
            var color = border.Color;
            var unifill = border.Unifill;
            if(null != border.Size)
                size = border.Size * g_dKoef_mm_to_pt;
            if (null == color)
                color = { r: 0, g: 0, b: 0 };
            res += name + ":" + size + "pt solid " + this.RGBToCSS(color, unifill) + ";";
        }
        return res;
    },
    _MarginToStyle : function(margins, styleName)
    {
        var res = "";
        var nMarginLeft = 1.9;
        var nMarginTop = 0;
        var nMarginRight = 1.9;
        var nMarginBottom = 0;
        if(null != margins.Left && tblwidth_Mm === margins.Left.Type && null != margins.Left.W)
            nMarginLeft = margins.Left.W;
        if(null != margins.Top && tblwidth_Mm === margins.Top.Type && null != margins.Top.W)
            nMarginTop = margins.Top.W;
        if(null != margins.Right && tblwidth_Mm === margins.Right.Type && null != margins.Right.W)
            nMarginRight = margins.Right.W;
        if(null != margins.Bottom && tblwidth_Mm === margins.Bottom.Type && null != margins.Bottom.W)
            nMarginBottom = margins.Bottom.W;
        res = styleName + ":"+(nMarginTop * g_dKoef_mm_to_pt)+"pt "+(nMarginRight * g_dKoef_mm_to_pt)+"pt "+(nMarginBottom * g_dKoef_mm_to_pt)+"pt "+(nMarginLeft * g_dKoef_mm_to_pt)+"pt;";
        return res;
    },
    _BordersToStyle : function(borders, bUseInner, bUseBetween, mso, alt)
    {
        var res = "";
        if(null == mso)
            mso = "";
        if(null == alt)
            alt = "";
        if(null != borders.Left)
            res += this._BorderToStyle(borders.Left, mso + "border-left" + alt);
        if(null != borders.Top)
            res += this._BorderToStyle(borders.Top, mso + "border-top" + alt);
        if(null != borders.Right)
            res += this._BorderToStyle(borders.Right, mso + "border-right" + alt);
        if(null != borders.Bottom)
            res += this._BorderToStyle(borders.Bottom, mso + "border-bottom" + alt);
		if(bUseInner)
		{
			if(null != borders.InsideV)
				res += this._BorderToStyle(borders.InsideV, "mso-border-insidev");
			if(null != borders.InsideH)
				res += this._BorderToStyle(borders.InsideH, "mso-border-insideh");
		}
		if(bUseBetween)
		{
			if(null != borders.Between)
				res += this._BorderToStyle(borders.Between, "mso-border-between");
		}
        return res;
    },
    _MergeProp : function(elem1, elem2)
    {
        if( !elem1 || !elem2 )
        {
            return;
        }

        var p, v;
        for(p in elem2)
        {
            if(elem2.hasOwnProperty(p) && !elem1.hasOwnProperty(p))
            {
                v = elem2[p];
                if(null != v)
                    elem1[p] = v;
            }
        }
    },
    CopyCell : function(tr, cell, tablePr, width, rowspan)
    {
        var tc = new CopyElement("td");
        //Pr
        var tcStyle = "";
        if(width > 0)
        {
            tc.oAttributes["width"] = Math.round(width * g_dKoef_mm_to_pix);
            tcStyle += "width:"+(width * g_dKoef_mm_to_pt)+"pt;";
        }
        if(rowspan > 1)
            tc.oAttributes["rowspan"] = rowspan;
        var cellPr = null;

		var tablePr = null;
        if(!PasteElementsId.g_bIsDocumentCopyPaste && editor.WordControl.m_oLogicDocument && null != cell.CompiledPr && null != cell.CompiledPr.Pr)
		{
			var presentation = editor.WordControl.m_oLogicDocument;
			var curSlide = presentation.Slides[presentation.CurPage];
			if(presentation && curSlide && curSlide.Layout && curSlide.Layout.Master && curSlide.Layout.Master.Theme)
        AscFormat.checkTableCellPr(cell.CompiledPr.Pr, curSlide, curSlide.Layout, curSlide.Layout.Master, curSlide.Layout.Master.Theme);
		}

		if(null != cell.CompiledPr && null != cell.CompiledPr.Pr)
        {
            cellPr = cell.CompiledPr.Pr;
			//Для первых и послених ячеек выставляются margin а не colspan
            if(null != cellPr.GridSpan && cellPr.GridSpan > 1)
				tc.oAttributes["colspan"] = cellPr.GridSpan;
        }
        if(null != cellPr && null != cellPr.Shd)
        {
			if (c_oAscShdNil !== cellPr.Shd.Value && (null != cellPr.Shd.Color || null != cellPr.Shd.Unifill)) {
				var _shdColor = cellPr.Shd.GetSimpleColor(this.oDocument.Get_Theme(), this.oDocument.Get_ColorMap());
				//todo проверить и убрать else, всегда использовать GetSimpleColor
				if (_shdColor) {
					_shdColor = this.RGBToCSS(_shdColor);
				} else {
					_shdColor = this.RGBToCSS(cellPr.Shd.Color, cellPr.Shd.Unifill);
				}
				tcStyle += "background-color:" + _shdColor + ";";
			}
        }
        else if(null != tablePr && null != tablePr.Shd)
        {
            if (c_oAscShdNil !== tablePr.Shd.Value && (null != tablePr.Shd.Color || null != tablePr.Shd.Unifill))
                tcStyle += "background-color:" + this.RGBToCSS(tablePr.Shd.Color, tablePr.Shd.Unifill) + ";";
        }
        var oCellMar = {};
        if(null != cellPr && null != cellPr.TableCellMar)
            this._MergeProp(oCellMar, cellPr.TableCellMar);
        if(null != tablePr && null != tablePr.TableCellMar)
            this._MergeProp(oCellMar, tablePr.TableCellMar);
        tcStyle += this._MarginToStyle(oCellMar, "padding");

        var oCellBorder = cell.Get_Borders();
        tcStyle += this._BordersToStyle(oCellBorder, false, false);

        if("" != tcStyle)
            tc.oAttributes["style"] = tcStyle;
        //Content
        this.CopyDocument2(tc, cell.Content);

        tr.addChild(tc);
    },
    CopyRow : function(oDomTarget, table, nCurRow, elems, nMaxRow)
    {
        var row = table.Content[nCurRow];
		if(null == elems)
			elems = {gridStart: 0, gridEnd: table.TableGrid.length - 1, indexStart: null, indexEnd: null, after: null, before: null, cells: row.Content};
        var tr = new CopyElement("tr");
        //Pr
		var gridSum = table.TableSumGrid;
        var trStyle = "";
        var nGridBefore = 0;
		var rowPr = null;

		var CompiledPr = row.Get_CompiledPr();
		if(null != CompiledPr)
			rowPr = CompiledPr;
        if(null != rowPr)
        {
			if(null == elems.before && null != rowPr.GridBefore && rowPr.GridBefore > 0)
			{
                elems.before = rowPr.GridBefore;
				elems.gridStart += rowPr.GridBefore;
			}
			if(null == elems.after && null != rowPr.GridAfter && rowPr.GridAfter > 0)
			{
                elems.after = rowPr.GridAfter;
				elems.gridEnd -= rowPr.GridAfter;
			}
            //height
            if(null != rowPr.Height && Asc.linerule_Auto != rowPr.Height.HRule && null != rowPr.Height.Value)
            {
                trStyle += "height:"+(rowPr.Height.Value * g_dKoef_mm_to_pt)+"pt;";
            }
        }
		//WBefore
        if(null != elems.before)
        {
            if(elems.before > 0)
            {
                nGridBefore = elems.before;
                var nWBefore = gridSum[elems.gridStart - 1] - gridSum[elems.gridStart - nGridBefore - 1];
				//Записываем margin
                trStyle += "mso-row-margin-left:"+(nWBefore * g_dKoef_mm_to_pt)+"pt;";
				//добавляем td для тех кто не понимает mso-row-margin-left
                var oNewTd = new CopyElement("td");
                oNewTd.oAttributes["style"] = "mso-cell-special:placeholder;border:none;padding:0cm 0cm 0cm 0cm";
                oNewTd.oAttributes["width"] = Math.round(nWBefore * g_dKoef_mm_to_pix);
                if(nGridBefore > 1)
                    oNewTd.oAttributes["colspan"] = nGridBefore;
                var oNewP = new CopyElement("p");
				oNewP.oAttributes["style"] = "margin:0cm";
                oNewP.addChild(new CopyElement("&nbsp;", true));
                oNewTd.addChild(oNewP);
                tr.addChild(oNewTd);
            }
        }

		var tablePr = null;
		var compiledTablePr = table.Get_CompiledPr();
		if(null != compiledTablePr && null != compiledTablePr.TablePr)
			tablePr = compiledTablePr.TablePr;
        //tc
        for(var i in elems.cells)
        {
            var cell = row.Content[i];
			if(vmerge_Continue !== cell.GetVMerge())
			{
				var StartGridCol = cell.Metrics.StartGridCol;
				var GridSpan = cell.Get_GridSpan();
				var width = gridSum[StartGridCol + GridSpan - 1] - gridSum[StartGridCol - 1];
				//вычисляем rowspan
				var nRowSpan = table.Internal_GetVertMergeCount(nCurRow, StartGridCol, GridSpan);
				if(nCurRow + nRowSpan - 1 > nMaxRow)
				{
					nRowSpan = nMaxRow - nCurRow + 1;
					if(nRowSpan <= 0)
						nRowSpan = 1;
				}
				this.CopyCell(tr, cell, tablePr, width, nRowSpan);
			}
        }
        //WAfter
        if(null != elems.after)
        {
            if(elems.after > 0)
            {
                var nGridAfter = elems.after;
                var nWAfter = gridSum[elems.gridEnd + nGridAfter] - gridSum[elems.gridEnd];
				//Записываем margin
                trStyle += "mso-row-margin-right:"+(nWAfter * g_dKoef_mm_to_pt)+"pt;";
				//добавляем td для тех кто не понимает mso-row-margin-left
                var oNewTd = new CopyElement("td");
                oNewTd.oAttributes["style"] = "mso-cell-special:placeholder;border:none;padding:0cm 0cm 0cm 0cm";
				oNewTd.oAttributes["width"] = Math.round(nWAfter * g_dKoef_mm_to_pix);
                if(nGridAfter > 1)
					oNewTd.oAttributes["colspan"] = nGridAfter;
                var oNewP = new CopyElement("p");
				oNewP.oAttributes["style"] = "margin:0cm";
				oNewP.addChild(new CopyElement("&nbsp;", true));
                oNewTd.addChild(oNewP);
                tr.addChild(oNewTd);
            }
        }
        if("" != trStyle)
            tr.oAttributes["style"] = trStyle;

        oDomTarget.addChild(tr);
    },
    CopyTable : function(oDomTarget, table, aRowElems)
    {
        var DomTable = new CopyElement("table");

		var compiledPr = table.Get_CompiledPr();

        var Pr = null;
        if(compiledPr && null != compiledPr.TablePr)
            Pr = compiledPr.TablePr;
        var tblStyle = "";
        var bBorder = false;
        if(null != Pr)
        {
			var align = "";
            if(true != table.Inline && null != table.PositionH)
			{
				var PositionH = table.PositionH;
				if(true === PositionH.Align)
				{
					switch(PositionH.Value)
					{
						case c_oAscXAlign.Outside:
						case c_oAscXAlign.Right: align = "right";break;
						case c_oAscXAlign.Center: align = "center";break;
					}
				}
				else if(table.TableSumGrid)
				{
					var TableWidth = table.TableSumGrid[ table.TableSumGrid.length - 1 ];
					var nLeft = PositionH.Value;
					var nRight = nLeft + TableWidth;
					var nFromLeft = Math.abs(nLeft - X_Left_Margin);
					var nFromCenter = Math.abs((Page_Width - X_Right_Margin + X_Left_Margin) / 2 - (nLeft + nRight) / 2);
					var nFromRight = Math.abs(Page_Width - nRight - X_Right_Margin);
					if(nFromRight < nFromLeft || nFromCenter < nFromLeft)
					{
						if(nFromRight < nFromCenter)
							align = "right";
						else
							align = "center";
					}
				}
			}
			else if(null != Pr.Jc)
            {
                switch(Pr.Jc)
                {
                    case align_Center:align = "center";break;
                    case align_Right:align = "right";break;
                }
            }
			if("" != align)
				DomTable.oAttributes["align"] = align;
            if(null != Pr.TableInd)
                tblStyle += "margin-left:"+(Pr.TableInd * g_dKoef_mm_to_pt)+"pt;";
            if (null != Pr.Shd && c_oAscShdNil !== Pr.Shd.Value && (null != Pr.Shd.Color || null != Pr.Shd.Unifill))
                tblStyle += "background:" + this.RGBToCSS(Pr.Shd.Color, Pr.Shd.Unifill) + ";";
            if(null != Pr.TableCellMar)
                tblStyle += this._MarginToStyle(Pr.TableCellMar, "mso-padding-alt");
            if(null != Pr.TableBorders)
                tblStyle += this._BordersToStyle(Pr.TableBorders, true, false);
        }
		//ищем cellSpacing
        var bAddSpacing = false;
        if(table.Content.length > 0)
        {
            var firstRow = table.Content[0];
			var rowPr = firstRow.Get_CompiledPr();
            if(null != rowPr && null != rowPr.TableCellSpacing)
            {
                bAddSpacing = true;
                var cellSpacingMM = rowPr.TableCellSpacing;
                tblStyle += "mso-cellspacing:"+(cellSpacingMM * g_dKoef_mm_to_pt)+"pt;";
				DomTable.oAttributes["cellspacing"] = Math.round(cellSpacingMM * g_dKoef_mm_to_pix);
            }
        }
        if(!bAddSpacing)
            DomTable.oAttributes["cellspacing"] = 0;
        DomTable.oAttributes["border"] = false == bBorder ? 0 : 1;
        DomTable.oAttributes["cellpadding"] = 0;
        if("" != tblStyle)
            DomTable.oAttributes["style"] = tblStyle;

        //rows
		table.Recalculate_Grid();
		if(null == aRowElems)
		{
			for(var i = 0, length = table.Content.length; i < length; i++)
				this.CopyRow(DomTable, table, i , null, table.Content.length - 1);
		}
		else
		{
			var nMaxRow = 0;
			for(var i = 0, length = aRowElems.length; i < length; ++i)
			{
				var elem = aRowElems[i];
				if(elem.row > nMaxRow)
					nMaxRow = elem.row;
			}
			for(var i = 0, length = aRowElems.length; i < length; ++i)
			{
				var elem = aRowElems[i];
				this.CopyRow(DomTable, table, elem.row, elem, nMaxRow);
			}
		}

        oDomTarget.addChild(DomTable);
    },

	CopyDocument2 : function(oDomTarget, oDocument, elementsContent, dNotGetBinary)
	{
		if(PasteElementsId.g_bIsDocumentCopyPaste)
		{
			if(!elementsContent && oDocument && oDocument.Content)
				elementsContent = oDocument.Content;

			for ( var Index = 0; Index < elementsContent.length; Index++ )
			{
				var Item;
				if(elementsContent[Index].Element)
					Item = elementsContent[Index].Element;
				else
					Item = elementsContent[Index];

				if(type_Table === Item.GetType() )
				{
					this.oBinaryFileWriter.copyParams.bLockCopyElems++;
					if(!this.onlyBinaryCopy)
						this.CopyTable(oDomTarget, Item, null);
					this.oBinaryFileWriter.copyParams.bLockCopyElems--;

					if(!dNotGetBinary)
						this.oBinaryFileWriter.CopyTable(Item, null);
				}
				else if ( type_Paragraph === Item.GetType() )
				{
					var SelectedAll = Index === elementsContent.length - 1 ? elementsContent[Index].SelectedAll : true;
					//todo может только для верхнего уровня надо Index == End
					if (!dNotGetBinary) {
						this.oBinaryFileWriter.CopyParagraph(Item, SelectedAll);
					}

					if (!this.onlyBinaryCopy) {
						var _nextElem;
						if (elementsContent[Index + 1] && elementsContent[Index + 1].Element) {
							_nextElem = elementsContent[Index + 1].Element;
						} else {
							_nextElem = elementsContent[Index + 1];
						}
						this.CopyParagraph(oDomTarget, Item, SelectedAll, _nextElem);
					}
				}
				else if(type_BlockLevelSdt === Item.GetType() )
				{
					this.oBinaryFileWriter.copyParams.bLockCopyElems++;
					if(!this.onlyBinaryCopy)
					{
						this.CopyDocument2(oDomTarget, oDocument, Item.Content.Content, true);
					}
					this.oBinaryFileWriter.copyParams.bLockCopyElems--;

					if(!dNotGetBinary)
						this.oBinaryFileWriter.CopySdt(Item);
				}
			}
		}
		else//presentation
		{
			this.copyPresentation2(oDomTarget, oDocument, elementsContent);
		}
    },

	copyPresentation2: function(oDomTarget, oDocument, elementsContent){
		//DocContent/ Drawings/ SlideObjects
		var presentation = this.oDocument;

		if(elementsContent && elementsContent.length){
			if(elementsContent[0].DocContent || (elementsContent[0].Drawings && elementsContent[0].Drawings.length) || (elementsContent[0].SlideObjects && elementsContent[0].SlideObjects.length))
			{
				var themeName = elementsContent[0].ThemeName ? elementsContent[0].ThemeName : "";

				this.oPresentationWriter.WriteString2(this.api.documentId);
				this.oPresentationWriter.WriteString2(themeName);
				this.oPresentationWriter.WriteDouble(presentation.GetWidthMM());
				this.oPresentationWriter.WriteDouble(presentation.GetHeightMM());
				//флаг о том, что множественный контент в буфере
				this.oPresentationWriter.WriteBool(true);
			}

			//записываем все варианты контента
			//в html записываем первый вариант - конечное форматирование
			//в банарник пишем: 1)конечное форматирование 2)исходное форматирование 3)картинка
			this.oPresentationWriter.WriteULong(elementsContent.length);
			for(var i = 0; i < elementsContent.length; i++)
			{
				if(i === 0)
				{
					this.copyPresentationContent(elementsContent[i], oDomTarget);
				}
				else
				{
					this.copyPresentationContent(elementsContent[i]);
				}
			}
		}
		/*else if(elementsContent)
		{
			//эту ветку оставляю для записи едиственного варианта контента, который используется функцией getSelectedBinary
			if(elementsContent.DocContent || (elementsContent.Drawings && elementsContent.Drawings.length) || (elementsContent.SlideObjects && elementsContent.SlideObjects.length))
			{
				this.oPresentationWriter.WriteString2(this.api.documentId);
				this.oPresentationWriter.WriteDouble(presentation.GetWidth());
				this.oPresentationWriter.WriteDouble(presentation.GetHeight());
			}
			this.copyPresentationContent(elementsContent, oDomTarget);
		}*/
		else
		{
			//для записи внутреннего контента таблицы
			this.copyPresentationContent(oDocument, oDomTarget);
		}
	},

	copyPresentationContent: function (elementsContent, oDomTarget) {

		if(elementsContent instanceof PresentationSelectedContent){
			this._writePresentationSelectedContent(elementsContent, oDomTarget);
		}
		else
		{
			//inner recursive call CopyDocument2 function
			if (elementsContent && elementsContent.Content && elementsContent.Content.length) {//пишем таблицу в html

				for (var Index = 0; Index < elementsContent.Content.length; Index++) {
					var Item = elementsContent.Content[Index];

					if (type_Table === Item.GetType()) {
						this.CopyTable(oDomTarget, Item, null);
					} else if (type_Paragraph === Item.GetType()) {
						this.CopyParagraph(oDomTarget, Item, true);
					}
				}

			}
		}
	},

	_writePresentationSelectedContent: function(elementsContent, oDomTarget){

		var oThis = this;
		var copyDocContent = function(){
			var docContent = elementsContent.DocContent;

			if (docContent.Elements) {
				var elements = docContent.Elements;

				//пишем метку и длины
				oThis.oPresentationWriter.WriteString2("DocContent");
				oThis.oPresentationWriter.WriteDouble(elements.length);

				//пишем контент
				for (var Index = 0; Index < elements.length; Index++) {
					var Item;
					if (elements[Index].Element) {
						Item = elements[Index].Element;
					} else {
						Item = elements[Index];
					}

					if (type_Paragraph === Item.GetType()) {
						oThis.oPresentationWriter.StartRecord(elements[Index].SelectedAll ? 1 : 0);
						oThis.oPresentationWriter.WriteParagraph(Item);
						oThis.oPresentationWriter.EndRecord();

						if (oDomTarget) {
							oThis.CopyParagraph(oDomTarget, Item, true);
						}
					}
				}
			}
		};

		var copyDrawings = function(){
			var elements = elementsContent.Drawings;

			//var selected_objects = graphicObjects.State.id === STATES_ID_GROUP ? graphicObjects.State.group.selectedObjects : graphicObjects.selectedObjects;

			//пишем метку и длины
			oThis.oPresentationWriter.WriteString2("Drawings");
			oThis.oPresentationWriter.WriteULong(elements.length);

			pptx_content_writer.Start_UseFullUrl();
			for (var i = 0; i < elements.length; ++i) {
				if (!elements[i].Drawing.isTable()) {
					oThis.oPresentationWriter.WriteBool(true);

					oThis.CopyGraphicObject(oDomTarget, elements[i].Drawing, elements[i]);

					oThis.oPresentationWriter.WriteDouble(elements[i].X);
					oThis.oPresentationWriter.WriteDouble(elements[i].Y);
					oThis.oPresentationWriter.WriteDouble(elements[i].ExtX);
					oThis.oPresentationWriter.WriteDouble(elements[i].ExtY);
					//TODO записывать base64 у картинок для разных контентов в единственном экземпляре
					if(elements[i].Drawing.isImage()) {
						oThis.oPresentationWriter.WriteString2("");
					} else {
						oThis.oPresentationWriter.WriteString2(elements[i].ImageUrl);
					}
				} else {
					var isOnlyTable = elements.length === 1;

					oThis.CopyPresentationTableFull(oDomTarget, elements[i].Drawing, isOnlyTable);

					oThis.oPresentationWriter.WriteDouble(elements[i].X);
					oThis.oPresentationWriter.WriteDouble(elements[i].Y);
					oThis.oPresentationWriter.WriteDouble(elements[i].ExtX);
					oThis.oPresentationWriter.WriteDouble(elements[i].ExtY);
					oThis.oPresentationWriter.WriteString2(elements[i].ImageUrl);
				}
			}
			pptx_content_writer.End_UseFullUrl();

		};

		var copySlideObjects = function(){
			var selected_slides = elementsContent.SlideObjects;

			oThis.oPresentationWriter.WriteString2("SlideObjects");
			oThis.oPresentationWriter.WriteULong(selected_slides.length);

			var layouts_map = {};
			var layout_count = 0;
			editor.WordControl.m_oLogicDocument.CalculateComments();

			//пишем слайд
			var slide;
			for (var i = 0; i < selected_slides.length; ++i) {
				slide = selected_slides[i];
				if(i === 0){
					oThis.CopySlide(oDomTarget, slide);
				} else{
					oThis.CopySlide(null, slide);
				}
			}
		};

		var copyLayouts = function(){
			var selected_layouts = elementsContent.Layouts;

			oThis.oPresentationWriter.WriteString2("Layouts");
			oThis.oPresentationWriter.WriteULong(selected_layouts.length);

			for (var i = 0; i < selected_layouts.length; ++i) {
				oThis.CopyLayout(selected_layouts[i]);
			}

		};

		var copyMasters = function(){
			var selected_masters = elementsContent.Masters;

			oThis.oPresentationWriter.WriteString2("Masters");
			oThis.oPresentationWriter.WriteULong(selected_masters.length);

			for (var i = 0; i < selected_masters.length; ++i) {
				oThis.oPresentationWriter.WriteSlideMaster(selected_masters[i]);
			}
		};

		var copyNotes = function(){
			var selected_notes = elementsContent.Notes;

			oThis.oPresentationWriter.WriteString2("Notes");
			oThis.oPresentationWriter.WriteULong(selected_notes.length);

			for (var i = 0; i < selected_notes.length; ++i) {
				oThis.oPresentationWriter.WriteSlideNote(selected_notes[i]);
			}
		};

		var copyNoteMasters = function(){
			var selected_note_master = elementsContent.NotesMasters;

			oThis.oPresentationWriter.WriteString2("NotesMasters");
			oThis.oPresentationWriter.WriteULong(selected_note_master.length);

			for (var i = 0; i < selected_note_master.length; ++i) {
				oThis.oPresentationWriter.WriteNoteMaster(selected_note_master[i]);
			}
		};

		var copyNoteTheme = function(){
			var selected_themes = elementsContent.NotesThemes;

			oThis.oPresentationWriter.WriteString2("NotesThemes");
			oThis.oPresentationWriter.WriteULong(selected_themes.length);

			for (var i = 0; i < selected_themes.length; ++i) {
				oThis.oPresentationWriter.WriteTheme(selected_themes[i]);
			}
		};

		var copyTheme = function(){
			var selected_themes = elementsContent.Themes;

			oThis.oPresentationWriter.WriteString2("Themes");
			oThis.oPresentationWriter.WriteULong(selected_themes.length);

			for (var i = 0; i < selected_themes.length; ++i) {
				oThis.oPresentationWriter.WriteTheme(selected_themes[i]);
			}
		};

		var copyIndexes = function(selected_indexes){
			oThis.oPresentationWriter.WriteULong(selected_indexes.length);
			for (var i = 0; i < selected_indexes.length; ++i) {
				oThis.oPresentationWriter.WriteULong(selected_indexes[i]);
			}
		};


		//получаем пишем количество
		var contentCount = 0;
		for(var i in elementsContent){
			if(elementsContent[i] && typeof elementsContent[i] === "object" && elementsContent[i].length){
				contentCount++;
			} else if(null !== elementsContent[i] && elementsContent[i] instanceof AscCommonWord.CSelectedContent){
				contentCount++;
			}
		}
		oThis.oPresentationWriter.WriteString2("SelectedContent");
        oThis.oPresentationWriter.WriteULong((elementsContent.PresentationWidth * 100000) >> 0);
        oThis.oPresentationWriter.WriteULong((elementsContent.PresentationHeight * 100000) >> 0);
		oThis.oPresentationWriter.WriteULong(contentCount);

		//DocContent
		if (elementsContent.DocContent) {//пишем контент
			copyDocContent();
		}
		//Drawings
		if (elementsContent.Drawings && elementsContent.Drawings.length) {
			copyDrawings();
		}
		//SlideObjects
		if (elementsContent.SlideObjects && elementsContent.SlideObjects.length) {
			copySlideObjects();
		}
		//Layouts
		if (elementsContent.Layouts && elementsContent.Layouts.length) {
			copyLayouts();
		}
		//LayoutsIndexes
		if (elementsContent.LayoutsIndexes && elementsContent.LayoutsIndexes.length) {
			oThis.oPresentationWriter.WriteString2("LayoutsIndexes");
			copyIndexes(elementsContent.LayoutsIndexes);
		}
		//Masters
		if (elementsContent.Masters && elementsContent.Masters.length) {
			copyMasters();
		}
		//MastersIndexes
		if (elementsContent.MastersIndexes && elementsContent.MastersIndexes.length) {
			oThis.oPresentationWriter.WriteString2("MastersIndexes");
			copyIndexes(elementsContent.MastersIndexes);
		}
		//Notes
		if (elementsContent.Notes && elementsContent.Notes.length) {
			//TODO если нет Notes, то Notes должен быть равен null. вместо этого приходится проверять 1 элемент массива
			if(!(elementsContent.Notes.length === 1 && null === elementsContent.Notes[0])){
				copyNotes();
			}
		}
		//NotesMasters
		if (elementsContent.NotesMasters && elementsContent.NotesMasters.length) {
			copyNoteMasters();
		}
		//NotesMastersIndexes
		if (elementsContent.NotesMastersIndexes && elementsContent.NotesMastersIndexes.length) {
			oThis.oPresentationWriter.WriteString2("NotesMastersIndexes");
			copyIndexes(elementsContent.NotesMastersIndexes);
		}
		//NotesThemes
		if (elementsContent.NotesThemes && elementsContent.NotesThemes.length) {
			copyNoteTheme();
		}
		//Themes
		if (elementsContent.Themes && elementsContent.Themes.length) {
			copyTheme();
		}
		//ThemesIndexes
		if (elementsContent.ThemesIndexes && elementsContent.ThemesIndexes.length) {
			oThis.oPresentationWriter.WriteString2("ThemeIndexes");
			copyIndexes(elementsContent.ThemesIndexes);
		}
	},

	getSelectedBinary : function()
	{
		var oDocument = this.oDocument;

		if(PasteElementsId.g_bIsDocumentCopyPaste)
		{
			var selectedContent = oDocument.GetSelectedContent();

			var elementsContent;
			if(selectedContent && selectedContent.Elements && selectedContent.Elements[0] && selectedContent.Elements[0].Element)
				elementsContent = selectedContent.Elements;
			else
				return false;

			var drawingUrls = [];
			if(selectedContent.DrawingObjects && selectedContent.DrawingObjects.length)
			{
				var url, correctUrl, graphicObj;
				for(var i = 0; i < selectedContent.DrawingObjects.length; i++)
				{
					graphicObj = selectedContent.DrawingObjects[i].GraphicObj;
					if(graphicObj.isImage())
					{
						url = graphicObj.getImageUrl();
						if(window["NativeCorrectImageUrlOnCopy"])
						{
							correctUrl = window["NativeCorrectImageUrlOnCopy"](url);

							drawingUrls[i] = correctUrl;
						}
					}
				}
			}

			//подменяем Document для копирования(если не подменить, то commentId будет не соответствовать)
			this.oBinaryFileWriter.Document = elementsContent[0].Element.LogicDocument;

			this.oBinaryFileWriter.CopyStart();
			this.CopyDocument2(null, oDocument, elementsContent);
			this.oBinaryFileWriter.CopyEnd();

			var sBase64 = this.oBinaryFileWriter.GetResult();
			var text = "";
            if (oDocument.GetSelectedText)
                text = oDocument.GetSelectedText();

			return {sBase64: "docData;" + sBase64, text: text, drawingUrls: drawingUrls};
		}
	},

	Start : function () {
		var oDocument = this.oDocument;
		var bFromPresentation;

		window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();

		var sBase64, oElem, sStyle;
		var selectedContent;
		if (PasteElementsId.g_bIsDocumentCopyPaste) {
			selectedContent = oDocument.GetSelectedContent();

			var elementsContent;
			if (selectedContent && selectedContent.Elements && selectedContent.Elements[0] && selectedContent.Elements[0].Element) {
				elementsContent = selectedContent.Elements;
			} else {
				return "";
			}

			//TODO заглушка для презентационных параграфов(выделен текст внутри диаграммы) - пока не пишем в бинарник
			if (selectedContent.Elements[0].Element && selectedContent.Elements[0].Element.bFromDocument === false) {
				this.oBinaryFileWriter.Document = this.oDocument;
			} else {
				//подменяем Document для копирования(если не подменить, то commentId будет не соответствовать)
				this.oBinaryFileWriter.Document = elementsContent[0].Element.LogicDocument;
			}

			this.oBinaryFileWriter.CopyStart();
			this.CopyDocument2(this.oRoot, oDocument, elementsContent, bFromPresentation);
			this.CopyFootnotes(this.oRoot, this.aFootnoteReference);
			this.oBinaryFileWriter.CopyEnd();
		} else {
			selectedContent = oDocument.GetSelectedContent2();
			if (!selectedContent[0].DocContent && (!selectedContent[0].Drawings ||
				(selectedContent[0].Drawings && !selectedContent[0].Drawings.length)) &&
				(!selectedContent[0].SlideObjects ||
				(selectedContent[0].SlideObjects && !selectedContent[0].SlideObjects.length))) {
				return false;
			}

			this.CopyDocument2(this.oRoot, oDocument, selectedContent);

			sBase64 = this.oPresentationWriter.GetBase64Memory();
			sBase64 = "pptData;" + this.oPresentationWriter.pos + ";" + sBase64;
			if (this.oRoot.aChildren && this.oRoot.aChildren.length === 1 && AscBrowser.isSafariMacOs) {
				oElem = this.oRoot.aChildren[0];
				sStyle = oElem.oAttributes["style"];
				if (null == sStyle) {
					oElem.oAttributes["style"] = "font-weight:normal";
				} else {
					oElem.oAttributes["style"] = sStyle + ";font-weight:normal";
				}//просто добавляем потому что в sStyle не могло быть font-weight, мы всегда пишем <b>
				this.oRoot.wrapChild(new CopyElement("b"));
			}
			if (this.oRoot.aChildren && this.oRoot.aChildren.length > 0) {
				this.oRoot.aChildren[0].oAttributes["class"] = sBase64;
			}
		}

		if (PasteElementsId.g_bIsDocumentCopyPaste && PasteElementsId.copyPasteUseBinary && this.oBinaryFileWriter.copyParams.itemCount > 0 && !bFromPresentation) {
			sBase64 = "docData;" + this.oBinaryFileWriter.GetResult();
			if (this.oRoot.aChildren && this.oRoot.aChildren.length == 1 && AscBrowser.isSafariMacOs) {
				oElem = this.oRoot.aChildren[0];
				sStyle = oElem.oAttributes["style"];
				if (null == sStyle) {
					oElem.oAttributes["style"] = "font-weight:normal";
				} else {
					oElem.oAttributes["style"] = sStyle + ";font-weight:normal";
				}//просто добавляем потому что в sStyle не могло быть font-weight, мы всегда пишем <b>
				this.oRoot.wrapChild(new CopyElement("b"));
			}
			if (this.oRoot.aChildren && this.oRoot.aChildren.length > 0) {
				this.oRoot.aChildren[0].oAttributes["class"] = sBase64;
			}
		}

		return sBase64;
	},


	CopySlide: function (oDomTarget, slide) {
		if (oDomTarget) {
			var sSrc = slide.getBase64Img();
			var _bounds_cheker = new AscFormat.CSlideBoundsChecker();
			slide.draw(_bounds_cheker, 0);
			var oImg = new CopyElement("img");
			oImg.oAttributes["width"] = Math.round((_bounds_cheker.Bounds.max_x - _bounds_cheker.Bounds.min_x + 1) * g_dKoef_mm_to_pix);
			oImg.oAttributes["height"] = Math.round((_bounds_cheker.Bounds.max_y - _bounds_cheker.Bounds.min_y + 1) * g_dKoef_mm_to_pix);
			oImg.oAttributes["src"] = sSrc;
			oDomTarget.addChild(oImg);
		}

		//пока записываю для копирования/вставки ссылку на стиль
		//TODO в дальнейшем необходимо пересмотреть и писать стили вместе со слайдом
		// - аналогично тому как это реализовано при записи таблицы
		var presentation = editor.WordControl.m_oLogicDocument;
		for(var key in presentation.TableStylesIdMap)
		{
			if(presentation.TableStylesIdMap.hasOwnProperty(key))
			{
				this.oPresentationWriter.tableStylesGuides[key] = key;
			}
		}

		//записываем slide
		this.oPresentationWriter.WriteSlide(slide);

	},

	CopyLayout: function (layout) {
		this.oPresentationWriter.WriteSlideLayout(layout);
	},


	CopyPresentationTableCells: function (oDomTarget, graphicFrame) {
		var aSelectedRows = [];
		var oRowElems = {};
		var Item = graphicFrame.graphicObject;
		if (Item.Selection.Data.length > 0) {
			for (var i = 0, length = Item.Selection.Data.length; i < length; ++i) {
				var elem = Item.Selection.Data[i];
				var rowElem = oRowElems[elem.Row];
				if (null == rowElem) {
					rowElem = {
						row: elem.Row,
						gridStart: null,
						gridEnd: null,
						indexStart: null,
						indexEnd: null,
						after: null,
						before: null,
						cells: {}
					};
					oRowElems[elem.Row] = rowElem;
					aSelectedRows.push(rowElem);
				}
				if (null == rowElem.indexEnd || elem.Cell > rowElem.indexEnd)
					rowElem.indexEnd = elem.Cell;
				if (null == rowElem.indexStart || elem.Cell < rowElem.indexStart)
					rowElem.indexStart = elem.Cell;
				rowElem.cells[elem.Cell] = 1;
			}
		}
		aSelectedRows.sort(function (a, b) {
			return a.row - b.row;
		});
		var nMinGrid = null;
		var nMaxGrid = null;
		var nPrevStartGrid = null;
		var nPrevEndGrid = null;
		var nPrevRowIndex = null;
		for (var i = 0, length = aSelectedRows.length; i < length; ++i) {
			var elem = aSelectedRows[i];
			var nRowIndex = elem.row;
			if (null != nPrevRowIndex) {
				if (nPrevRowIndex + 1 !== nRowIndex) {
					nMinGrid = null;
					nMaxGrid = null;
					break;
				}
			}
			nPrevRowIndex = nRowIndex;
			var row = Item.Content[nRowIndex];
			var cellFirst = row.Get_Cell(elem.indexStart);
			var cellLast = row.Get_Cell(elem.indexEnd);
			var nCurStartGrid = cellFirst.Metrics.StartGridCol;
			var nCurEndGrid = cellLast.Metrics.StartGridCol + cellLast.Get_GridSpan() - 1;
			if (null != nPrevStartGrid && null != nPrevEndGrid) {
				//учитываем вертикальный merge, раздвигаем границы
				if (nCurStartGrid > nPrevStartGrid) {
					for (var j = elem.indexStart - 1; j >= 0; --j) {
						var cellCur = row.Get_Cell(j);
						if (vmerge_Continue === cellCur.GetVMerge()) {
							var nCurGridCol = cellCur.Metrics.StartGridCol;
							if (nCurGridCol >= nPrevStartGrid) {
								nCurStartGrid = nCurGridCol;
								elem.indexStart = j;
							} else
								break;
						} else
							break;
					}
				}
				if (nCurEndGrid < nPrevEndGrid) {
					for (var j = elem.indexEnd + 1; j < row.Get_CellsCount(); ++j) {
						var cellCur = row.Get_Cell(j);
						if (vmerge_Continue === cellCur.GetVMerge()) {
							var nCurGridCol = cellCur.Metrics.StartGridCol + cellCur.Get_GridSpan() - 1;
							if (nCurGridCol <= nPrevEndGrid) {
								nCurEndGrid = nCurGridCol;
								elem.indexEnd = j;
							} else
								break;
						} else
							break;
					}
				}
			}
			elem.gridStart = nPrevStartGrid = nCurStartGrid;
			elem.gridEnd = nPrevEndGrid = nCurEndGrid;
			if (null == nMinGrid || nMinGrid > nCurStartGrid)
				nMinGrid = nCurStartGrid;
			if (null == nMaxGrid || nMaxGrid < nCurEndGrid)
				nMaxGrid = nCurEndGrid;
		}
		if (null != nMinGrid && null != nMaxGrid) {
			//выставляем after, before
			for (var i = 0, length = aSelectedRows.length; i < length; ++i) {
				var elem = aSelectedRows[i];
				elem.before = elem.gridStart - nMinGrid;
				elem.after = nMaxGrid - elem.gridEnd;
			}
			this.CopyTable(oDomTarget, Item, aSelectedRows);
		}
		History.TurnOff();
		var graphic_frame = new AscFormat.CGraphicFrame(graphicFrame.parent);
		var grid = [];

		for (var i = nMinGrid; i <= nMaxGrid; ++i) {
			grid.push(graphicFrame.graphicObject.TableGrid[i]);
		}
		var table = new CTable(editor.WordControl.m_oDrawingDocument, graphicFrame, false, aSelectedRows.length, nMaxGrid - nMinGrid + 1, grid);
		table.setStyleIndex(graphicFrame.graphicObject.styleIndex);
		graphic_frame.setGraphicObject(table);
		graphic_frame.setXfrm(0, 0, 20, 30, 0, false, false);
		var b_style_index = false;
		if (AscFormat.isRealNumber(graphic_frame.graphicObject.styleIndex) && graphic_frame.graphicObject.styleIndex > -1) {
			b_style_index = true;
		}

		this.oPresentationWriter.WriteULong(1);
		this.oPresentationWriter.WriteBool(false);
		this.oPresentationWriter.WriteBool(b_style_index);
		if (b_style_index) {
			this.oPresentationWriter.WriteULong(graphic_frame.graphicObject.styleIndex);
		}
		var old_style_index = graphic_frame.graphicObject.styleIndex;
		graphic_frame.graphicObject.styleIndex = -1;
		this.oPresentationWriter.WriteGrFrame(graphic_frame);
		graphic_frame.graphicObject.styleIndex = old_style_index;

		History.TurnOn();

		this.oBinaryFileWriter.copyParams.itemCount = 0;
	},

	CopyPresentationTableFull: function (oDomTarget, graphicFrame, isOnlyTable) {
		var aSelectedRows = [];
		var oRowElems = {};
		var Item = graphicFrame.graphicObject;

		var b_style_index = false;
		var presentation = editor.WordControl.m_oLogicDocument;
		if (Item.TableStyle && presentation.globalTableStyles.Style[Item.TableStyle]) {
			b_style_index = true;
		}

		for (var key in presentation.TableStylesIdMap) {
			if (presentation.TableStylesIdMap.hasOwnProperty(key)) {
				this.oPresentationWriter.tableStylesGuides[key] = "{" + AscCommon.GUID() + "}"
			}
		}


		this.oPresentationWriter.WriteBool(!b_style_index);
		if (b_style_index) {
			var tableStyle = presentation.globalTableStyles.Style[Item.TableStyle];
			this.oPresentationWriter.WriteBool(true);
			this.oPresentationWriter.WriteTableStyle(Item.TableStyle, tableStyle);
			this.oPresentationWriter.WriteBool(true);
			this.oPresentationWriter.WriteString2(Item.TableStyle);
		}

		History.TurnOff();
		this.oPresentationWriter.WriteGrFrame(graphicFrame);

		//для случая, когда копируем 1 таблицу из презентаций, в бинарник заносим ещё одну такую же табличку, но со скомпиоированными стилями(для вставки в word / excel)
		if (isOnlyTable) {
			this.convertToCompileStylesTable(Item);
			this.oPresentationWriter.WriteGrFrame(graphicFrame);
		}
		History.TurnOn();

		if (oDomTarget) {
			this.CopyTable(oDomTarget, Item, null);
		}
	},

	convertToCompileStylesTable: function (table) {
		var t = this;

		for (var i = 0; i < table.Content.length; i++) {
			var row = table.Content[i];
			for (var j = 0; j < row.Content.length; j++) {
				var cell = row.Content[j];
				var compilePr = cell.Get_CompiledPr();

				cell.Pr = compilePr;

				var shd = compilePr.Shd;
				var color = shd.Get_Color2(this.oDocument.Get_Theme(), this.oDocument.Get_ColorMap());
				if(color) {
					cell.Pr.Shd.Unifill = AscFormat.CreateSolidFillRGB(color.r, color.g, color.b);
				}

				if (compilePr.TableCellBorders.Bottom) {
					color = compilePr.TableCellBorders.Bottom.Get_Color2(this.oDocument.Get_Theme(), this.oDocument.Get_ColorMap());
					if(color) {
						cell.Pr.TableCellBorders.Bottom.Unifill = AscFormat.CreateSolidFillRGB(color.r, color.g, color.b);
					}
				}

				if (compilePr.TableCellBorders.Top) {
					color = compilePr.TableCellBorders.Top.Get_Color2(this.oDocument.Get_Theme(), this.oDocument.Get_ColorMap());
					if(color) {
						cell.Pr.TableCellBorders.Top.Unifill = AscFormat.CreateSolidFillRGB(color.r, color.g, color.b);
					}
				}

				if (compilePr.TableCellBorders.Left) {
					color = compilePr.TableCellBorders.Left.Get_Color2(this.oDocument.Get_Theme(), this.oDocument.Get_ColorMap());
					if(color) {
						cell.Pr.TableCellBorders.Left.Unifill = AscFormat.CreateSolidFillRGB(color.r, color.g, color.b);
					}
				}

				if (compilePr.TableCellBorders.Right) {
					color = compilePr.TableCellBorders.Right.Get_Color2(this.oDocument.Get_Theme(), this.oDocument.Get_ColorMap());
					if(color) {
						cell.Pr.TableCellBorders.Right.Unifill = AscFormat.CreateSolidFillRGB(color.r, color.g, color.b);
					}
				}
			}
		}
	},

	CopyGraphicObject: function (oDomTarget, oGraphicObj, drawingCopyObject) {
		var sSrc = drawingCopyObject.ImageUrl;
		if (oDomTarget && sSrc.length > 0) {
			var _bounds_cheker = new AscFormat.CSlideBoundsChecker();
			oGraphicObj.draw(_bounds_cheker, 0);

			var width, height;
			if (drawingCopyObject && drawingCopyObject.ExtX)
				width = Math.round(drawingCopyObject.ExtX * g_dKoef_mm_to_pix);
			else
				width = Math.round((_bounds_cheker.Bounds.max_x - _bounds_cheker.Bounds.min_x + 1) * g_dKoef_mm_to_pix);

			if (drawingCopyObject && drawingCopyObject.ExtY)
				height = Math.round(drawingCopyObject.ExtY * g_dKoef_mm_to_pix);
			else
				height = Math.round((_bounds_cheker.Bounds.max_y - _bounds_cheker.Bounds.min_y + 1) * g_dKoef_mm_to_pix);

			var oImg = new CopyElement("img");
			oImg.oAttributes["width"] = width;
			oImg.oAttributes["height"] = height;
			oImg.oAttributes["src"] = sSrc;
			if (this.api.DocumentReaderMode)
				oImg.oAttributes["style"] = "max-width:100%;";
			oDomTarget.addChild(oImg);
		}
		this.oPresentationWriter.WriteSpTreeElem(oGraphicObj);
	},

	CopyFootnotes: function (oDomTarget, aFootnotes) {
		if (aFootnotes && aFootnotes.length) {

			/*<div style='mso-element:footnote-list'>
				<br clear=all>
				<hr align=left size=1 width="33%">
				<div style='mso-element:footnote' id=ftn1>
					<p class=MsoFootnoteText>
						<a style='mso-footnote-id:ftn1' href="#_ftnref1" name="_ftn1" title="">
							<span class=MsoFootnoteReference>
								<span style='mso-special-character:footnote'>
									<span class=MsoFootnoteReference>
										<span style=''>[1]</span>
									</span>
								</span>
							</span>
						</a>
						Link here
					</p>
				</div>
				<div style="mso-element:footnote" id="ftn2">
					<p class="MsoFootnoteText">
						<a style="mso-footnote-id:ftn2" href="#_ftnref2" name="_ftn2" title="">
							<span class="MsoFootnoteReference">
								<span style="mso-special-character:footnote">
									<span class="MsoFootnoteReference">
										<span style="">[2]</span>
									</span>
								</span>
							</span>
						</a>
						<span><b><i><u>Sdfsfsdsdf</u></i></b></span>
					</p>
					<p class="MsoFootnoteText"><span>Sdfsdfsdfsdf</span></p>
					<p class="MsoFootnoteText"><span>sdfsdfsdfsdfsdf</span></p>
				</div>
			</div>*/

			let _mainDiv = new CopyElement("div");
			_mainDiv.oAttributes["style"] = "mso-element:footnote-list";

			let _br = new CopyElement("br");
			_br.oAttributes["clear"] = "all";
			_mainDiv.addChild(_br);

			let _hr = new CopyElement("hr");
			_hr.oAttributes["align"] = "left";
			_hr.oAttributes["size"] = "1";
			_hr.oAttributes["width"] = "33%";
			_mainDiv.addChild(_hr);

			for (let i = 0; i < aFootnotes.length; i++) {
				let prefix = "ftn";
				let index = i + 1;
				let _div = new CopyElement("div");
				_div.oAttributes["style"] = "mso-element:footnote";
				_div.oAttributes["id"] = prefix + index;

				if (!aFootnotes[i] || !aFootnotes[i].Content) {
					continue;
				}

				//in first paragraph put link and paragraphs contents
				for (let j = 0; j < aFootnotes[i].Content.length; j++) {
					let _p = new CopyElement("p");
					_p.oAttributes["class"] = "MsoFootnoteText";

					let _link;
					if (j === 0) {

						/*<a style="mso-footnote-id:ftn2" href="#_ftnref2" name="_ftn2" title="">
							<span class="MsoFootnoteReference">
								<span style="mso-special-character:footnote">
									<span class="MsoFootnoteReference">
										<span style="">[2]</span>
									</span>
								</span>
							</span>
						</a>*/

						_link = new CopyElement("a");

						_link.oAttributes["style"] = "mso-footnote-id:" + prefix + index;
						_link.oAttributes["href"] = "#_" + prefix + "ref" + index;
						_link.oAttributes["name"] = "_" + prefix + index;
						_link.oAttributes["title"] = "";

						//skip 2 inner spans(MsoFootnoteReference + last)
						let spanMsoFootnoteReference = new CopyElement("span");
						spanMsoFootnoteReference.oAttributes["class"] = "MsoFootnoteReference";
						let spanMsoSpecialCharacter = new CopyElement("span");
						spanMsoSpecialCharacter.oAttributes["style"] = "mso-special-character:footnote";

						spanMsoFootnoteReference.addChild(spanMsoSpecialCharacter);
						spanMsoFootnoteReference.addChild(new CopyElement(CopyPasteCorrectString("[" + index + "]"), true));

						_link.addChild(spanMsoFootnoteReference);
					}

					if (_link) {
						_p.addChild(_link);
						//add spans from aFootnotes[0]
						let container = new CopyElement("div");
						this.CopyDocument2(container, null, [aFootnotes[i].Content[j]], true);
						for (let i = 0; i < container.aChildren.length; i++) {
							container.aChildren[i].moveChildTo(_p);
						}
					} else {
						this.CopyDocument2(_p, null, [aFootnotes[i].Content[j]], true);
					}

					_div.addChild(_p);
				}
				_mainDiv.addChild(_div);
			}

			oDomTarget.addChild(_mainDiv);
		}
	}
};

function CopyPasteCorrectString(str)
{
    /*
    // эта реализация на порядок быстрее. Перед выпуском не меняю ничего
    var _ret = "";
    var _len = str.length;

    for (var i = 0; i < _len; i++)
    {
        var _symbol = str[i];
        if (_symbol == "&")
            _ret += "&amp;";
        else if (_symbol == "<")
            _ret += "&lt;";
        else if (_symbol == ">")
            _ret += "&gt;";
        else if (_symbol == "'")
            _ret += "&apos;";
        else if (_symbol == "\"")
            _ret += "&quot;";
        else
            _ret += _symbol;
    }

    return _ret;
    */
    if (!str)
    	return "";

    var res = str;
    res = res.replace(/&/g,'&amp;');
    res = res.replace(/</g,'&lt;');
    res = res.replace(/>/g,'&gt;');
    res = res.replace(/'/g,'&apos;');
    res = res.replace(/"/g,'&quot;');
    return res;
}

function GetContentFromHtml(api, html, callback) {
	if (!html) {
		callback && callback(null);
		return;
	}

	//need document -> write, because some props not compile if use innerHTML(sub/sup tags give only "baseline")
	//TODO use iframe from clipboard_base
	AscCommon.g_clipboardBase.CommonIframe_PasteStart(html, null, function (oHtmlElem) {
		if (oHtmlElem) {
			var oPasteProcessor = new PasteProcessor(api, true, true, false, undefined, function (selectedContent) {
				if (selectedContent) {
					callback(selectedContent);
				}
			});
			oPasteProcessor.doNotInsertInDoc = true;
			oHtmlElem && oPasteProcessor.Start(oHtmlElem);
		}
	});
}

function Editor_Paste_Exec(api, _format, data1, data2, text_data, specialPasteProps, callback)
{
    var oPasteProcessor = new PasteProcessor(api, true, true, false, undefined, callback);
	window['AscCommon'].g_specialPasteHelper.endRecalcDocument = false;

	if(undefined === specialPasteProps)
	{
		window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
		window['AscCommon'].g_specialPasteHelper.specialPasteData._format = _format;
		window['AscCommon'].g_specialPasteHelper.specialPasteData.data1 = data1;
		window['AscCommon'].g_specialPasteHelper.specialPasteData.data2 = data2;
		window['AscCommon'].g_specialPasteHelper.specialPasteData.text_data = text_data;
		if (_format !== AscCommon.c_oAscClipboardDataFormat.Text) {
			text_data = null;
		}
	}
	else
	{
		window['AscCommon'].g_specialPasteHelper.specialPasteProps = specialPasteProps;

		_format = window['AscCommon'].g_specialPasteHelper.specialPasteData._format;
		data1 = window['AscCommon'].g_specialPasteHelper.specialPasteData.data1;
		data2 = window['AscCommon'].g_specialPasteHelper.specialPasteData.data2;
		text_data = window['AscCommon'].g_specialPasteHelper.specialPasteData.text_data;

		if(!(specialPasteProps === Asc.c_oSpecialPasteProps.keepTextOnly && _format !== AscCommon.c_oAscClipboardDataFormat.Text && text_data))
		{
			text_data = null;
		}
	}

	switch (_format)
	{
		case AscCommon.c_oAscClipboardDataFormat.HtmlElement:
		{
			oPasteProcessor.Start(data1, data2, null, null, text_data);
			break;
		}
		case AscCommon.c_oAscClipboardDataFormat.Internal:
		{
			oPasteProcessor.Start(null, null, null, data1, text_data);
			break;
		}
		case AscCommon.c_oAscClipboardDataFormat.Text:
		{
			oPasteProcessor.Start(null, null, null, null, data1);
			break;
		}
	}
}
function trimString( str ){
    return str.replace(/^\s+|\s+$/g, '') ;
}
function sendImgUrls(api, images, callback, bNotShowError, token) {
  if (window["NATIVE_EDITOR_ENJINE"] === true && window["IS_NATIVE_EDITOR"] !== true)
  {
    var _data = [];
    for (var i = 0; i < images.length; i++)
    {
      var _url = window["native"]["getImageUrl"](images[i]);
      var _full_path = window["native"]["getImagesDirectory"]() + "/" + _url;
      var _local_url = "media/" + _url;
      AscCommon.g_oDocumentUrls.addUrls({_local_url:_full_path});
      _data[i] = {url:_full_path, path:_local_url};
    }
    callback(_data);
    return;
  }
  if (window["AscDesktopEditor"])
  {
    // correct local images
    for (var nIndex = images.length - 1; nIndex >= 0; nIndex--)
    {
      if (0 == images[nIndex].indexOf("file:/"))
        images[nIndex] = window["AscDesktopEditor"]["GetImageBase64"](images[nIndex]);
    }
  }

  if (AscCommon.EncryptionWorker && AscCommon.EncryptionWorker.isCryptoImages())
  {
      return AscCommon.EncryptionWorker.addCryproImagesFromUrls(images, callback);
  }

  if(window["IS_NATIVE_EDITOR"])
  {
	callback([]);
	return;
  }

  var rData = {
    "id": api.documentId, "c": "imgurls", "userid": api.documentUserId, "saveindex": g_oDocumentUrls.getMaxIndex(),
    "tokenDownload": token, "data": images
  };
  if (!api.isOpenedChartFrame) {
      api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);
  }

  api.fCurCallback = function (input) {
    if (!api.isOpenedChartFrame) {
        api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);
    }
    var nError = c_oAscError.ID.No;
    var data;
    if (null != input && "imgurls" == input["type"]) {
      if ("ok" == input["status"]) {
        data = input["data"]["urls"];
        nError = AscCommon.mapAscServerErrorToAscError(input["data"]["error"]);
        var urls = {};
        for (var i = 0, length = data.length; i < length; ++i) {
          var elem = data[i];
          if (null != elem.url) {
            urls[elem.path] = elem.url;
          }
        }
        g_oDocumentUrls.addUrls(urls);
      } else {
        nError = AscCommon.mapAscServerErrorToAscError(parseInt(input["data"]));
      }
    } else {
      nError = c_oAscError.ID.Unknown;
    }
    if ( c_oAscError.ID.No !== nError && !bNotShowError) {
        api.sendEvent("asc_onError", nError, c_oAscError.Level.NoCritical);
    }
    if (!data) {
      //todo сделать функцию очистки, чтобы можно было оборвать paste и показать error
      data = [];
      for ( var i = 0; i < images.length; ++i) {
        data.push({'url': 'error', 'path': 'error'});
      }
    }
    callback(data);
  };
  if (api.isEditOleMode) {
    const sendInformation = {
      "type": AscCommon.c_oAscFrameDataType.SendImageUrls,
      "information": {
          "images": images,
          "bNotShowError": bNotShowError,
          "token": token
        }
    }
    api.sendFromFrameToGeneralEditor(sendInformation);
    return;
    }
  AscCommon.sendCommand(api, null, rData);
}
function PasteProcessor(api, bUploadImage, bUploadFonts, bNested, pasteInExcel, pasteCallback)
{
    this.oRootNode = null;
    this.api = api;
    this.bIsDoublePx  = api.WordControl.bIsDoublePx;
    this.oDocument = api.WordControl.m_oLogicDocument;
    this.oLogicDocument = this.oDocument;
    this.oRecalcDocument = this.oDocument;
    this.map_font_index = api.FontLoader.map_font_index;
    this.bUploadImage = bUploadImage;
    this.bUploadFonts = bUploadFonts;
    this.bNested = bNested;//для параграфов в таблицах
    this.oFonts = {};
    this.oImages = {};
	this.aContent = [];
	this.AddedFootEndNotes = {};
	this.aContentForNotes = [];

	this.bIsForFootEndnote = false;
	this.pasteInExcel = pasteInExcel;
	this.pasteInPresentationShape = null;
	this.pasteCallback = pasteCallback;

	this.maxTableCell = null;

	//для вставки текста в ячейку, при копировании из word в chrome появляются лишние пробелы вне <p>
    this.bIgnoreNoBlockText = false;

    this.oCurRun = null;
    this.oCurRunContentPos = 0;
    this.oCurPar = null;
    this.oCurParContentPos = 0;
    this.oCurHyperlink = null;
    this.oCurHyperlinkContentPos = 0;
    this.oCur_rPr = new CTextPr();

	//Br копятся потомы что есть случаи когда не надо вывобить br, хотя он и присутствует.
    this.nBrCount = 0;
	//bInBlock указывает блочный ли элемент(рассматриваются только элементы дочерние от child)
	//Если после окончания вставки true != this.bInBlock значит последний элемент не параграф и не надо добавлять новый параграф
    this.bInBlock = null;

	//ширина элемента в который вставляем страница или ячейка
    this.dMaxWidth = Page_Width - X_Left_Margin - X_Right_Margin;
	//коэфициент сжатия(например при вставке таблица сжалась, значит при вставке содержимого ячейки к картинкам и таблице будет применен этот коэффициент)
    this.dScaleKoef = 1;
    this.bUseScaleKoef = false;
	this.bIsPlainText = false;

	this.defaultImgWidth = 50;
	this.defaultImgHeight = 50;

    this.MsoStyles = {"mso-style-type": 1, "mso-pagination": 1, "mso-line-height-rule": 1, "mso-style-textfill-fill-color": 1, "mso-tab-count": 1,
        "tab-stops": 1, "list-style-type": 1, "mso-special-character": 1, "mso-column-break-before": 1, "mso-break-type": 1, "mso-padding-alt": 1, "mso-border-insidev": 1,
        "mso-border-insideh": 1, "mso-row-margin-left": 1, "mso-row-margin-right": 1, "mso-cellspacing": 1, "mso-border-alt": 1,
        "mso-border-left-alt": 1, "mso-border-top-alt": 1, "mso-border-right-alt": 1, "mso-border-bottom-alt": 1, "mso-border-between": 1, "mso-list": 1,
		"mso-comment-reference": 1, "mso-comment-date": 1, "mso-comment-continuation": 1, "mso-data-placement": 1};
    this.oBorderCache = {};

	this.msoListMap = [];

	//пока ввожу эти параметры для специальной вставки. возможно, нужно будет пересмотреть и убрать их
	this.pasteTypeContent = undefined;
	this.pasteList = undefined;
	this.pasteIntoElem = undefined;//ссылка на элемент контента, который был выделен до вставки

	this.apiEditor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window["editor"];

	this.msoComments = [];

	this.startMsoAnnotation = undefined;
	this.needAddCommentStart = null;
	this.needAddCommentEnd = null;

	this.aMsoHeadStylesStr = null;
	this.oMsoHeadStylesListMap = [];
	this.oMsoStylesParser = null;

	this.pasteTextIntoList = null;

	this.rtfImages = null;

	this.aNeedRecalcImgSize = null;

	this.doNotInsertInDoc = null;
}
PasteProcessor.prototype =
{
    _GetTargetDocument : function(oDocument)
    {
        if(PasteElementsId.g_bIsDocumentCopyPaste)
        {
			var nDocPosType = oDocument.GetDocPosType();
			if (docpostype_HdrFtr === nDocPosType)
			{
				if (null != oDocument.HdrFtr && null != oDocument.HdrFtr.CurHdrFtr && null != oDocument.HdrFtr.CurHdrFtr.Content)
				{
					oDocument  = oDocument.HdrFtr.CurHdrFtr.Content;
					this.oRecalcDocument = oDocument;
				}
			}
			else if (nDocPosType === docpostype_DrawingObjects)
			{
				var content = oDocument.DrawingObjects.getTargetDocContent(true);
				if (content)
				{
					oDocument = content;
				}
			}
			else if (nDocPosType === docpostype_Footnotes)
			{
				if (oDocument.Footnotes && oDocument.Footnotes.CurFootnote)
					oDocument = oDocument.Footnotes.CurFootnote
			}
			else if (nDocPosType === docpostype_Endnotes)
			{
				if (oDocument.Endnotes && oDocument.Endnotes.CurEndnote)
					oDocument = oDocument.Endnotes.CurEndnote
			}

			// Отдельно обрабатываем случай, когда курсор находится внутри таблицы
			var Item = oDocument.Content[oDocument.CurPos.ContentPos];
			if (type_Table === Item.GetType() && null != Item.CurCell)
			{
				this.dMaxWidth = this._CalcMaxWidthByCell(Item.CurCell);
				oDocument = this._GetTargetDocument(Item.CurCell.Content);
			}
        }
        else
        {

        }
        return oDocument;
    },
    _CalcMaxWidthByCell : function(cell)
    {
        var row = cell.Row;
        var table = row.Table;
        var grid = table.TableGrid;
        var nGridBefore = 0;
        if(null != row.Pr && null != row.Pr.GridBefore)
            nGridBefore = row.Pr.GridBefore;
        var nCellIndex = cell.Index;
        var nCellGrid = 1;
        if(null != cell.Pr && null != cell.Pr.GridSpan)
            nCellGrid = cell.Pr.GridSpan;
        var nMarginLeft = 0;
        if(null != cell.Pr && null != cell.Pr.TableCellMar && null != cell.Pr.TableCellMar.Left && tblwidth_Mm === cell.Pr.TableCellMar.Left.Type && null != cell.Pr.TableCellMar.Left.W)
            nMarginLeft = cell.Pr.TableCellMar.Left.W;
        else if(null != table.Pr && null != table.Pr.TableCellMar && null != table.Pr.TableCellMar.Left && tblwidth_Mm === table.Pr.TableCellMar.Left.Type && null != table.Pr.TableCellMar.Left.W)
            nMarginLeft = table.Pr.TableCellMar.Left.W;
        var nMarginRight = 0;
        if(null != cell.Pr && null != cell.Pr.TableCellMar && null != cell.Pr.TableCellMar.Right && tblwidth_Mm === cell.Pr.TableCellMar.Right.Type && null != cell.Pr.TableCellMar.Right.W)
            nMarginRight = cell.Pr.TableCellMar.Right.W;
        else if(null != table.Pr && null != table.Pr.TableCellMar && null != table.Pr.TableCellMar.Right && tblwidth_Mm === table.Pr.TableCellMar.Right.Type && null != table.Pr.TableCellMar.Right.W)
            nMarginRight = table.Pr.TableCellMar.Right.W;
        var nPrevSumGrid = nGridBefore;
        for(var i = 0; i < nCellIndex; ++i)
        {
            var oTmpCell = row.Content[i];
            var nGridSpan = 1;
            if(null != cell.Pr && null != cell.Pr.GridSpan)
                nGridSpan = cell.Pr.GridSpan;
            nPrevSumGrid += nGridSpan;
        }
        var dCellWidth = 0;
        for(var i = nPrevSumGrid, length = grid.length; i < nPrevSumGrid + nCellGrid && i < length; ++i)
            dCellWidth += grid[i];

        if(dCellWidth - nMarginLeft - nMarginRight <= 0)
            dCellWidth = 4;
        else
            dCellWidth -= nMarginLeft + nMarginRight;
        return dCellWidth;
    },
    InsertInDocument : function(dNotShowOptions)
    {
        var oDocument = this.oDocument;

		//TODO ориентируюсь при специальной вставке на SelectionState. возможно стоит пересмотреть.
		this._initSelectedElem();

        var nInsertLength = this.aContent.length;
        if(nInsertLength > 0)
        {
			this.InsertInPlace(oDocument, this.aContent);

            if(false === PasteElementsId.g_bIsDocumentCopyPaste)
            {
                oDocument.Recalculate();
                if(oDocument.Parent != null && oDocument.Parent.txBody != null)
                {
                    oDocument.Parent.txBody.recalculate();
                }
            }
        }

        var bNeedRecalculate = (false === this.bNested && nInsertLength > 0);

		//for special paste
		if(dNotShowOptions && !window['AscCommon'].g_specialPasteHelper.specialPasteStart) {
			window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
		} else {
			this._specialPasteSetShowOptions();
		}

		window['AscCommon'].g_specialPasteHelper.Paste_Process_End(true);
        if(bNeedRecalculate)
        {
            this.oRecalcDocument.Recalculate();
            this.oLogicDocument.Document_UpdateInterfaceState();
            this.oLogicDocument.Document_UpdateSelectionState();
        }
    },
    InsertInPlace : function(oDoc, aNewContent)
    {
		if (!PasteElementsId.g_bIsDocumentCopyPaste)
			return;

		var specialPasteHelper = window['AscCommon'].g_specialPasteHelper;
		var bIsSpecialPaste = specialPasteHelper.specialPasteStart;
		var paragraph = oDoc.GetCurrentParagraph();
		var oTable = oDoc.IsInTable() ? oDoc.GetParent().GetTable() : null;

		//pasteTypeContent - если все содержимое одного типа
		//TODO пересмотреть pasteTypeContent
		this.pasteTypeContent = null;
		var oSelectedContent = new AscCommonWord.CSelectedContent();

		//by section
		//oSelectedContent.LastSection = this.getLastSectionPr();

		var tableSpecialPaste = false;
		let pasteHelperElement = null;

		if (oTable && !aNewContent[0].IsTable() && oTable.IsCellSelection())
		{
			var arrSelectedCells = oTable.GetSelectionArray(true);
			var nPrevRow      = -1;
			var nElementIndex = -1;
			var nRowIndex     = -1;
			var arrNewContent = [];
			for (var i = aNewContent.length - 1; i >= 0; i--)
			{
				arrNewContent.unshift([]);
				var oElem = aNewContent[i].Copy();
				if (oElem.IsParagraph())
				{
					for (var j = oElem.Content.length - 1; j >= 0; j--)
					{
						var oNestedElem = oElem.Content[j];
						if (oNestedElem.Type === para_Run && oNestedElem.Content.length)
						{
							for (var k = oNestedElem.Content.length - 1; k >= 0; k--)
							{
								var oThirdElem = oNestedElem.Content[k];
								if (oThirdElem.Type === para_Tab)
								{
									var oPar = new Paragraph(oDoc.DrawingDocument);
									arrNewContent[0].unshift(oPar);
									var oRun = oNestedElem.Split2(k);
									oRun.RemoveFromContent(0, 1);
									oPar.AddToContent(0, oRun);
								}
							}
						}	
					}
				}
				arrNewContent[0].unshift(oElem);
			}
			for (var nIndex = 0, nCount = arrSelectedCells.length; nIndex < nCount; ++nIndex)
			{
				var oPos  = arrSelectedCells[nIndex];
				var oRow  = oTable.GetRow(oPos.Row);

				if (!oRow)
					continue;

				var oCell = oRow.GetCell(oPos.Cell);
				if (!oCell)
					continue;

				var isMerged = oCell.GetVMerge() === vmerge_Continue;
				if (isMerged)
					oCell = oTable.GetStartMergedCell(oPos.Cell, oPos.Row);

				var oCellContent = oCell.GetContent();
				var oPara;
				if (isMerged) {
					oPara = new Paragraph(oDoc.DrawingDocument);
					oCellContent.AddToContent(oCellContent.Content.length, oPara, true);
					oPara.Document_SetThisElementCurrent(false);
				}
				else
				{
					oCellContent.ClearContent(true);
					oPara = oCellContent.GetElement(0);
				}

				if (!oPara || !oPara.IsParagraph())
					continue;

				if (oPos.Row !== nPrevRow)
				{
					nPrevRow = oPos.Row;
					nRowIndex++;
					nElementIndex = -1;

					if (nRowIndex > arrNewContent.length - 1)
						nRowIndex = 0;
				}

				nElementIndex++;

				if (nElementIndex > arrNewContent[nRowIndex].length - 1)
					nElementIndex = 0;

				if (nRowIndex < 0 || nElementIndex < 0)
					break;

				oSelectedContent.Reset();

				var NewElem = arrNewContent[nRowIndex][nElementIndex].Copy();

				var NearPos = oPara.GetCurrentAnchorPosition();
				if (bIsSpecialPaste)
				{
					if (Asc.c_oSpecialPasteProps.insertAsNestedTable === specialPasteHelper.specialPasteProps ||
						Asc.c_oSpecialPasteProps.overwriteCells === specialPasteHelper.specialPasteProps)
					{
						tableSpecialPaste = true;
						oSelectedContent.SetInsertOptionForTable(specialPasteHelper.specialPasteProps);
					}
				}
				else
				{
					oSelectedContent.SetInsertOptionForTable(Asc.c_oSpecialPasteProps.insertAsNestedTable);
				}


				if (bIsSpecialPaste && !tableSpecialPaste)
				{
					var parseItem = this._specialPasteItemConvert(NewElem);
					if (parseItem && parseItem.length)
					{
						for (var j = 0; j < parseItem.length; j++)
						{
							if (j === 0)
							{
								arrNewContent.splice(nRowIndex + j, 1, parseItem[j]);
							}
							else
							{
								arrNewContent.splice(nRowIndex + j, 0, parseItem[j]);
							}
						}
					}
				}


				var oSelectedElement     = new AscCommonWord.CSelectedElement();
				oSelectedElement.Element = NewElem;

				var type = this._specialPasteGetElemType(NewElem);
				if (0 === nIndex)
				{
					this.pasteTypeContent = type;
				}
				else if (type !== this.pasteTypeContent)
				{
					this.pasteTypeContent = null;
				}

				oSelectedElement.SelectedAll = false;
				oSelectedContent.Add(oSelectedElement);

				oSelectedContent.EndCollect(this.oLogicDocument);
				oSelectedContent.SetCopyComments(false);

				if (!this.pasteInExcel && !oSelectedContent.CanInsert(NearPos))
				{
					this.oLogicDocument.Document_Undo();
					History.Clear_Redo();
					return;
				}

				oSelectedContent.Insert(NearPos, false);
			}
		}
		else
		{
			oTable = null;
			var paragraph = oDoc.GetCurrentParagraph();
			var NearPos = paragraph.GetCurrentAnchorPosition();
			//делаем небольшой сдвиг по y, потому что сама точка TargetPos для двухстрочного параграфа определяется как верхняя
			//var NearPos = oDoc.Get_NearestPos(this.oLogicDocument.TargetPos.PageNum, this.oLogicDocument.TargetPos.X, this.oLogicDocument.TargetPos.Y + 0.05);//0.05 == 2pix

			if(bIsSpecialPaste){
				if (Asc.c_oSpecialPasteProps.insertAsNestedTable === specialPasteHelper.specialPasteProps ||
					Asc.c_oSpecialPasteProps.overwriteCells === specialPasteHelper.specialPasteProps)
				{
					tableSpecialPaste = true;
					oSelectedContent.SetInsertOptionForTable(specialPasteHelper.specialPasteProps);
				}
			}
			for (var i = 0; i < aNewContent.length; ++i) {
				if(bIsSpecialPaste && !tableSpecialPaste)
				{
					var parseItem = this._specialPasteItemConvert(aNewContent[i]);
					if(parseItem && parseItem.length)
					{
						for(var j = 0; j < parseItem.length; j++)
						{
							if(j === 0)
							{
								aNewContent.splice(i + j, 1, parseItem[j]);
							}
							else
							{
								aNewContent.splice(i + j, 0, parseItem[j]);
							}
						}
					}
				}

				var oSelectedElement = new AscCommonWord.CSelectedElement();
				oSelectedElement.Element = aNewContent[i];

				var type = this._specialPasteGetElemType(aNewContent[i]);
				if(0 === i)
				{
					this.pasteTypeContent = type;
				}
				else if(type !== this.pasteTypeContent)
				{
					this.pasteTypeContent = null;
				}

				if (i === aNewContent.length - 1 && true != this.bInBlock && type_Paragraph === oSelectedElement.Element.GetType())
					oSelectedElement.SelectedAll = false;
				else
					oSelectedElement.SelectedAll = true;
				oSelectedContent.Add(oSelectedElement);
			}

			//проверка на возможность втавки в формулу
			//TODO проверку на excel пеерсмотреть!!!!
			oSelectedContent.EndCollect(this.oLogicDocument);
			oSelectedContent.SetCopyComments(false);

			if (this.doNotInsertInDoc) {
				this.pasteCallback && this.pasteCallback(oSelectedContent);
				return;
			}

			if(!this.pasteInExcel && !oSelectedContent.CanInsert(NearPos))
			{
				this.oLogicDocument.Document_Undo();
				History.Clear_Redo();
				return;
			}

			oSelectedContent.Insert(NearPos, false);
			paragraph.Clear_NearestPosArray(aNewContent);
			pasteHelperElement = oSelectedContent.GetPasteHelperElement();
		}

		//если вставляем таблицу в ячейку таблицы
		if (this.pasteIntoElem && 1 === this.aContent.length && type_Table === this.aContent[0].GetType() &&
			this.pasteIntoElem.Parent && this.pasteIntoElem.Parent.IsInTable() && (!bIsSpecialPaste || (bIsSpecialPaste &&
			Asc.c_oSpecialPasteProps.overwriteCells === specialPasteHelper.specialPasteProps))) {
			//TODO пересмотреть положение кнопки специальной вставки при вставке в таблицу
			var table;
			var tableCell = paragraph && paragraph.Parent && paragraph.Parent.Parent;
			if (tableCell && tableCell.GetTable) {
				table = tableCell.GetTable()
			} else {
				table = this.pasteIntoElem.Parent.Parent.Get_Table();
			}
			specialPasteHelper.showButtonIdParagraph = table.Id;
		} else {
			if (pasteHelperElement)
			{
				specialPasteHelper.showButtonIdParagraph = pasteHelperElement.GetId();
			}
			else if (oSelectedContent.Elements.length === 1 && !oTable)
			{
				let currentParagraph = this.oDocument.GetCurrentParagraph();
				if (currentParagraph)
					specialPasteHelper.showButtonIdParagraph = currentParagraph.GetId();
			}
			else
			{
				specialPasteHelper.showButtonIdParagraph = oSelectedContent.Elements[oSelectedContent.Elements.length - 1].Element.Id;
			}
		}


		if(this.oLogicDocument && this.oLogicDocument.DrawingObjects)
		{
			var oTargetTextObject = AscFormat.getTargetTextObject(this.oLogicDocument.DrawingObjects);
			oTargetTextObject && oTargetTextObject.checkExtentsByDocContent && oTargetTextObject.checkExtentsByDocContent();
		}

		this._selectShapesBeforeInsert(aNewContent, oDoc);

    },

	//***functions for special paste***
	_specialPasteGetElemType: function(elem)
	{
		var type = elem.GetType();

		if(type_Paragraph === type)
		{
			//проверяем, возможно это графический объект
			for(var i = 0; i < elem.Content.length; i++)
			{
				if(elem.Content[i] && elem.Content[i].Content)
				{
					for(var j = 0; j < elem.Content[i].Content.length; j++)
					{
						var contentElem = elem.Content[i].Content[j];
						if(!(contentElem instanceof AscWord.CRunParagraphMark))
						{
							var typeElem = contentElem.GetType ? contentElem.GetType() : null;
							if(para_Drawing === typeElem)
							{
								type = para_Drawing;
							}
							else
							{
								if(para_Drawing !== type)
								{
									type = type_Paragraph;
								}
								else
								{
									type = null;
								}

								break;
							}
						}
					}
				}

				if(type_Paragraph === type)
				{
					if(elem.Pr && elem.Pr.NumPr)
					{
						if(undefined === this.pasteList)
						{
							this.pasteList = elem.Pr.NumPr;
						}
						else if(this.pasteList && !elem.Pr.NumPr.Is_Equal(this.pasteList))
						{
							this.pasteList = null;
						}
					}
				}
			}
		}

		return type;
	},

	_specialPasteSetShowOptions: function()
	{
		//специальная вставка:
		//выдаем стандартные параметры всавки(paste, merge, value) во всех ситуация, за исключением:
		//если вставляем единственную таблицу в таблицу - особые параметры вставки(как извне, так и внутри)
		//если вставляем список - должны совпадать типы с уже существующими(как извне, так и внутри)
		//изображения / шейпы

		//отдельно диаграммы - для них есть отдельный пункт. посмотреть, нужно ли это добавлять

		//для формул параметры как и при обычной вставке. но нужно уметь их преобразовывать в текст при вставке только текста
		//особые параметры при вставке таблиц из EXCEL


		//если вставляются только изображения, пока не показываем параметры специальной
		if(para_Drawing === this.pasteTypeContent)
		{
			window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
			if(window['AscCommon'].g_specialPasteHelper.buttonInfo)
			{
				window['AscCommon'].g_specialPasteHelper.showButtonIdParagraph = null;
				window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
			}
			return;
		}

		var specialPasteShowOptions = !window['AscCommon'].g_specialPasteHelper.buttonInfo.isClean() ? window['AscCommon'].g_specialPasteHelper.buttonInfo : null;
		if(!window['AscCommon'].g_specialPasteHelper.specialPasteStart)
		{
			specialPasteShowOptions = window['AscCommon'].g_specialPasteHelper.buttonInfo;

			var sProps = Asc.c_oSpecialPasteProps;
			var aContent = this.aContent;

			var props = null;
			//table into table
			//this.pasteTypeContent и this.pasteList нужны для вставки таблиц/списков и тд
			//TODO пока вставка будет работать только с текстом(форматированный/не форматированный)
			/*if(insertToElem && 1 === aContent.length && type_Table === this.aContent[0].GetType() && type_Table === insertToElem.GetType())
			{
				props = [sProps.paste, sProps.insertAsNestedTable, sProps.uniteIntoTable, sProps.insertAsNewRows, sProps.pasteOnlyValues];
			}
			else if(this.pasteList && insertToElem && type_Paragraph === insertToElem.GetType() && insertToElem.Pr && insertToElem.Pr.NumPr && insertToElem.Pr.NumPr.Is_Equal(this.pasteList))
			{
				//вставка нумерованного списка в нумерованный список
				props = [sProps.paste, sProps.uniteList, sProps.doNotUniteList];
			}*/

			//если вставляем одну таблицу в ячейку другой таблицы
			if (this.pasteIntoElem && 1 === aContent.length && type_Table === this.aContent[0].GetType() &&
				this.pasteIntoElem.Parent && this.pasteIntoElem.Parent.IsInTable())
			{
				props = [sProps.overwriteCells, sProps.insertAsNestedTable, sProps.keepTextOnly];
			}
			else
			{
				props = [sProps.sourceformatting/*, sProps.mergeFormatting*/, sProps.keepTextOnly];
			}

			if(null !== props)
			{
				specialPasteShowOptions.asc_setOptions(props);
			}
			else
			{
				window['AscCommon'].g_specialPasteHelper.showButtonIdParagraph = null;
				window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
			}
		}

		if(specialPasteShowOptions)
		{
			//SpecialPasteButtonById_Show вызываю здесь, если пересчет документа завершился раньше, чем мы попали сюда и сгенерировали параметры вставки
			//в противном случае вызываю SpecialPasteButtonById_Show в drawingDocument->OnEndRecalculate
			if (window['AscCommon'].g_specialPasteHelper.endRecalcDocument) {
				window['AscCommon'].g_specialPasteHelper.SpecialPasteButtonById_Show();
			}
		}
	},

	_specialPasteItemConvert: function(item)
	{
		//TODO рассмотреть вариант вставки текста ("text/plain")
		//для вставки простого текста, можно было бы использовать ("text/plain")
		//но в данном случае вставка текста будет работать не совсем корретно внутри приложения, поскольку
		//когда мы пишем в буфер текст, функция GetSelectedText отдаёт вместо табуляции пробелы
		//так же некорректно будут вставляться таблицы, поскольку табуляции между ячейками мы потеряем
		//внутренние таблицы мы вообще теряем
		//для реализации необходимо менять функцию GetSelectedText
		//посмотреть, какие браузер могут заменить табуляцию на пробел при занесении текста в буфер обмена

		var res = item;
		var type = item.GetType();
		switch(type)
		{
			case type_Paragraph:
			{
				res = this._specialPasteParagraphConvert(item);
				break;
			}
			case type_Table:
			{
				res = this._specialPasteTableConvert(item);
				break;
			}
		}
		return res;
	},

	_specialPasteTableConvert: function(table)
	{
		//TODO временная функция
		var res = table;

		var props = window['AscCommon'].g_specialPasteHelper.specialPasteProps;
		if(props === Asc.c_oSpecialPasteProps.keepTextOnly)
		{
			res = this._convertTableToText(table);
		}
		else
		{
			for(var i = 0; i < table.Content.length; i++)
			{
				for(var j = 0; j < table.Content[i].Content.length; j++)
				{
					var cDocumentContent = table.Content[i].Content[j].Content;
					for(var n = 0; n < cDocumentContent.Content.length; n++)
					{
						if(cDocumentContent.Content[n] instanceof Paragraph)
						{
							this._specialPasteParagraphConvert(cDocumentContent.Content[n]);
						}
						else if(cDocumentContent.Content[n] instanceof CTable)
						{
							this._specialPasteTableConvert(cDocumentContent.Content[n]);
						}
					}
				}
			}
		}

		return res;
	},

	_specialPasteParagraphConvert: function(paragraph)
	{
		var res = paragraph;
		var props = window['AscCommon'].g_specialPasteHelper.specialPasteProps;

		//стиль текущего параграфа/рана, в который вставляем
		var pasteIntoParagraphPr = this.oDocument.GetDirectParaPr();
		var pasteIntoParaRunPr = this.oDocument.GetDirectTextPr();

		switch(props)
		{
			case Asc.c_oSpecialPasteProps.paste:
			{
				break;
			}
			case Asc.c_oSpecialPasteProps.keepTextOnly:
			{
				//TODO check it and remove/modify this code
				var numbering =  paragraph.GetNumPr();
				if(numbering)
				{
					//проставляем параграфам NumInfo
					var parentContent = paragraph.Parent instanceof CDocument ? this.aContent : paragraph.Parent.Content;
					for(var i = 0; i < parentContent.length; i++)
					{
						var tempParagraph = parentContent[i];
						var numbering2 =  tempParagraph.GetNumPr ? tempParagraph.GetNumPr() : null;

						if(numbering2)
						{
							var NumberingEngine = new CDocumentNumberingInfoEngine(tempParagraph.Id, numbering2, this.oLogicDocument.Get_Numbering());
							var numInfo2 = tempParagraph.Numbering.Internal.NumInfo;

							if(!numInfo2 || (numInfo2 && !numInfo2[numbering.Lvl]))
							{
								for (var nIndex = 0, nCount = parentContent.length; nIndex < nCount; ++nIndex)
								{
									parentContent[nIndex].GetNumberingInfo(NumberingEngine);
								}

								if (NumberingEngine.NumInfo.length && NumberingEngine.NumInfo[0] !== undefined) {
									tempParagraph.Numbering.Internal.NumInfo = NumberingEngine.NumInfo;
								}
							}
						}

					}

					this._checkNumberingText(paragraph, paragraph.Numbering.Internal.NumInfo, numbering);
				}


				if(pasteIntoParagraphPr)
				{
					paragraph.Set_Pr(pasteIntoParagraphPr.Copy());

					if(paragraph.TextPr && pasteIntoParaRunPr)
					{
						paragraph.TextPr.Value = pasteIntoParaRunPr.Copy();
					}
				}
				this._specialPasteParagraphContentConvert(paragraph.Content, pasteIntoParaRunPr);

				break;
			}
			case Asc.c_oSpecialPasteProps.mergeFormatting:
			{
				//ms почему-то при merge игнорирует заливку текста
				if(pasteIntoParagraphPr)
				{
					paragraph.Pr.Merge(pasteIntoParagraphPr);
					if(paragraph.TextPr)
					{
						paragraph.TextPr.Value.Merge(pasteIntoParaRunPr);
					}
				}
				this._specialPasteParagraphContentConvert(paragraph.Content, pasteIntoParaRunPr);

				break;
			}
		}

		return res;
	},

	_specialPasteParagraphContentConvert: function(paragraphContent, pasteIntoParaRunPr)
	{
		var props = window['AscCommon'].g_specialPasteHelper.specialPasteProps;

		var checkInsideDrawings = function(runContent)
		{
			for(var j = 0; j < runContent.length; j++)
			{
				var item = runContent[j];

				switch(item.Type)
				{
					case para_Run:
					{
						checkInsideDrawings(item.Content);
						break;
					}
					case para_Drawing:
					{
						runContent.splice(j, 1);
						break;
					}
				}
			}
		};

		switch(props)
		{
			case Asc.c_oSpecialPasteProps.paste:
			{
				break;
			}
			case Asc.c_oSpecialPasteProps.keepTextOnly:
			{
				//в данному случае мы должны применить к вставленному фрагменту стиль paraRun, в который вставляем
				if(pasteIntoParaRunPr)
				{
					for(var i = 0; i < paragraphContent.length; i++)
					{
						var elem = paragraphContent[i];
						var type = elem.Type;
						switch(type)
						{
							case para_Run:
							{
								//проверить, есть ли внутри изображение
								if(pasteIntoParaRunPr && elem.Set_Pr)
								{
									elem.Set_Pr( pasteIntoParaRunPr.Copy() );
								}

								checkInsideDrawings(elem.Content);

								break;
							}
							case para_Field:
							case para_InlineLevelSdt:
							case para_Hyperlink:
							{
								//изменить hyperlink на pararun
								//проверить, есть ли внутри изображение

								paragraphContent.splice(i, 1);
								for(var n = 0; n < elem.Content.length; n++)
								{
									paragraphContent.splice(i + n, 0, elem.Content[n]);
								}
								i--;

								break;
							}
							case para_Math:
							{
								//преобразуем в текст
								var mathToParaRun = this._convertParaMathToText(elem);
								if(mathToParaRun)
								{
									paragraphContent.splice(i, 1, mathToParaRun);
									i--;
								}

								break;
							}
							case para_Comment:
							{
								//TODO в дальнейшем лучше удалять коммент а не заменять его
								paragraphContent.splice(i, 1, new ParaRun());
								i--;

								break;
							}
						}
					}
				}

				break;
			}
			case Asc.c_oSpecialPasteProps.mergeFormatting:
			{
				//ms почему-то при merge игнорирует заливку текста
				if(pasteIntoParaRunPr)
				{
					for(var i = 0; i < paragraphContent.length; i++)
					{
						var elem = paragraphContent[i];
						if(pasteIntoParaRunPr && elem.Pr)
						{
							elem.Pr.Merge(pasteIntoParaRunPr);
						}
					}
				}

				break;
			}
		}
	},

	_convertParaMathToText: function(paraMath)
	{
		var res = null;
		var oDoc = this.oLogicDocument;

		var mathText = paraMath.Root.GetTextContent();
		if(mathText && mathText.str)
		{
			var newParaRun = new ParaRun();
			addTextIntoRun(newParaRun, mathText.str);

			res = newParaRun;
		}

		return res;
	},

	_convertTableToText: function(table, obj, newParagraph)
	{
		var oDoc = this.oLogicDocument;
		var t = this;
		if(!obj)
		{
			obj = [];
		}

		//row
		for(var i = 0; i < table.Content.length; i++)
		{
			if(!newParagraph)
			{
				newParagraph = new Paragraph(oDoc.DrawingDocument, oDoc);
			}

			//col
			for(var j = 0; j < table.Content[i].Content.length; j++)
			{
				//content
				var cDocumentContent = table.Content[i].Content[j].Content;

				var createNewParagraph = false;
				var previousTableAdd = false;
				for(var n = 0; n < cDocumentContent.Content.length; n++)
				{
					previousTableAdd = false;
					if(createNewParagraph)
					{
						newParagraph = new Paragraph(oDoc.DrawingDocument, oDoc);
						createNewParagraph = false;
					}

					if(cDocumentContent.Content[n] instanceof Paragraph)
					{
						//TODO пересмотреть обработку. получаем текст из контента, затем делаем контент из текста!
						this._specialPasteParagraphConvert(cDocumentContent.Content[n]);

						var value = cDocumentContent.Content[n].GetText();
						var newParaRun = new ParaRun();

						var bIsAddTabBefore = false;
						if(newParagraph.Content.length > 1)
						{
							bIsAddTabBefore = true;
						}

						addTextIntoRun(newParaRun, value, bIsAddTabBefore, true);

						newParagraph.Internal_Content_Add(newParagraph.Content.length - 1, newParaRun, false);
					}
					else if(cDocumentContent.Content[n] instanceof CTable)
					{
						t._convertTableToText(cDocumentContent.Content[n], obj, newParagraph);
						createNewParagraph = true;
						previousTableAdd = true;
					}

					if(!previousTableAdd && cDocumentContent.Content.length > 1 && n !== cDocumentContent.Content.length - 1)
					{
						obj.push(newParagraph);
						createNewParagraph = true;
					}
				}
			}

			obj.push(newParagraph);
			newParagraph = null;
		}

		return obj;
	},

	_checkNumberingText: function (paragraph, oNumInfo, oNumPr) {
		if (oNumPr && oNumInfo) {
			var oNum = this.oLogicDocument.GetNumbering().GetNum(oNumPr.NumId);
			if (oNum) {
				var sNumberingText = oNum.GetText(oNumPr.Lvl, oNumInfo);

				var newParaRun = new ParaRun();
				addTextIntoRun(newParaRun, sNumberingText, false, true, true);
				paragraph.Internal_Content_Add(0, newParaRun, false);
			}
		}
	},

	_initSelectedElem: function () {
		this.curDocSelection = this.oDocument.GetSelectionState();
		if (this.curDocSelection && this.curDocSelection[1] && this.curDocSelection[1].CurPos) {
			this.pasteIntoElem = this.oDocument.Content[this.curDocSelection[1].CurPos.ContentPos];
		}
	},

	//***end special paste***

	InsertInPlacePresentation: function(aNewContent, isText)
	{
		var presentation = editor.WordControl.m_oLogicDocument;

		var presentationSelectedContent = new PresentationSelectedContent();
		presentationSelectedContent.DocContent = new AscCommonWord.CSelectedContent();
		for (var i = 0, length = aNewContent.length; i < length; ++i) {
			var oSelectedElement = new AscCommonWord.CSelectedElement();

			if (window['AscCommon'].g_specialPasteHelper.specialPasteStart && !isText) {
				aNewContent[i]= this._specialPasteItemConvert(aNewContent[i]);
			}

			oSelectedElement.Element = aNewContent[i];
			presentationSelectedContent.DocContent.Elements[i] = oSelectedElement;
		}

		if(presentation.InsertContent(presentationSelectedContent)) {
			presentation.Recalculate();
            editor.checkChangesSize();
			presentation.Document_UpdateInterfaceState();

			this._setSpecialPasteShowOptionsPresentation();
		} else {
			//window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
		}

		window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
	},

	_setSpecialPasteShowOptionsPresentation: function(props){
		var presentation = editor.WordControl.m_oLogicDocument;
		var stateSelection = presentation.GetSelectionState();
		var curPage = stateSelection.CurPage;
		var pos = presentation.GetTargetPosition();
		props = !props ? [Asc.c_oSpecialPasteProps.sourceformatting, Asc.c_oSpecialPasteProps.keepTextOnly] : props;
		var x, y, w, h;
		if (null === pos) {
			pos = presentation.GetSelectedBounds();
			w = pos.w;
			h = pos.h;
			x = pos.x + w;
			y = pos.y + h;
		} else {
			x = pos.X;
			y = pos.Y;
		}
		let screenPos;
		let bThumbnals = presentation.IsFocusOnThumbnails();
		let sSlideId = null;

		let aSelectedSlides = presentation.GetSelectedSlides();
		if(bThumbnals && aSelectedSlides.length > 0) {
			let nSlideIndex = aSelectedSlides[aSelectedSlides.length - 1];
			let oSlide = presentation.GetSlide(nSlideIndex);
			sSlideId = oSlide.Get_Id();
		}
		if(sSlideId) {
			screenPos = editor.WordControl.Thumbnails.getSpecialPasteButtonCoords(sSlideId);
			w = 1;
			h = 1;
		}
		else {
			screenPos = presentation.DrawingDocument.ConvertCoordsToCursorWR(x, y, curPage);
		}

		var specialPasteShowOptions = window['AscCommon'].g_specialPasteHelper.buttonInfo;
		specialPasteShowOptions.asc_setOptions(props);

		var targetDocContent = presentation.Get_TargetDocContent();
		if(targetDocContent && targetDocContent.Id) {
			specialPasteShowOptions.setShapeId(targetDocContent.Id);
		} else {
			specialPasteShowOptions.setShapeId(null);
		}

		var curCoord = new AscCommon.asc_CRect( screenPos.X, screenPos.Y, 0, 0 );
		specialPasteShowOptions.asc_setCellCoord(curCoord);
		specialPasteShowOptions.setFixPosition({x: x, y: y, pageNum: curPage, w: w, h: h, slideId: sSlideId});
	},

    insertInPlace2: function(oDoc, aNewContent)
    {
        var nNewContentLength = aNewContent.length;
		//Часть кода из Document.Add_NewParagraph

        for(var i = 0; i < aNewContent.length; ++i)
        {
            aNewContent[i].Clear_TextFormatting();
            aNewContent[i].Clear_Formatting(true);
        }
        oDoc.Remove(1, true, true);
        var Item = oDoc.Content[oDoc.CurPos.ContentPos];
        if( type_Paragraph === Item.GetType() )
        {
            if(/*true != this.bInBlock &&*/ 1 === nNewContentLength && type_Paragraph === aNewContent[0].GetType() && Item.CurPos.ContentPos !== 1)
            {
				//Вставка строки в параграф
                var oInsertPar = aNewContent[0];
                var nContentLength = oInsertPar.Content.length;
                if(nContentLength > 2)
                {
                    var oFindObj = Item.Internal_FindBackward(Item.CurPos.ContentPos, [para_TextPr]);
                    var TextPr = null;
					if ( true === oFindObj.Found && para_TextPr === oFindObj.Type )
						TextPr = Item.Content[oFindObj.LetterPos].Copy();
					else
						TextPr = new ParaTextPr();
                    var nContentPos = Item.CurPos.ContentPos;
                    for(var i = 0; i < nContentLength - 2; ++i)// -2 на спецсимволы конца параграфа
                    {
                        var oCurInsItem = oInsertPar.Content[i];
                        if(para_Numbering !== oCurInsItem.Type)
                        {
                            Item.Internal_Content_Add(nContentPos, oCurInsItem);
                            nContentPos++;
                        }
                    }
                    Item.Internal_Content_Add(nContentPos, TextPr);
                }
                Item.RecalcInfo.Set_Type_0(pararecalc_0_All);
                Item.RecalcInfo.NeedSpellCheck();
            }
            else
            {
                var LastPos = this.oRecalcDocument.CurPos.ContentPos;
                var LastPosCurDoc = oDoc.CurPos.ContentPos;
				//Нужно разрывать параграф
                var oSourceFirstPar = Item;
                var oSourceLastPar = new Paragraph(oDoc.DrawingDocument, oDoc);
                if(true !== oSourceFirstPar.IsCursorAtEnd() || oSourceFirstPar.IsEmpty())
                    oSourceFirstPar.Split(oSourceLastPar);
                var oInsFirstPar = aNewContent[0];
                var oInsLastPar = null;
                if(nNewContentLength > 1)
                    oInsLastPar = aNewContent[nNewContentLength - 1];

                var nStartIndex = 0;
                var nEndIndex = nNewContentLength - 1;

                if(type_Paragraph === oInsFirstPar.GetType())
                {
					//копируем свойства первого вставляемого параграфа в первый исходный параграф
					//CopyPr_Open - заносим в историю, т.к. этот параграф уже в документе
                    oInsFirstPar.CopyPr_Open( oSourceFirstPar );
					//Копируем содержимое вставляемого параграфа
                    oSourceFirstPar.Concat(oInsFirstPar);
                    if(AscCommon.isRealObject(oInsFirstPar.bullet))
                    {
                        oSourceFirstPar.setPresentationBullet(oInsFirstPar.bullet.createDuplicate());
                    }
					//Сдвигаем стартовый индекс чтобы больше не учитывать этот параграф
                    nStartIndex++;
                }
                else if(type_Table === oInsFirstPar.GetType())
                {
					//если вставляем таблицу в пустой параграф, то не разрываем его
                    if(oSourceFirstPar.IsEmpty())
                    {
                        oSourceFirstPar = null;
                    }
                }
				//Если не скопирован символ конца параграфа, то добавляем содержимое последнего параграфа в начело второй половины разбитого параграфа
                if(null != oInsLastPar && type_Paragraph == oInsLastPar.GetType() && true != this.bInBlock)
                {
                    var nNewContentPos = oInsLastPar.Content.length - 2;
					//копируем свойства последнего исходного параграфа в последний  вставляемый параграф
					//CopyPr - не заносим в историю, т.к. в историю добавится вставка этого параграфа в документ
                    var ind = oInsLastPar.Pr.Ind;
                    if(null != oInsLastPar)
                        oSourceLastPar.CopyPr( oInsLastPar );
                    if(oInsLastPar.bullet)
                    {
                        oInsLastPar.Set_Ind(ind);
                    }
                    oInsLastPar.Concat(oSourceLastPar);
                    oInsLastPar.CurPos.ContentPos = nNewContentPos;
                    oSourceLastPar = oInsLastPar;
                    nEndIndex--;
                }
				//вставляем
                for(var i = nStartIndex; i <= nEndIndex; ++i )
                {
                    var oElemToAdd = aNewContent[i];
                    LastPosCurDoc++;
                    oDoc.Internal_Content_Add(LastPosCurDoc, oElemToAdd);
                }
                if(null != oSourceLastPar)
                {
					//вставляем последний параграф
                    LastPosCurDoc++;
                    oDoc.Internal_Content_Add(LastPosCurDoc, oSourceLastPar);
                }
                if(null == oSourceFirstPar)
                {
					//Удаляем первый параграф, потому что будут ошибки если в документе не будет ни одного параграфа
                    oDoc.Internal_Content_Remove(LastPosCurDoc, 1);
                    LastPosCurDoc--;
                }
                Item.RecalcInfo.Set_Type_0(pararecalc_0_All);
                Item.RecalcInfo.NeedSpellCheck();
                oDoc.CurPos.ContentPos = LastPosCurDoc;
            }
        }

		var content = oDoc.Content;
		for(var  i = 0;  i < content.length; ++i)
		{
            content[i].Recalc_CompiledPr();
			content[i].RecalcInfo.Set_Type_0(pararecalc_0_All);
		}
	},
    ReadFromBinary : function(sBase64, oDocument)
	{
        var oDocumentParams = PasteElementsId.g_bIsDocumentCopyPaste ? this.oDocument : null;
		var openParams = { checkFileSize: false, charCount: 0, parCount: 0, bCopyPaste: true, oDocument: oDocumentParams };
		var doc = oDocument ? oDocument : this.oLogicDocument;
        var oBinaryFileReader = new AscCommonWord.BinaryFileReader(doc, openParams);
        var oRes = oBinaryFileReader.ReadFromString(sBase64, {wordCopyPaste: true});
        this.bInBlock = oRes.bInBlock;

		if(!oRes.content.length && !oRes.aPastedImages.length && !oRes.images.length) {
			oRes = null;
		}

        return oRes;
	},

    SetShortImageId: function(aPastedImages)
    {
        if(!aPastedImages)
			return;

		for(var i = 0, length = aPastedImages.length; i < length; ++i)
        {
            var imageElem = aPastedImages[i];
            if(null != imageElem)
            {
                imageElem.SetUrl(imageElem.Url);
            }
        }
    },

	getRtfImages: function(rtf, html) {

		var getRtfImg = function (sRtf) {
			var res = [];
			var rg_rtf = /\{\\pict[\s\S]+?\\bliptag\-?\d+(\\blipupi\-?\d+)?(\{\\\*\\blipuid\s?[\da-fA-F]+)?[\s\}]*?/, d;
			var rg_rtf_all = new RegExp("(?:(" + rg_rtf.source + "))([\\da-fA-F\\s]+)\\}", "g");
			var pngStr = "\\pngblip";
			var jpegStr = "\\jpegblip";
			var pngTypeStr = "image/png";
			var jpegTypeStr = "image/jpeg";
			var type;

			sRtf = sRtf.match(rg_rtf_all);

			if (!sRtf) {
				return res;
			}

			for (var i = 0; i < sRtf.length; i++) {
				if (rg_rtf.test(sRtf[i])) {
					if (-1 !== sRtf[i].indexOf(jpegStr)) {
						type = pngTypeStr;
					} else if (-1 !== sRtf[i].indexOf(pngStr)) {
						type = jpegTypeStr;
					} else {
						continue;
					}

					res.push({
						data: sRtf[i].replace(rg_rtf, "").replace(/[^\da-fA-F]/g, ""), type: type
					})
				}
			}
			return res
		};

		var getHtmlImg = function (sHtml) {
			var rg_html = /<img[^>]+src="([^"]+)[^>]+/g;
			var res = [];
			var img;
			while (true) {
				img = rg_html.exec(sHtml);
				if (!img) {
					break;
				}
				res.push(img[1]);
			}
			return res;
		};

		function hexToBytes(hex) {
			var res = [];
			for (var i = 0; i < hex.length; i += 2) {
				res.push(parseInt(hex.substr(i, 2), 16));
			}
			return res;
		}

		function bytesToBase64(val) {
			var res = "";
			var bytes = new Uint8Array(val);
			for (var i = 0; i < bytes.byteLength; i++) {
				res += String.fromCharCode(bytes[i]);
			}
			//TODO проверить данный метод на разных браузерах и системах
			return window.btoa(res);
		}

		var rtfImages = getRtfImg(rtf);
		var htmlImages = getHtmlImg(html);
		var map = {};
		if(rtfImages.length === htmlImages.length) {
			for(var i = 0; i < rtfImages.length; i++) {
				var a = rtfImages[i];
				if(a.type) {
					var bytes = hexToBytes(a.data);
					map[htmlImages[i]] = "data:" + a.type + ";base64," + bytesToBase64(bytes);
				}
			}
		}

		return map;
	},

	Start : function (node, nodeDisplay, bDuplicate, fromBinary, text, callback) {
		//PASTE
		var tempPresentation = !PasteElementsId.g_bIsDocumentCopyPaste && editor && editor.WordControl ? editor.WordControl.m_oLogicDocument : null;
		var insertToPresentationWithoutSlides = tempPresentation && tempPresentation.Slides && !tempPresentation.Slides.length;

		//this.oDocument.Remove(1, false, false, false, false, true);

		var base64FromExcel, base64FromWord, base64FromPresentation
		if (PasteElementsId.copyPasteUseBinary) {
			//get binary
			var binaryObj = this._getClassBinaryFromHtml(node, fromBinary);
			base64FromExcel = binaryObj.base64FromExcel;
			base64FromWord = binaryObj.base64FromWord;
			base64FromPresentation = binaryObj.base64FromPresentation;
		}

		if (text) {
			if (insertToPresentationWithoutSlides) {
				window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
				window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
				return;
			}

			//при вставке списка в список, ms вставляет именно html и фильтрует текст(убирает всё, что относится к списку)
			//идея такая, если видим, что вставляем в список, то здесь не делаем pasteText, смотрим на функции prepeare есть ли внутри html списки
			//причём в любом виде - стандартные списки/mso-list
			//если есть, то парсим стандратно html, далее переводим в текст её без элементов списка и вставляем
			//+ если из нас в нас вставляем, тоже отсекам всё что связано со списками

			//TODO pasteTextIntoList - ввожу временно, искючение для вставки текста в список. позже сделать общую отдельную обработку для подобных исключений

			if (PasteElementsId.g_bIsDocumentCopyPaste && (node || ("" !== fromBinary && base64FromWord))) {
				this._initSelectedElem();
				if (this.pasteIntoElem && this.pasteIntoElem.GetNumPr && this.pasteIntoElem.GetNumPr()) {
					this.pasteTextIntoList = text;
				}
			}
			if (!this.pasteTextIntoList) {
				this.oLogicDocument.RemoveBeforePaste();
				this.oDocument = this._GetTargetDocument(this.oDocument);
				this._pasteText(text);
				return;
			}
		}

		var bInsertFromBinary = false;

		if (!node && "" === fromBinary) {
			return;
		}

		if (PasteElementsId.copyPasteUseBinary) {
			var bTurnOffTrackRevisions = false;
			if (PasteElementsId.g_bIsDocumentCopyPaste)//document
			{
				var oThis = this;
				//удаляем в начале, иначе может получиться что будем вставлять в элементы, которое потом удалим.
				//todo с удалением в начале есть проблема, что удаляем элементы даже при пустом буфере

				// Для вставки текста по выделению ячеек таблицы, мы должны сохранить выделенные ячейки
				var oDocState = null;
				if (this.oDocument instanceof CDocument && this.oDocument.IsTableCellSelection())
					oDocState = this.oDocument.SaveDocumentState(false);

				this.oLogicDocument.RemoveBeforePaste();

				if (oDocState)
					this.oDocument.LoadDocumentState(oDocState);

				this.oDocument = this._GetTargetDocument(this.oDocument);


				if (this.oDocument && this.oDocument.bPresentation) {
					if (oThis.api.WordControl.m_oLogicDocument.IsTrackRevisions()) {
						bTurnOffTrackRevisions = oThis.api.WordControl.m_oLogicDocument.GetLocalTrackRevisions();
						oThis.api.WordControl.m_oLogicDocument.SetLocalTrackRevisions(false);
					}
				}
			}

			//paste form word/excel/html into empty presentation without slides
			if (insertToPresentationWithoutSlides && !base64FromPresentation) {
				window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
				window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
				return;
			}

			//insert from binary
			if (base64FromExcel)//вставка из редактора таблиц
			{
				if (PasteElementsId.g_bIsDocumentCopyPaste) {
					bInsertFromBinary = null !== this._pasteBinaryFromExcelToWord(base64FromExcel);
				} else {
					bInsertFromBinary = null !== this._pasteBinaryFromExcelToPresentation(base64FromExcel);
				}
			} else if (base64FromWord)//вставка из редактора документов
			{
				if (PasteElementsId.g_bIsDocumentCopyPaste) {
					bInsertFromBinary = null !== this._pasteBinaryFromWordToWord(base64FromWord, !!(fromBinary));
				} else {
					bInsertFromBinary = null !== this._pasteBinaryFromWordToPresentation(base64FromWord, !!(fromBinary));
				}
			} else if (base64FromPresentation)//вставка из редактора презентаций
			{
				if (PasteElementsId.g_bIsDocumentCopyPaste) {
					bInsertFromBinary = null !== this._pasteBinaryFromPresentationToWord(base64FromPresentation, bDuplicate);
				} else {
					bInsertFromBinary = null !== this._pasteBinaryFromPresentationToPresentation(base64FromPresentation);
				}
			}
		}

		if (true === bInsertFromBinary) {
			if (false !== bTurnOffTrackRevisions) {
				oThis.api.WordControl.m_oLogicDocument.SetLocalTrackRevisions(bTurnOffTrackRevisions);
			}
		} else if (node) {
			this._pasteFromHtml(node, bTurnOffTrackRevisions);
		} else {
			window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
			window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
		}
	},

	//from EXCEL to WORD
	_pasteBinaryFromExcelToWord: function (base64FromExcel) {
		var oThis = this;

		var fPrepasteCallback = function () {
			if (false === oThis.bNested) {
				oThis.InsertInDocument();
				if (oThis.aContent.bAddNewStyles) {
					oThis.api.GenerateStyles();
				}
				if (oThis.pasteCallback) {
					oThis.pasteCallback();
				}
			}
		};

		History.TurnOff();
		var aContentExcel = this._readFromBinaryExcel(base64FromExcel);
		History.TurnOn();


		if (null === aContentExcel) {
			return null;
		}

		var oldLocale = AscCommon.g_oDefaultCultureInfo ? AscCommon.g_oDefaultCultureInfo.LCID : AscCommon.g_oDefaultCultureInfo;
		AscCommon.setCurrentCultureInfo(aContentExcel.workbook.Core.language);
		var revertLocale = function () {
			if (oldLocale) {
				AscCommon.setCurrentCultureInfo(oldLocale);
			} else {
				AscCommon.g_oDefaultCultureInfo = oldLocale;
			}
		};

		var aContent;
		if (window['AscCommon'].g_specialPasteHelper.specialPasteStart &&
			Asc.c_oSpecialPasteProps.keepTextOnly === window['AscCommon'].g_specialPasteHelper.specialPasteProps) {
			aContent = oThis._convertExcelBinary(aContentExcel);
			revertLocale();

			oThis.aContent = aContent.content;
			fPrepasteCallback();
		} else {
            let oObjectsForDownload = null;
            if(aContentExcel.arrImages && aContentExcel.arrImages.length) {
                oObjectsForDownload = GetObjectsForImageDownload(aContentExcel.arrImages);
            }
            if (oObjectsForDownload && oObjectsForDownload.aUrls.length > 0) {
                AscCommon.sendImgUrls(oThis.api, oObjectsForDownload.aUrls, function (data) {
                    var oImageMap = {};
                    ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);
                    var aContent = oThis._convertExcelBinary(aContentExcel, aContentExcel ? aContentExcel.pDrawings : null);
                    revertLocale();
                    oThis.aContent = aContent.content;
	                addThemeImagesToMap(oImageMap, oObjectsForDownload.aUrls, aContentExcel.arrImages);
                    oThis.api.pre_Paste(aContent.fonts, oImageMap, fPrepasteCallback);
                }, true);
            } else {
                aContent = oThis._convertExcelBinary(aContentExcel);
                revertLocale();
                oThis.aContent = aContent.content;
                oThis.api.pre_Paste(aContent.fonts, aContent.images, fPrepasteCallback);
            }
        }
	},

	//from EXCEL to PRESENTATION
	_pasteBinaryFromExcelToPresentation: function (base64FromExcel) {
		var oThis = this;
		var presentation = editor.WordControl.m_oLogicDocument;

		var excelContent = AscFormat.ExecuteNoHistory(this._readFromBinaryExcel, this, [base64FromExcel]);
		if (null === excelContent) {
			return null;
		}

		var aContentExcel = excelContent.workbook;
		var aPastedImages = excelContent.arrImages;

		//если есть шейпы, то вставляем их из excel
		var aContent;
		var _sheet = aContentExcel && aContentExcel.aWorksheets && aContentExcel.aWorksheets[0];
		var drawings = excelContent.pDrawings ? excelContent.pDrawings : _sheet && _sheet.Drawings;
		if (drawings && drawings.length) {
			var paste_callback = function () {
				if (false === oThis.bNested) {
					var oIdMap = {};
					var aCopies = [];
					var oCopyPr = new AscFormat.CCopyObjectProperties();
					oCopyPr.idMap = oIdMap;
					var l = null, t = null, r = null, b = null, oXfrm;

					for (var i = 0; i < arr_shapes.length; ++i) {
						shape = arr_shapes[i].graphicObject.copy(oCopyPr);
						aCopies.push(shape);
						oIdMap[arr_shapes[i].graphicObject.Id] = shape.Id;
						shape.worksheet = null;
						shape.drawingBase = null;

						arr_shapes[i] = new DrawingCopyObject(shape, 0, 0, 0, 0);
						if (shape.spPr && shape.spPr.xfrm && AscFormat.isRealNumber(shape.spPr.xfrm.offX) && AscFormat.isRealNumber(shape.spPr.xfrm.offY)
							&& AscFormat.isRealNumber(shape.spPr.xfrm.extX) && AscFormat.isRealNumber(shape.spPr.xfrm.extY)) {
							oXfrm = shape.spPr.xfrm;
							if (l === null) {
								l = oXfrm.offX;
							} else {
								if (oXfrm.offX < l) {
									l = oXfrm.offX;
								}
							}
							if (t === null) {
								t = oXfrm.offY;
							} else {
								if (oXfrm.offY < t) {
									t = oXfrm.offY;
								}
							}
							if (r === null) {
								r = oXfrm.offX + oXfrm.extX;
							} else {
								if (oXfrm.offX + oXfrm.extX > r) {
									r = oXfrm.offX + oXfrm.extX;
								}
							}
							if (b === null) {
								b = oXfrm.offY + oXfrm.extY;
							} else {
								if (oXfrm.offY + oXfrm.extY > b) {
									b = oXfrm.offY + oXfrm.extY;
								}
							}
						}
					}
					if (AscFormat.isRealNumber(l) && AscFormat.isRealNumber(t)
						&& AscFormat.isRealNumber(r) && AscFormat.isRealNumber(b)) {
						var fSlideCX = presentation.GetWidthMM() / 2.0;
						var fSlideCY = presentation.GetHeightMM() / 2.0;
						var fBoundsCX = (r + l) / 2.0;
						var fBoundsCY = (t + b) / 2.0;
						var fDiffX = fBoundsCX - fSlideCX;
						var fDiffY = fBoundsCY - fSlideCY;
						if (!AscFormat.fApproxEqual(fDiffX, 0) || !AscFormat.fApproxEqual(fDiffY, 0)) {
							for (var i = 0; i < arr_shapes.length; ++i) {
								shape = arr_shapes[i].Drawing;
								if (shape.spPr && shape.spPr.xfrm && AscFormat.isRealNumber(shape.spPr.xfrm.offX) && AscFormat.isRealNumber(shape.spPr.xfrm.offY)
									&& AscFormat.isRealNumber(shape.spPr.xfrm.extX) && AscFormat.isRealNumber(shape.spPr.xfrm.extY)) {
									shape.spPr.xfrm.setOffX(shape.spPr.xfrm.offX - fDiffX);
									shape.spPr.xfrm.setOffY(shape.spPr.xfrm.offY - fDiffY);
								}
							}
						}
					}
					AscFormat.fResetConnectorsIds(aCopies, oIdMap);

					var presentationSelectedContent = new PresentationSelectedContent();
					presentationSelectedContent.Drawings = arr_shapes;


					if (presentation.InsertContent(presentationSelectedContent)) {
						presentation.Recalculate();
                        editor.checkChangesSize();
						presentation.Document_UpdateInterfaceState();
					} else {
						window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
					}

					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
				}
			};


			var arr_shapes = drawings;

			var aImagesToDownload = [];
			for (var i = 0; i < aPastedImages.length; i++) {
				aImagesToDownload.push(aPastedImages[i].Url);
			}

			aContent = {aPastedImages: aPastedImages, images: aImagesToDownload};

			//fonts
			var font_map = {};
			for (var i = 0; i < arr_shapes.length; ++i) {
				var shape = arr_shapes[i].graphicObject;
				if (shape) {
					shape.getAllFonts(font_map);
				}
			}

			var fonts = [];
			//грузим картинки и фонты
			for (var i in font_map) {
				fonts.push(new CFont(i));
			}

			//images
			var images = aContent.images;
			var arrImages = aContent.aPastedImages;
			var oObjectsForDownload = GetObjectsForImageDownload(arrImages);
			if (oObjectsForDownload.aUrls.length > 0) {
				AscCommon.sendImgUrls(oThis.api, oObjectsForDownload.aUrls, function (data) {
					var oImageMap = {};
					ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);

					addThemeImagesToMap(oImageMap, oObjectsForDownload.aUrls, arrImages);
					oThis.api.pre_Paste(fonts, oImageMap, paste_callback);
				}, true);
			} else {
				this.SetShortImageId(arrImages);
				this.api.pre_Paste(fonts, images, paste_callback);
			}
		} else {
			var presentationSelectedContent = new PresentationSelectedContent();
			presentationSelectedContent.DocContent = new AscCommonWord.CSelectedContent();

			aContent = AscFormat.ExecuteNoHistory(this._convertExcelBinary, this, [excelContent]);

			var selectedElement, element, pDrawings = [], drawingCopyObject;
			//var defaultTableStyleId = presentation.DefaultTableStyleId;
			for (var i = 0; i < aContent.content.length; ++i) {
				selectedElement = new AscCommonWord.CSelectedElement();
				element = aContent.content[i];

				if (type_Table === element.GetType())//table
				{
					//TODO переделать количество строк и ширину
					var W = 100;
					var Rows = 3;
					var graphic_frame = new AscFormat.CGraphicFrame();
					graphic_frame.setSpPr(new AscFormat.CSpPr());
					graphic_frame.spPr.setParent(graphic_frame);
					graphic_frame.spPr.setXfrm(new AscFormat.CXfrm());
					graphic_frame.spPr.xfrm.setParent(graphic_frame.spPr);
					graphic_frame.spPr.xfrm.setOffX((this.oDocument.GetWidthMM() - W) / 2);
					graphic_frame.spPr.xfrm.setOffY(this.oDocument.GetHeightMM() / 5);
					graphic_frame.spPr.xfrm.setExtX(W);
					graphic_frame.spPr.xfrm.setExtY(7.478268771701388 * Rows);
					graphic_frame.setNvSpPr(new AscFormat.UniNvPr());

					element = this._convertTableToPPTX(element);
					graphic_frame.setGraphicObject(element.Copy(graphic_frame));
					//graphic_frame.graphicObject.Set_TableStyle(presentation.DefaultTableStyleId);

					drawingCopyObject = new DrawingCopyObject();
					drawingCopyObject.Drawing = graphic_frame;
					pDrawings.push(drawingCopyObject);
				}
			}
			presentationSelectedContent.Drawings = pDrawings;

			//вставка
			var paste_callback_presentation = function () {
				if (false == oThis.bNested) {

					if (presentation.InsertContent(presentationSelectedContent)) {
						presentation.Recalculate();
                        editor.checkChangesSize();
						presentation.Document_UpdateInterfaceState();

						var props = [Asc.c_oSpecialPasteProps.destinationFormatting, Asc.c_oSpecialPasteProps.keepTextOnly];
						oThis._setSpecialPasteShowOptionsPresentation(props);
					} else {
						window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
					}

					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
				}
			};

			oThis.api.pre_Paste(aContent.fonts, null, paste_callback_presentation);
		}
	},

	//from WORD to WORD
	_pasteBinaryFromWordToWord: function (base64FromWord, bIsOnlyFromBinary) {
		var oThis = this;
		//при чтении документа создаётся новый DocPart, который добавляется в DocParts, но не добавляется в g_oTableId
		//чтобы он добавлялся в g_oTableId и соответсвенно в историю, делаю ему Copy()

		//но далее вызывается функция InsertInDocument-> ... -> CheckDocPartNames, где берутся все СС во вставляемом фрагменте
		//смотрится есть ли плесйхолдер и если он есть, то переименовывается плейсходер и DocPart->Name
		//поскольку в DocParts уже есть несколько DocPart с одинаковым именем(один при чтении, второй и при Copy)
		//то при попытке переименования берётся первый DocPart(который добавлен в DocParts, но не в g_oTableId)-> и переименовывается
		//соответсвенно далее при накатывании изменений в g_oTableId отсутвует DocPart с нужным именем
		//чтобы от этой проблемы уйти - удаляю тот первый, который был создан при чтении - delete glossaryDoc.DocParts[aDelIndexes[i]];

		//но появляется другая проблема - при повтроном копировании делается CDocPart-> Copy и снова создаётся DocPart с таким именем
		//и добавляется в DocParts. далее при вставке снова всё повторяется...
		var aContent = this.ReadFromBinary(base64FromWord);
		if (null === aContent) {
			return null;
		}

		if (Asc.c_oSpecialPasteProps.keepTextOnly === window['AscCommon'].g_specialPasteHelper.specialPasteProps) {
			this.oLogicDocument.RemoveBeforePaste();
			this.oDocument = this._GetTargetDocument(this.oDocument);

			var oPr = {NewLineParagraph: true, Numbering: true};

			this._initSelectedElem();
			if (this.pasteIntoElem && this.pasteIntoElem.GetNumPr && this.pasteIntoElem.GetNumPr()) {
				oPr.Numbering = false;
			}

			this._pasteText(this._getTextFromContent(aContent.content, oPr));
			return;
		}

		//вставляем в заголовок диаграммы, предварительно конвертируем все параграфы в презентационный формат
		if (aContent && aContent.content && this.oDocument.bPresentation && oThis.oDocument && oThis.oDocument.Parent &&
			oThis.oDocument.Parent.parent && oThis.oDocument.Parent.parent.parent &&
			oThis.oDocument.Parent.parent.parent.getObjectType &&
			oThis.oDocument.Parent.parent.parent.getObjectType() === AscDFH.historyitem_type_Chart) {

			//не грузим изображения при вставке в заголовок диаграммы
			aContent.images = [];
			aContent.aPastedImages = [];

			var newContent = [];
			for (var i = 0; i < aContent.content.length; i++) {
				if (type_Paragraph === aContent.content[i].Get_Type()) {
					newContent.push(
						AscFormat.ConvertParagraphToPPTX(aContent.content[i], this.oDocument.DrawingDocument,
							this.oDocument, false, true));
				}
			}

			aContent.content = newContent;
		}

		var fPrepasteCallback = function () {
			if (false === oThis.bNested) {
				oThis.InsertInDocument();
				if (aContent.bAddNewStyles) {
					oThis.api.GenerateStyles();
				}
				if (oThis.pasteCallback) {
					oThis.pasteCallback();
				}
			}
		};

		this.aContent = aContent.content;
		//проверяем список фонтов
		aContent.fonts = oThis._checkFontsOnLoad(aContent.fonts);

		var oObjectsForDownload = GetObjectsForImageDownload(aContent.aPastedImages);
		if (window['AscCommon'].g_specialPasteHelper.specialPasteStart &&
			Asc.c_oSpecialPasteProps.keepTextOnly === window['AscCommon'].g_specialPasteHelper.specialPasteProps) {
			oThis.api.pre_Paste([], [], fPrepasteCallback);
		} else if (oObjectsForDownload.aUrls.length > 0) {
			if (window["IS_NATIVE_EDITOR"]) {
				oThis.api.pre_Paste(aContent.fonts, aContent.images, fPrepasteCallback);
			} else if (bIsOnlyFromBinary && window["NativeCorrectImageUrlOnPaste"]) {
				var url;
				for (var i = 0; i < aContent.aPastedImages.length; ++i) {
					url = window["NativeCorrectImageUrlOnPaste"](aContent.aPastedImages[i].Url);
					aContent.images[i] = url;

					var imageElem = aContent.aPastedImages[i];
					if (null != imageElem) {
						imageElem.SetUrl(url);
					}
				}
				oThis.api.pre_Paste(aContent.fonts, aContent.images, fPrepasteCallback);
			} else {
				AscCommon.sendImgUrls(oThis.api, oObjectsForDownload.aUrls, function (data) {
					ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl,
						aContent.images);
					oThis.api.pre_Paste(aContent.fonts, aContent.images, fPrepasteCallback);
				}, true);
			}
		} else {
			oThis.SetShortImageId(aContent.aPastedImages);
			oThis.api.pre_Paste(aContent.fonts, aContent.images, fPrepasteCallback);
		}
	},

	//from WORD to PRESENTATION
	_pasteBinaryFromWordToPresentation: function (base64FromWord) {
		var oThis = this;
		var presentation = editor.WordControl.m_oLogicDocument;
		var trueDocument = this.oDocument;

		var tempCDocument = function () {
			return new CDocument(oThis.oDocument.DrawingDocument, false);
		};
		//создаём темповый CDocument
		this.oDocument = AscFormat.ExecuteNoHistory(tempCDocument, this, []);

		AscCommon.g_oIdCounter.m_bRead = true;
		var aContent = AscFormat.ExecuteNoHistory(this.ReadFromBinary, this, [base64FromWord, this.oDocument]);
		AscCommon.g_oIdCounter.m_bRead = false;

		if (null === aContent) {
			return null;
		}

		//возврщаем обратно переменные и историю, документ которой заменяется при создании CDocument
		this.oDocument = trueDocument;
		History.Document = trueDocument;

		var presentationSelectedContent = new PresentationSelectedContent();
		presentationSelectedContent.DocContent = new AscCommonWord.CSelectedContent();

		var parseContent = function (content) {
			for (var i = 0; i < content.length; ++i) {
				selectedElement = new AscCommonWord.CSelectedElement();
				element = content[i];
				//drawings
				element.GetAllDrawingObjects(drawings);
				if (type_Paragraph === element.GetType())//paragraph
				{
					selectedElement.Element = AscFormat.ConvertParagraphToPPTX(element, null, null, true, false);
					elements.push(selectedElement);
				} else if (type_Table === element.GetType())//table
				{
					element = oThis._convertTableToPPTX(element, true);

					//TODO переделать количество строк и ширину
					var W = oThis.oDocument.GetWidthMM() / 1.45;
					var Rows = element.GetRowsCount();
					var H = Rows * 7.478268771701388;
					var graphic_frame = new AscFormat.CGraphicFrame();
					graphic_frame.setSpPr(new AscFormat.CSpPr());
					graphic_frame.spPr.setParent(graphic_frame);
					graphic_frame.spPr.setXfrm(new AscFormat.CXfrm());
					graphic_frame.spPr.xfrm.setParent(graphic_frame.spPr);
					graphic_frame.spPr.xfrm.setOffX(oThis.oDocument.GetWidthMM() / 2 - W / 2);
					graphic_frame.spPr.xfrm.setOffY(oThis.oDocument.GetHeightMM() / 2 - H / 2);
					graphic_frame.spPr.xfrm.setExtX(W);
					graphic_frame.spPr.xfrm.setExtY(H);
					graphic_frame.setNvSpPr(new AscFormat.UniNvPr());

					graphic_frame.setGraphicObject(element.Copy(graphic_frame));
					graphic_frame.graphicObject.Set_TableStyle(defaultTableStyleId);

					drawingCopyObject = new DrawingCopyObject();
					drawingCopyObject.Drawing = graphic_frame;
					pDrawings.push(drawingCopyObject);

				} else if (type_BlockLevelSdt === element.GetType())//TOC
				{
					parseContent(element.Content.Content);
				}
			}
		};

		var elements = [], selectedElement, element, drawings = [], pDrawings = [], drawingCopyObject;
		var defaultTableStyleId = presentation.DefaultTableStyleId;
		parseContent(aContent.content);

		var onlyImages = false;
		if (drawings && drawings.length) {
			//если массив содержит только изображения
			if (elements && 1 === elements.length && elements[0].Element && type_Paragraph === elements[0].Element.Get_Type()) {
				if (true === this._isParagraphContainsOnlyDrawing(elements[0].Element)) {
					elements = [];
					onlyImages = true;
				}
			}

			for (var j = 0; j < drawings.length; j++) {
				drawingCopyObject = new DrawingCopyObject();
				drawingCopyObject.Drawing = drawings[j].GraphicObj;
				pDrawings.push(drawingCopyObject);
			}
		}
		presentationSelectedContent.DocContent.Elements = elements;
		presentationSelectedContent.Drawings = pDrawings;

		//вставка
		var paste_callback = function () {
			if (false === oThis.bNested) {
				//для таблиц необходимо рассчитать их размер, чтобы разместить в центре
				var slide = presentation.Slides[0];
				for (var i = 0; i < presentationSelectedContent.Drawings.length; i++) {
					if (presentationSelectedContent.Drawings[i].Drawing instanceof AscFormat.CGraphicFrame) {
						var drawing = presentationSelectedContent.Drawings[i].Drawing;
						var oldParent = drawing.parent;
						var bDeleted = drawing.bDeleted;
						drawing.parent = slide;
						drawing.bDeleted = false;
						drawing.recalculate();
						drawing.parent = oldParent;
						drawing.bDeleted = bDeleted;

						drawing.spPr.xfrm.setOffX(oThis.oDocument.GetWidthMM() / 2 - drawing.extX / 2);
						drawing.spPr.xfrm.setOffY(oThis.oDocument.GetHeightMM() / 2 - drawing.extY / 2);

					}
				}

				if (presentation.InsertContent(presentationSelectedContent)) {
					presentation.Recalculate();
                    editor.checkChangesSize();
					presentation.Document_UpdateInterfaceState();

					if (!onlyImages) {
						var props = [Asc.c_oSpecialPasteProps.destinationFormatting, Asc.c_oSpecialPasteProps.keepTextOnly];
						oThis._setSpecialPasteShowOptionsPresentation(props);
					} else {
						window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
					}
				} else {
					window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
				}

				window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
			}
		};


		var font_map = {};
		var images = [];
		//shape.getAllFonts(font_map);

		//перебираем шрифты
		var fonts = [];
		for (var i in font_map)
			fonts.push(new CFont(i));

		var oObjectsForDownload = GetObjectsForImageDownload(aContent.aPastedImages);
		if (oObjectsForDownload.aUrls.length > 0) {
			AscCommon.sendImgUrls(oThis.api, oObjectsForDownload.aUrls, function (data) {
				var oImageMap = {};
				ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);
				//ковертим изображения в презентационный формат
				for (var i = 0; i < presentationSelectedContent.Drawings.length; i++) {
					if (!(presentationSelectedContent.Drawings[i].Drawing instanceof AscFormat.CGraphicFrame)) {
						AscFormat.ExecuteNoHistory(function () {
							if (presentationSelectedContent.Drawings[i].Drawing.setBDeleted2) {
								presentationSelectedContent.Drawings[i].Drawing.setBDeleted2(true);
							} else {
								presentationSelectedContent.Drawings[i].Drawing.setBDeleted(true);
							}
						}, this, []);
						presentationSelectedContent.Drawings[i].Drawing = presentationSelectedContent.Drawings[i].Drawing.convertToPPTX(oThis.oDocument.DrawingDocument, undefined, true);
						AscFormat.checkBlipFillRasterImages(presentationSelectedContent.Drawings[i].Drawing);
					}
				}
				addThemeImagesToMap(oImageMap, oObjectsForDownload.aUrls, aContent.aPastedImages);
				oThis.api.pre_Paste(fonts, oImageMap, paste_callback);
			}, true);
		} else {
			//ковертим изображения в презентационный формат
			for (var i = 0; i < presentationSelectedContent.Drawings.length; i++) {
				if (!(presentationSelectedContent.Drawings[i].Drawing instanceof AscFormat.CGraphicFrame)) {
					presentationSelectedContent.Drawings[i].Drawing = presentationSelectedContent.Drawings[i].Drawing.convertToPPTX(oThis.oDocument.DrawingDocument, undefined, true);
					AscFormat.checkBlipFillRasterImages(presentationSelectedContent.Drawings[i].Drawing);
				}
			}

			oThis.api.pre_Paste(aContent.fonts, aContent.images, paste_callback);
		}
	},

	//from PRESENTATION to WORD
	_pasteBinaryFromPresentationToWord: function (base64) {
		var oThis = this;

		var fPrepasteCallback = function () {
			if (false === oThis.bNested) {
				oThis.InsertInDocument();
				if (oThis.pasteCallback) {
					oThis.pasteCallback();
				}
			}
		};

		var oSelectedContent2 = this._readPresentationSelectedContent2(base64);
		var selectedContent2 = oSelectedContent2.content;
		var defaultSelectedContent = selectedContent2[0] ? selectedContent2[0] : selectedContent2[1];
		var bSlideObjects = defaultSelectedContent && defaultSelectedContent.content && defaultSelectedContent.content.SlideObjects && defaultSelectedContent.content.SlideObjects.length > 0;
		var pasteObj = bSlideObjects && PasteElementsId.g_bIsDocumentCopyPaste ? selectedContent2[2] : defaultSelectedContent;

		var arr_Images, fonts, content = null, font_map = {};
		if (pasteObj) {
			arr_Images = pasteObj.images;
			fonts = pasteObj.fonts;
			content = pasteObj.content;
		}

		if (null === content) {
			return null;
		}

		if (content && content.DocContent) {
			let aElements = content.DocContent.Elements;
			let aContent = [];
			let bNeedDocElement = !this.oDocument.bPresentation;
			let oNewParent = this.oDocument;
			for (let nElement = 0; nElement < aElements.length; nElement++) {
				let oContentElement = aElements[nElement].Element;
				if(bNeedDocElement) {
					aContent.push(AscFormat.ConvertParagraphToWord(oContentElement, oNewParent));
				} else {
					aContent.push(oContentElement.Copy(oNewParent, oNewParent.DrawingDocument));
				}
			}
			this.aContent = aContent;
			oThis.api.pre_Paste(fonts, arr_Images, fPrepasteCallback);

		} else if (content && content.Drawings) {

			var arr_shapes = content.Drawings;
			var arrImages = pasteObj.images;
			if (!bSlideObjects && content.Drawings.length === selectedContent2[1].content.Drawings.length) {
				AscFormat.checkDrawingsTransformBeforePaste(content, selectedContent2[1].content, null);
			}
			//****если записана одна табличка, то вставляем html и поддерживаем все цвета и стили****
			if (!arrImages.length && arr_shapes.length === 1 && arr_shapes[0] && arr_shapes[0].Drawing &&
				arr_shapes[0].Drawing.graphicObject) {

				var drawing = arr_shapes[0].Drawing;

				if (typeof CGraphicFrame !== "undefined" && drawing instanceof CGraphicFrame) {
					var aContent = [];
					var table = AscFormat.ConvertGraphicFrameToWordTable(drawing, this.oLogicDocument);
					table.Document_Get_AllFontNames(font_map);

					//перебираем шрифты
					for (var i in font_map) {
						fonts.push(new CFont(i));
					}

					//TODO стиль не прокидывается. в будущем нужно реализовать
					table.TableStyle = null;
					aContent.push(table);

					this.aContent = aContent;
					oThis.api.pre_Paste(fonts, aContent.images, fPrepasteCallback);

					return;
				}
			}


			//если несколько графических объектов, то собираем base64 у таблиц(graphicFrame)
			if (arr_shapes.length > 1) {
				for (var i = 0; i < arr_shapes.length; i++) {
					if (arr_shapes[i].Drawing && arr_shapes[i].Drawing.isTable()) {
						arrImages.push(new AscCommon.CBuilderImages(null, arr_shapes[i].base64, arr_shapes[i], null, null));
					}
				}
			}

			var oObjectsForDownload = GetObjectsForImageDownload(arrImages);
			var aImagesToDownload = oObjectsForDownload.aUrls;

			AscCommon.sendImgUrls(oThis.api, aImagesToDownload, function (data) {
				var image_map = {};
				for (var i = 0, length = Math.min(data.length, arrImages.length); i < length; ++i) {
					var elem = data[i];
					if (null != elem.url) {
						var name = g_oDocumentUrls.imagePath2Local(elem.path);
						var imageElem = oObjectsForDownload.aBuilderImagesByUrl[i];
						if (null != imageElem) {
							if (Array.isArray(imageElem)) {
								for (var j = 0; j < imageElem.length; ++j) {
									var curImageElem = imageElem[j];
									if (null != curImageElem) {
										if (curImageElem.ImageShape && curImageElem.ImageShape.base64) {
											curImageElem.ImageShape.base64 = name;
										} else {
											curImageElem.SetUrl(name);
										}
									}
								}
							} else {
								//для вставки graphicFrame в виде картинки(если было при копировании выделено несколько графических объектов)
								if (imageElem.ImageShape && imageElem.ImageShape.base64) {
									imageElem.ImageShape.base64 = name;
								} else {
									imageElem.SetUrl(name);
								}
							}
						}
						image_map[i] = name;
					} else {
						image_map[i] = aImagesToDownload[i];
					}
				}

				aContent = oThis._convertExcelBinary(null, arr_shapes);
				oThis.aContent = aContent.content;
				oThis.api.pre_Paste(fonts, image_map, fPrepasteCallback);
			}, true);
		}
	},

	//from PRESENTATION to PRESENTATION
	_pasteBinaryFromPresentationToPresentation: function (base64, bDuplicate) {
		var oThis = this;
		var presentation = editor.WordControl.m_oLogicDocument;

		var oSelectedContent2 = this._readPresentationSelectedContent2(base64, bDuplicate);
		var p_url = oSelectedContent2.p_url;
		var p_theme = oSelectedContent2.p_theme;
		var selectedContent2 = oSelectedContent2.content;
		var multipleParamsCount = selectedContent2 ? selectedContent2.length : 0;
		if (multipleParamsCount) {
			var aContents = [];
			for (var i = 0; i < multipleParamsCount; i++) {
				var curContent = selectedContent2[i];
				aContents.push(curContent.content);
			}

			var specialOptionsArr = [];
			var specialProps = Asc.c_oSpecialPasteProps;
			if (1 === multipleParamsCount) {
				specialOptionsArr = [specialProps.destinationFormatting];
			} else if (2 === multipleParamsCount) {
				specialOptionsArr = [specialProps.destinationFormatting, specialProps.sourceformatting];
			} else if (3 === multipleParamsCount) {
				specialOptionsArr = [specialProps.destinationFormatting, specialProps.sourceformatting, specialProps.picture];
			}

			var pasteObj = selectedContent2[0];
			var nIndex = 0;
			if (window['AscCommon'].g_specialPasteHelper.specialPasteStart) {
				var props = window['AscCommon'].g_specialPasteHelper.specialPasteProps;
				switch (props) {
					case Asc.c_oSpecialPasteProps.destinationFormatting: {
						break;
					}
					case Asc.c_oSpecialPasteProps.sourceformatting: {
						if (selectedContent2[1]) {
							pasteObj = selectedContent2[1];
							nIndex = 1;
						}
						break;
					}
					case Asc.c_oSpecialPasteProps.picture: {
						if (selectedContent2[2]) {
							pasteObj = selectedContent2[2];
							nIndex = 2;
						}
						break;
					}
					case Asc.c_oSpecialPasteProps.keepTextOnly: {
						//в идеале у этом случае нужно использовать данные plain text из буфера обмена
						//pasteObj = selectedContent2[2];
						break;
					}
				}
			}

			var arr_Images = pasteObj.images;
			var fonts = pasteObj.fonts;
			var presentationSelectedContent = pasteObj.content;

			if (null === presentationSelectedContent) {
				return null;
			}

			if (presentationSelectedContent.Drawings && presentationSelectedContent.Drawings.length > 0) {
				var controller = this.oDocument.GetCurrentController();
				var curTheme = controller ? controller.getTheme() : null;
				if (curTheme && curTheme.name === p_theme) {
					specialOptionsArr.splice(1, 1);
				}
			}

			var paste_callback = function () {
				if (false === oThis.bNested) {
					var bPaste = presentation.InsertContent2(aContents, nIndex);

					presentation.Recalculate();
                    editor.checkChangesSize();
					presentation.Document_UpdateInterfaceState();

					//пока не показываю значок специальной вставки после copy/paste слайдов
					var bSlideObjects = aContents[nIndex] && aContents[nIndex].SlideObjects && aContents[nIndex].SlideObjects.length > 0;
					if (specialOptionsArr.length >= 1 /*&& !bSlideObjects*/ && bPaste) {
						if (presentationSelectedContent && presentationSelectedContent.DocContent) {
							specialOptionsArr.push(Asc.c_oSpecialPasteProps.keepTextOnly);
						}

						oThis._setSpecialPasteShowOptionsPresentation(specialOptionsArr);
					} else {
						window['AscCommon'].g_specialPasteHelper.CleanButtonInfo();
					}

					window['AscCommon'].g_specialPasteHelper.Paste_Process_End();
				}
			};

			var oObjectsForDownload = GetObjectsForImageDownload(arr_Images, p_url === this.api.documentId);
			if (oObjectsForDownload.aUrls.length > 0) {
				AscCommon.sendImgUrls(oThis.api, oObjectsForDownload.aUrls, function (data) {
					let oImageMap = {};
					ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);
					addThemeImagesToMap(oImageMap, oObjectsForDownload.aUrls, arr_Images);
					oThis.api.pre_Paste(fonts, oImageMap, paste_callback);
				}, true);
			} else {
				let oImageMap = {};
				for(let nImg = 0; nImg < arr_Images.length; ++nImg) {
					oImageMap[nImg] = arr_Images[nImg].Url
				}
				oThis.api.pre_Paste(fonts, oImageMap, paste_callback);
			}
		} else {
			return null;
		}
	},

	_readPresentationSelectedContent2: function (base64, bDuplicate) {
		pptx_content_loader.Clear();

		var _stream = AscFormat.CreateBinaryReader(base64, 0, base64.length);
		var stream = new AscCommon.FileStream(_stream.data, _stream.size);
		var p_url = stream.GetString2();
		var p_theme = stream.GetString2();
		var p_width = stream.GetULong() / 100000;
		var p_height = stream.GetULong() / 100000;

		var bIsMultipleContent = stream.GetBool();
		var selectedContent2 = [];
		if (true === bIsMultipleContent) {
			var multipleParamsCount = stream.GetULong();
			for (var i = 0; i < multipleParamsCount; i++) {
				selectedContent2.push(this._readPresentationSelectedContent(stream, bDuplicate));
			}
		}
		return {content: selectedContent2, p_url: p_url, p_theme: p_theme};
	},

	_readPresentationSelectedContent: function (stream, bDuplicate) {
		return AscFormat.ExecuteNoHistory(function () {
			var presentationSelectedContent = null;
			var fonts = [];
			var arr_Images = [];
			var oThis = this;
			var oFontMap = {};

			var readContent = function () {
				var docContent = oThis.ReadPresentationText(stream);
				if (docContent.length === 0) {
					return;
				}
				presentationSelectedContent.DocContent = new AscCommonWord.CSelectedContent();
				presentationSelectedContent.DocContent.Elements = docContent;

				//перебираем шрифты
				for (var i in oThis.oFonts) {
					oFontMap[i] = 1;
				}

				bIsEmptyContent = false;
			};

			var readDrawings = function () {

				if (PasteElementsId.g_bIsDocumentCopyPaste) {
					History.TurnOff();
				}
				var objects = oThis.ReadPresentationShapes(stream);
				if (PasteElementsId.g_bIsDocumentCopyPaste) {
					History.TurnOn();
				}

				presentationSelectedContent.Drawings = objects.arrShapes;

				var arr_shapes = objects.arrShapes;
				for (var i = 0; i < arr_shapes.length; ++i) {
					if (arr_shapes[i].Drawing.getAllFonts) {
						arr_shapes[i].Drawing.getAllFonts(oFontMap);
					}
				}
				arr_Images = arr_Images.concat(objects.arrImages);
			};

			var readSlideObjects = function () {
				var arr_slides = [];
				var loader = new AscCommon.BinaryPPTYLoader();
				if (!(bDuplicate === true)) {
					loader.Start_UseFullUrl();
				}
				loader.stream = stream;
				loader.presentation = editor.WordControl.m_oLogicDocument;
				loader.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;

				//для вставки таблицы со стилем
				//TODO в дальнейшем необходимо пересмотреть и писать стили вместе со слайдом
				var _globalTableStyles = editor.WordControl.m_oLogicDocument.globalTableStyles;
				if (_globalTableStyles) {
					for (var key in _globalTableStyles.Style) {
						if (_globalTableStyles.Style.hasOwnProperty(key)) {
							loader.map_table_styles[_globalTableStyles.Style[key].Id] =_globalTableStyles.Style[key];
						}
					}
				}

				//read slides
				var slide_count = stream.GetULong();
				//var arr_arrTransforms = [];

				for (var i = 0; i < slide_count; ++i) {
					if (PasteElementsId.g_bIsDocumentCopyPaste) {
						loader.stream.GetUChar();
						loader.stream.SkipRecord();
						arr_slides[i] = null;
					} else {
						arr_slides[i] = loader.ReadSlide(0);
					}
				}

				//images and fonts
				var font_map = {};
				var slideCopyObjects = [];
				for (var i = 0; i < arr_slides.length; ++i) {
					if (arr_slides[i] && arr_slides[i].getAllFonts) {
						arr_slides[i].fontMap = {};
						arr_slides[i].getAllFonts(arr_slides[i].fontMap);
					}

					slideCopyObjects[i] = arr_slides[i];
				}

				arr_Images = arr_Images.concat(loader.End_UseFullUrl());

				presentationSelectedContent.SlideObjects = slideCopyObjects;
			};


			var readLayouts = function () {
				var loader = new AscCommon.BinaryPPTYLoader();
				if (!(bDuplicate === true)) {
					loader.Start_UseFullUrl();
				}
				loader.stream = stream;
				loader.presentation = editor.WordControl.m_oLogicDocument;
				loader.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;

				var selected_layouts = stream.GetULong();

				var layouts = [];
				for (var i = 0; i < selected_layouts; ++i) {
					if (PasteElementsId.g_bIsDocumentCopyPaste) {
						loader.stream.GetUChar();
						loader.stream.SkipRecord();
					} else {
						let oLayout = loader.ReadSlideLayout();
						layouts.push(oLayout);
						oLayout.getAllFonts(oFontMap);
					}
				}

				arr_Images = arr_Images.concat(loader.End_UseFullUrl());
				presentationSelectedContent.Layouts = layouts;
			};

			var readMasters = function () {
				var loader = new AscCommon.BinaryPPTYLoader();
				if (!(bDuplicate === true)) {
					loader.Start_UseFullUrl();
				}
				loader.stream = stream;
				loader.presentation = editor.WordControl.m_oLogicDocument;
				loader.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;

				var count = stream.GetULong();

				var array = [];
				for (var i = 0; i < count; ++i) {
					if (PasteElementsId.g_bIsDocumentCopyPaste) {
						loader.stream.GetUChar();
						loader.stream.SkipRecord();
					} else {
						let oMaster = loader.ReadSlideMaster();
						array.push(oMaster);
						oMaster.getAllFonts(oFontMap);
					}
				}

				arr_Images = arr_Images.concat(loader.End_UseFullUrl());
				presentationSelectedContent.Masters = array;
			};

			var readNotes = function () {
				var loader = new AscCommon.BinaryPPTYLoader();
				loader.stream = stream;
				loader.presentation = editor.WordControl.m_oLogicDocument;
				loader.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;

				var selected_notes = stream.GetULong();

				var notes = [];
				for (var i = 0; i < selected_notes; ++i) {
					if (PasteElementsId.g_bIsDocumentCopyPaste) {
						loader.stream.GetUChar();
						loader.stream.SkipRecord();
					} else {
						let oNotes = loader.ReadNote();
						notes.push(oNotes);
						oNotes.getAllFonts(oFontMap);
					}
				}

				presentationSelectedContent.Notes = notes;
			};

			var readNotesMasters = function () {
				var loader = new AscCommon.BinaryPPTYLoader();
				loader.stream = stream;
				loader.presentation = editor.WordControl.m_oLogicDocument;
				loader.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;

				var count = stream.GetULong();

				var array = [];
				for (var i = 0; i < count; ++i) {
					if (PasteElementsId.g_bIsDocumentCopyPaste) {
						loader.stream.GetUChar();
						loader.stream.SkipRecord();
					} else {
						let oNotesMaster = loader.ReadNoteMaster();
						array.push(oNotesMaster);
						oNotesMaster.getAllFonts(oFontMap);
					}
				}

				presentationSelectedContent.NotesMasters = array;
			};

			var readNotesThemes = function () {
				var loader = new AscCommon.BinaryPPTYLoader();
				loader.stream = stream;
				loader.presentation = editor.WordControl.m_oLogicDocument;
				loader.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;

				var count = stream.GetULong();

				//TODO возможно стоит пропустить при чтении в документах
				var array = [];
				for (var i = 0; i < count; ++i) {
					let oTheme = loader.ReadTheme();
					array.push(oTheme);
					oTheme.Document_Get_AllFontNames(oFontMap);
				}

				presentationSelectedContent.NotesThemes = array;
			};

			var readThemes = function () {
				var loader = new AscCommon.BinaryPPTYLoader();
				if (!(bDuplicate === true)) {
					loader.Start_UseFullUrl();
				}
				loader.stream = stream;
				loader.presentation = editor.WordControl.m_oLogicDocument;
				loader.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;

				var count = stream.GetULong();

				//TODO возможно стоит пропустить при чтении в документах
				var array = [];
				for (var i = 0; i < count; ++i) {
					let oTheme = loader.ReadTheme();
					oTheme.Document_Get_AllFontNames(oFontMap);
					array.push(oTheme);
				}

				arr_Images = arr_Images.concat(loader.End_UseFullUrl());
				presentationSelectedContent.Themes = array;
			};

			var readIndexes = function () {
				var count = stream.GetULong();

				var array = [];
				for (var i = 0; i < count; ++i) {
					array.push(stream.GetULong());
				}

				return array;
			};

			var bIsEmptyContent = true;
			var first_content = stream.GetString2();
			if (first_content === "SelectedContent") {
				var PresentationWidth = stream.GetULong() / 100000.0;
				var PresentationHeight = stream.GetULong() / 100000.0;
				var countContent = stream.GetULong();
				for (var i = 0; i < countContent; i++) {
					if (null === presentationSelectedContent) {
						presentationSelectedContent = typeof PresentationSelectedContent !== "undefined" ? new PresentationSelectedContent() : {};
						presentationSelectedContent.PresentationWidth = PresentationWidth;
						presentationSelectedContent.PresentationHeight = PresentationHeight;
					}
					var first_string = stream.GetString2();
					if ("DocContent" !== first_string) {
						bIsEmptyContent = false;
					}

					switch (first_string) {
						case "DocContent": {
							readContent();
							break;
						}
						case "Drawings": {
							readDrawings();
							break;
						}
						case "SlideObjects": {
							readSlideObjects();
							break;
						}
						case "Layouts": {
							History.TurnOff();
							readLayouts();
							History.TurnOn();
							break;
						}
						case "LayoutsIndexes": {
							presentationSelectedContent.LayoutsIndexes = readIndexes();
							break;
						}
						case "Masters": {
							readMasters();
							break;
						}
						case "MastersIndexes": {
							presentationSelectedContent.MastersIndexes = readIndexes();
							break;
						}
						case "Notes": {
							readNotes();
							break;
						}
						case "NotesMasters": {
							History.TurnOff();
							readNotesMasters();
							History.TurnOn();
							break;
						}
						case "NotesMastersIndexes": {
							presentationSelectedContent.NotesMastersIndexes = readIndexes();
							break;
						}
						case "NotesThemes": {
							readNotesThemes();
							break;
						}
						case "Themes": {
							readThemes();
							break;
						}
						case "ThemeIndexes": {
							presentationSelectedContent.ThemesIndexes = readIndexes();
							break;
						}
					}
				}
			}

			if (presentationSelectedContent) {
				var aSlides = presentationSelectedContent.SlideObjects;
				if (Array.isArray(aSlides)) {
					for (var i = 0; i < aSlides.length; ++i) {
						var oCurSlide = aSlides[i];
						var oSlideFontMap = oCurSlide ? oCurSlide.fontMap : null;
						if (oSlideFontMap) {
							var oTheme = null;
							if (Array.isArray(presentationSelectedContent.LayoutsIndexes)) {
								var nLayoutIndex = presentationSelectedContent.LayoutsIndexes[i];
								if (AscFormat.isRealNumber(nLayoutIndex)) {
									if (Array.isArray(presentationSelectedContent.MastersIndexes)) {
										var nMasterIndex = presentationSelectedContent.MastersIndexes[nLayoutIndex];
										if (AscFormat.isRealNumber(nMasterIndex)) {
											if (Array.isArray(presentationSelectedContent.ThemesIndexes)) {
												var nThemeIndex = presentationSelectedContent.ThemesIndexes[nMasterIndex];
												if (AscFormat.isRealNumber(nThemeIndex)) {
													if (Array.isArray(presentationSelectedContent.Themes)) {
														oTheme = presentationSelectedContent.Themes[nThemeIndex];
													}
												}
											}
										}
									}
								}
							}
							if (oTheme) {
								AscFormat.checkThemeFonts(oSlideFontMap, oTheme.themeElements.fontScheme);
							}
							for (var key in oSlideFontMap) {
								if (oSlideFontMap.hasOwnProperty(key)) {
									oFontMap[key] = 1;
								}
							}
							AscFormat.checkThemeFonts(oFontMap, {});
							delete oCurSlide.fontMap;
						}
					}
				}
				for (var key in oFontMap) {
					fonts.push(new CFont(key));
				}
			}
			if (bIsEmptyContent) {
				presentationSelectedContent = null;
			}

			return {content: presentationSelectedContent, fonts: fonts, images: arr_Images};
		}, this, []);
	},

	_pasteFromHtml: function (node, bTurnOffTrackRevisions) {
		let oThis = this;
		let isPasteTextIntoList = !!this.pasteTextIntoList;
		let oAPI = oThis.api;

		let fPasteHtmlPresentationCallback = function (fonts, images) {
			oThis.aContent = [];
			let aShapes = [], aImages = [], aTables = [];
			let aCopyObjects = [];
			let oPresentation = oAPI.WordControl.m_oLogicDocument;
			let fExecutePastePresentation = function () {
				//prepare content

				//remove single shape with empty content
				if (aShapes.length === 1) {
					let oFirstShape = aShapes[0];
					let oTxBody = oFirstShape.txBody;
					if(oTxBody) {
						let oDocContent = oTxBody.content;
						if(!oDocContent || oDocContent.IsEmpty()) {
							aShapes.length = 0;
						}
					}
				}

				for (let nSp = 0; nSp < aShapes.length; ++nSp) {
					let oSp = aShapes[nSp];
					let oTxBody = oSp.txBody;
					let oTxContent = oTxBody.content;
					let aTxContent = oTxContent.Content;
					if (aTxContent.length > 1) {
						let oFirstElement = aTxContent[0];
						if(oFirstElement.IsEmpty()) {
							oTxContent.Internal_Content_Remove(0, 1);
						}
					}
					let dWidth, dHeight;
					dWidth = oTxBody.getRectWidth(oPresentation.GetWidthMM() * 2 / 3);
					dHeight = oTxContent.GetSummaryHeight();
					AscFormat.CheckSpPrXfrm(oSp);
					let oXfrm = oSp.spPr.xfrm;
					oXfrm.setExtX(dWidth);
					oXfrm.setExtY(dHeight);
					oXfrm.setOffX((oPresentation.GetWidthMM() - dWidth) / 2.0);
					oXfrm.setOffY((oPresentation.GetHeightMM() - dHeight) / 2.0);
					let oBodyPr = oSp.getBodyPr().createDuplicate();
					oBodyPr.rot = 0;
					oBodyPr.spcFirstLastPara = false;
					oBodyPr.vertOverflow = AscFormat.nVOTOverflow;
					oBodyPr.horzOverflow = AscFormat.nHOTOverflow;
					oBodyPr.vert = AscFormat.nVertTThorz;
					oBodyPr.setDefaultInsets();
					oBodyPr.numCol = 1;
					oBodyPr.spcCol = 0;
					oBodyPr.rtlCol = 0;
					oBodyPr.fromWordArt = false;
					oBodyPr.anchor = 4;
					oBodyPr.anchorCtr = false;
					oBodyPr.forceAA = false;
					oBodyPr.compatLnSpc = true;
					oBodyPr.prstTxWarp = AscFormat.CreatePrstTxWarpGeometry("textNoShape");
					oBodyPr.textFit = new AscFormat.CTextFit();
					oBodyPr.textFit.type = AscFormat.text_fit_Auto;
					oSp.txBody.setBodyPr(oBodyPr);
					oSp.txBody.content.MoveCursorToEndPos();
					aCopyObjects.push(new DrawingCopyObject(oSp, 0, 0, dWidth, dHeight));
				}

				for (let nTable = 0; nTable < aTables.length; ++nTable) {
					let oTableFrame = aTables[nTable];
					//Todo: to think about width and height of inserted graphic frame with table
					let dWidth = 100;
					let dHeight = 100;
					//----------------
					AscFormat.CheckSpPrXfrm(oTableFrame);
					let oXfrm = oTableFrame.spPr.xfrm;
					oXfrm.setExtX(dWidth);
					oXfrm.setExtY(dHeight);
					oXfrm.setOffX(0);
					oXfrm.setOffY(0);
					aCopyObjects.push(new DrawingCopyObject(oTableFrame, 0, 0, dWidth, dHeight));
				}

				for (let nImage = 0; nImage < aImages.length; ++nImage) {
					let oImage = aImages[nImage];
					AscFormat.CheckSpPrXfrm(oImage);
					let oXfrm = oImage.spPr.xfrm;
					let dWidth = oXfrm.extX;
					let dHeight = oXfrm.extY;
					oXfrm.setOffX(0);
					oXfrm.setOffY(0);
					aCopyObjects.push(new DrawingCopyObject(oImage, 0, 0, dWidth, dHeight));
				}


				//INSERT CONTENT
				let oController = oPresentation.GetCurrentController();
				let oTargetContent = oController && oController.getTargetDocContent();
				if (oTargetContent && aCopyObjects.length === 1 && aImages.length === 0 && aTables.length === 0) {
					//не проверяем на лок т. к. это делается в asc_docs_api.prototype.asc_PasteData.
					// При двух последовательных проверках в совместном редактировании вторая проверка всегда будет возвращать лок
					let aNewContent = aCopyObjects[0].Drawing.txBody.content.Content;
					oThis.InsertInPlacePresentation(aNewContent);
				} else {
					let oSelectedContent = new PresentationSelectedContent();
					oSelectedContent.Drawings = aCopyObjects;
					let bPaste = oPresentation.InsertContent(oSelectedContent);
					oPresentation.Recalculate();
					oPresentation.UpdateInterface();
					oAPI.checkChangesSize();

					//check only images
					let bOnlyImg = false;
					if (oSelectedContent.isDrawingsContent()) {
						bOnlyImg = true;
						let aDrawingCopyObjects = oSelectedContent.Drawings;
						for (let nDrawing = 0; nDrawing < aDrawingCopyObjects.length; nDrawing++) {
							let oSp = aDrawingCopyObjects[nDrawing].Drawing;
							if (oSp.getObjectType() !== AscDFH.historyitem_type_ImageShape) {
								bOnlyImg = false;
								break;
							}
						}
					}
					let oPasteHelper = window['AscCommon'].g_specialPasteHelper;
					if (bOnlyImg || !bPaste) {
						oPasteHelper.SpecialPasteButton_Hide();
						if (oPasteHelper.buttonInfo) {
							oPasteHelper.showButtonIdParagraph = null;
							oPasteHelper.CleanButtonInfo();
						}
					} else {
						oThis._setSpecialPasteShowOptionsPresentation([Asc.c_oSpecialPasteProps.sourceformatting, Asc.c_oSpecialPasteProps.keepTextOnly]);
					}
					oPasteHelper.Paste_Process_End();
				}
			};
			if (oPresentation.GetSlidesCount() === 0) {
				oPresentation.addNextSlide();
			}
			let oShape = new CShape();
			oShape.setParent(oPresentation.GetCurrentSlide());
			oShape.setTxBody(AscFormat.CreateTextBodyFromString("", oPresentation.DrawingDocument, oShape));
			aShapes.push(oShape);

			oThis._Execute(node, {}, true, true, false, aShapes, aImages, aTables);

			if (!fonts) {
				fonts = [];
			}
			if (!images) {
				images = [];
			}
			oAPI.pre_Paste(fonts, images, fExecutePastePresentation);
		};

		var fPasteHtmlWordCallback = function (fonts, images) {
			var executePasteWord = function () {

				if (false === oThis.bNested) {
					if (oThis.aNeedRecalcImgSize) {
						for (var i = 0; i < oThis.aNeedRecalcImgSize.length; i++) {
							if (PasteElementsId.g_bIsDocumentCopyPaste) {
								var drawing = oThis.aNeedRecalcImgSize[i].drawing;
								var img = oThis.aNeedRecalcImgSize[i].img;
								if (drawing && img) {
									var imgSize = oThis._getImgSize(img);
									let fitPagePictureSize = oThis.fitPictureSizeToPage(imgSize.width, imgSize.height);
									let nWidth = fitPagePictureSize.nWidth;
									let nHeight = fitPagePictureSize.nHeight;

									if (imgSize && drawing.Extent && (drawing.Extent.H !== nHeight || drawing.Extent.W !== nWidth)) {
										drawing.setExtent(nWidth, nHeight);
										drawing.GraphicObj.spPr.xfrm.setExtX(nWidth);
										drawing.GraphicObj.spPr.xfrm.setExtY(nHeight);
									}
								}
							} else {
							}
						}
					}
					oThis.InsertInDocument();
				}
				if (false !== bTurnOffTrackRevisions) {
					oThis.api.WordControl.m_oLogicDocument.SetLocalTrackRevisions(bTurnOffTrackRevisions);
				}
				if (false === oThis.bNested) {
					if (oThis.pasteCallback) {
						oThis.pasteCallback();
					}
				}
			};

			//если в итоге во вставляемом контенте нет следов списков, тогда вставляем просто текст, в противном случае - чистим списки
			if (oThis.pasteTextIntoList) {
				oThis._pasteText(oThis.pasteTextIntoList);
				return;
			}

			oThis.aContent = [];
			//если находимся внутри текстовой области диаграммы, то не вставляем ссылки
			if (oThis.oDocument && oThis.oDocument.Parent && oThis.oDocument.Parent.parent && oThis.oDocument.Parent.parent.parent && oThis.oDocument.Parent.parent.parent.getObjectType && oThis.oDocument.Parent.parent.parent.getObjectType() == AscDFH.historyitem_type_Chart) {
				var hyperlinks = node.getElementsByTagName("a");
				if (hyperlinks && hyperlinks.length) {
					var newElement;
					for (var i = 0; i < hyperlinks.length; i++) {
						newElement = document.createElement("span");

						var cssText = hyperlinks[i].getAttribute('style');
						if (cssText)
							newElement.getAttribute('style', cssText);

						$(newElement).append(hyperlinks[i].children);

						hyperlinks[i].parentNode.replaceChild(newElement, hyperlinks[i]);
					}
				}

				//Todo пока сделал так, чтобы не вставлялись графические объекты в название диаграммы, потом нужно будет сделать так же запутанно, как в MS
				var htmlImages = node.getElementsByTagName("img");
				if (htmlImages && htmlImages.length) {
					for (var i = 0; i < htmlImages.length; i++) {
						htmlImages[i].parentNode.removeChild(htmlImages[i]);
					}
				}
			}
			if (false !== bTurnOffTrackRevisions) {
				oThis.api.WordControl.m_oLogicDocument.SetLocalTrackRevisions(bTurnOffTrackRevisions);
			}
			//на время заполнения контента для вставки отключаем историю
			oThis._Execute(node, {}, true, true, false);
			//by section
			//oThis._applyMsoSections(oThis.aContent, oThis.oMsoSections);
			oThis._AddNextPrevToContent(oThis.oDocument);

			if (isPasteTextIntoList) {
				oThis._pasteText(oThis._getTextFromContent(oThis.aContent, {NewLineParagraph: true, Numbering: false}));
				return;
			}

			oThis.api.pre_Paste(fonts, images, executePasteWord);
		};

		this.oRootNode = node;

		if (AscCommon.g_clipboardBase.pastedFrom === AscCommon.c_oClipboardPastedFrom.Word) {
			this.aMsoHeadStylesStr = this._findMsoHeadStyle(node.parentElement);
		}

		if (PasteElementsId.g_bIsDocumentCopyPaste) {
			this.bIsPlainText = this._CheckIsPlainText(node);

			var specialPasteOnlyText = Asc.c_oSpecialPasteProps.keepTextOnly === window['AscCommon'].g_specialPasteHelper.specialPasteProps;
			if (specialPasteOnlyText && !this.pasteTextIntoList) {
				fPasteHtmlWordCallback();
			} else {
				this._Prepeare(node, fPasteHtmlWordCallback);
			}

			if (false !== bTurnOffTrackRevisions) {
				oThis.api.WordControl.m_oLogicDocument.SetLocalTrackRevisions(bTurnOffTrackRevisions);
			}
		} else {
			this.oRootNode = node;
			if (window['AscCommon'].g_specialPasteHelper.specialPasteStart && Asc.c_oSpecialPasteProps.keepTextOnly === window['AscCommon'].g_specialPasteHelper.specialPasteProps) {
				fPasteHtmlPresentationCallback();
			} else {
				this._Prepeare(node, fPasteHtmlPresentationCallback);
			}
		}
	},

	_getClassBinaryFromHtml: function (node, onlyBinary) {
		var classNode, base64FromExcel = null, base64FromWord = null, base64FromPresentation = null;

		if (onlyBinary) {
			if (typeof onlyBinary === "object") {
				var prefix = String.fromCharCode(onlyBinary[0], onlyBinary[1], onlyBinary[2], onlyBinary[3]);

				if ("PPTY" === prefix) {
					base64FromPresentation = onlyBinary;
				} else if ("DOCY" === prefix) {
					base64FromWord = onlyBinary;
				} else if ("XLSY" === prefix) {
					base64FromExcel = onlyBinary;
				}
			} else {
				if (onlyBinary.indexOf("pptData;") > -1) {
					base64FromPresentation = onlyBinary.split('pptData;')[1];
				} else if (onlyBinary.indexOf("docData;") > -1) {
					base64FromWord = onlyBinary.split('docData;')[1];
				} else if (onlyBinary.indexOf("xslData;") > -1) {
					base64FromExcel = onlyBinary.split('xslData;')[1];
				}
			}
		} else if (node) {
			classNode = searchBinaryClass(node);

			if (classNode != null) {
				var cL = classNode.split(" ");
				for (var i = 0; i < cL.length; i++) {
					if (cL[i].indexOf("xslData;") > -1) {
						base64FromExcel = cL[i].split('xslData;')[1];
					} else if (cL[i].indexOf("docData;") > -1) {
						base64FromWord = cL[i].split('docData;')[1];
					} else if (cL[i].indexOf("pptData;") > -1) {
						base64FromPresentation = cL[i].split('pptData;')[1];
					}
				}
			}
		}

		return {
			base64FromExcel: base64FromExcel,
			base64FromWord: base64FromWord,
			base64FromPresentation: base64FromPresentation
		};
	},

	_pasteText: function (text) {
		var oThis = this;

		var fPasteTextWordCallback = function () {
			var executePasteWord = function () {
				if (false === oThis.bNested) {
					oThis.InsertInDocument(!window['AscCommon'].g_specialPasteHelper.specialPasteStart);
				}
			};

			oThis.aContent = [];
			oThis._getContentFromText(text, true);
			oThis._AddNextPrevToContent(oThis.oDocument);

			oThis.api.pre_Paste([], [], executePasteWord);
		};

		var fPasteTextPresentationCallback = function () {
			var executePastePresentation = function () {
				oThis.InsertInPlacePresentation(oThis.aContent, true);
			};

			var presentation = editor.WordControl.m_oLogicDocument;
			if (presentation.Slides.length === 0) {
				presentation.addNextSlide();
			}
			var shape = new CShape();
			shape.setParent(presentation.Slides[presentation.CurPage]);
			shape.setTxBody(AscFormat.CreateTextBodyFromString("", presentation.DrawingDocument, shape));
			oThis.aContent = shape.txBody.content.Content;

			text = text.replace(/^(\r|\t)+|(\r|\t)+$/g, '');
			if (text.length > 0) {
                //TODO: May be use CDocumentContent.AddText instead
                var oContent = shape.txBody.content;
				oThis.oDocument = oContent;
				var bAddParagraph = false;
                var oCurParagraph = oContent.Content[0];
                var oCurRun = new ParaRun(oCurParagraph, false);
                var nCharPos = 0;
                oCurParagraph.Internal_Content_Add(0, oCurRun);
				for (var oIterator = text.getUnicodeIterator(); oIterator.check(); oIterator.next()) {
					if (bAddParagraph) {
                        oCurParagraph = new Paragraph(oContent.DrawingDocument, oContent, oContent.bPresentation === true);
                        oContent.Internal_Content_Add(oContent.Content.length, oCurParagraph);
                        oCurRun = new ParaRun(oCurParagraph, false);
                        oCurParagraph.Internal_Content_Add(0, oCurRun);
						bAddParagraph = false;
                        nCharPos = 0;
					}
					var nUnicode = oIterator.value();
                    if(null !== nUnicode) {
                        if (null !== nUnicode && 13 !== nUnicode) {
                            if (0x0A === nUnicode || 0x0D === nUnicode) {
                                bAddParagraph = true;
                            }
                            else if (9 === nUnicode) // \t
                                oCurRun.AddToContent(nCharPos++, new AscWord.CRunTab(), true);
                            else if (10 === nUnicode) // \n
                                oCurRun.AddToContent(nCharPos++, new AscWord.CRunBreak(AscWord.break_Line), true);
                            else if (13 === nUnicode) // \r
                                continue;
                            else if (AscCommon.IsSpace(nUnicode)) // space
                                oCurRun.AddToContent(nCharPos++, new AscWord.CRunSpace(nUnicode), true);
                            else
                                oCurRun.AddToContent(nCharPos++, new AscWord.CRunText(nUnicode), true);
                        }
                    }
				}
			}

			var oTextPr = presentation.GetCalculatedTextPr();
			shape.txBody.content.SetApplyToAll(true);
			var paraTextPr = new AscCommonWord.ParaTextPr(oTextPr);
			shape.txBody.content.AddToParagraph(paraTextPr);
			shape.txBody.content.SetApplyToAll(false);

			oThis.api.pre_Paste([], [], executePastePresentation);
		};

		if (PasteElementsId.g_bIsDocumentCopyPaste) {
			fPasteTextWordCallback();
		} else {
			fPasteTextPresentationCallback();
		}
	},

	_getContentFromText: function (text, getStyleCurSelection) {
		var t = this;
		var Count = text.length;

		var pasteIntoParagraphPr = this.oDocument.GetDirectParaPr();
		var pasteIntoParaRunPr = this.oDocument.GetDirectTextPr();
		var bPresentation = false;
		var oCurParagraph = this.oDocument.GetCurrentParagraph();
		var Parent = t.oDocument;
		if (oCurParagraph && !oCurParagraph.bFromDocument) {
			bPresentation = true;
			Parent = oCurParagraph.Parent;
		}

		var getNewParagraph = function () {
			var paragraph = new Paragraph(t.oDocument.DrawingDocument, Parent, bPresentation);
			var copyParaPr;
			if (getStyleCurSelection) {
				if (pasteIntoParagraphPr) {
					copyParaPr = pasteIntoParagraphPr.Copy();
					copyParaPr.NumPr = undefined;
					paragraph.Set_Pr(copyParaPr);

					if (paragraph.TextPr && pasteIntoParaRunPr) {
						paragraph.TextPr.Value = pasteIntoParaRunPr.Copy();
					}
				}
			}
			return paragraph;
		};

		var getNewParaRun = function () {
			var paraRun = new ParaRun();
			if (getStyleCurSelection) {
				if (pasteIntoParaRunPr && paraRun.Set_Pr) {
					paraRun.Set_Pr(pasteIntoParaRunPr.Copy());
				}
			}
			return paraRun;
		};


		var _addToRun = function (_nUnicode) {
			var Item;
			if (0x2009 === _nUnicode || 9 === _nUnicode) {
				Item = new AscWord.CRunTab();
			} else if (0x20 !== _nUnicode && 0xA0 !== _nUnicode) {
				Item = new AscWord.CRunText(_nUnicode);
			} else {
				Item = new AscWord.CRunSpace();
			}

			//add text
			newParaRun.AddToContent(-1, Item, false);
		};

		var newParagraph = getNewParagraph();
		var partTextCount = 0;
		var newParaRun = getNewParaRun();
		for (var oIterator = text.getUnicodeIterator(); oIterator.check(); oIterator.next()) {
			var pos = oIterator.position();
			var nUnicode = oIterator.value();

			if (0x0A === nUnicode || pos === Count - 1) {
				if (pos === Count - 1 && 0x0A !== nUnicode) {
					_addToRun(nUnicode);
				}

				newParagraph.Internal_Content_Add(newParagraph.Content.length - 1, newParaRun, false);
				this.aContent.push(newParagraph);

				newParagraph = getNewParagraph();
				newParaRun = getNewParaRun();
				partTextCount = 0;
			} else if (partTextCount === Asc.c_dMaxParaRunContentLength) {//max run length
				_addToRun(nUnicode);

				newParagraph.Internal_Content_Add(newParagraph.Content.length - 1, newParaRun, false);
				newParaRun = getNewParaRun();
				partTextCount = 0;
			} else if (13 === nUnicode) {
				continue;
			} else {
				partTextCount++;
				_addToRun(nUnicode);
			}
		}

		if (partTextCount) {
			newParagraph.Internal_Content_Add(newParagraph.Content.length - 1, newParaRun, false);
			this.aContent.push(newParagraph);
		}
	},

	_getTextFromContent: function (aContent, oPr) {
		var ResultText = "";
		if (aContent) {
			for (var Index = 0; Index < aContent.length; Index++)
			{
				aContent[Index].SelectAll();
				aContent[Index].ApplyToAll = true;
				ResultText += aContent[Index].GetSelectedText(false, oPr);
			}
		}
		return ResultText;
	},

	_isParagraphContainsOnlyDrawing: function (par) {
		var res = true;

		if (par.Content) {
			for (var i = 0; i < par.Content.length; i++) {
				if (par.Content[i] && par.Content[i].Content && par.Content[i].Content.length) {
					for (var j = 0; j < par.Content[i].Content.length; j++) {
						var elem = par.Content[i].Content[j];
						if (!(para_Drawing === elem.Get_Type() || para_End === elem.Get_Type())) {
							res = false;
						}
					}
				}
			}
		}

		return res;
	},

	_convertExcelBinary: function (aContentExcel, pDrawings) {
		//пока только распознаём только графические объекты
		var aContent = null, tempParagraph = null;
		var imageUrl, isGraphicFrame, extX, extY;
		var fonts = null;

		var drawings = pDrawings ? pDrawings : aContentExcel.workbook.aWorksheets[0].Drawings;
		if (drawings && drawings.length) {
			var drawing, graphicObj, paraRun, tempParaRun;

			aContent = [];

			var font_map = {};
			//из excel в word они вставляются в один параграф
			for (var i = 0; i < drawings.length; i++) {
				drawing = drawings[i] && drawings[i].Drawing ? drawings[i].Drawing : drawings[i];

				//TODO нужна отдельная обработка для таблиц из презентаций
				isGraphicFrame = typeof CTable !== "undefined" && drawing.graphicObject instanceof CTable;
				if (isGraphicFrame && drawings.length > 1 && drawings[i].base64)//если кроме таблички(при вставке из презентаций) содержатся ещё данные, вставляем в виде base64
				{
					if (!tempParagraph)
						tempParagraph = new Paragraph(this.oDocument.DrawingDocument, this.oDocument);

					extX = drawings[i].ExtX;
					extY = drawings[i].ExtY;
					imageUrl = drawings[i].base64;

					graphicObj = AscFormat.DrawingObjectsController.prototype.createImage(imageUrl, 0, 0, extX, extY);

					tempParaRun = new ParaRun();
					tempParaRun.Paragraph = null;
					tempParaRun.Add_ToContent(0, new ParaDrawing(), false);

					tempParaRun.Content[0].Set_GraphicObject(graphicObj);
					tempParaRun.Content[0].GraphicObj.setParent(tempParaRun.Content[0]);
					tempParaRun.Content[0].CheckWH();

					tempParagraph.Content.splice(tempParagraph.Content.length - 1, 0, tempParaRun);
				} else if (isGraphicFrame) {
					
					var copyObj = drawing.getWordTable();
					if(copyObj) {
						copyObj.Set_Inline && copyObj.Set_Inline(true);
						copyObj.Set_Parent(this.oDocument);
						aContent[aContent.length] = copyObj;
						drawing.getAllFonts(font_map);
					}

				} else {
					if (!tempParagraph)
						tempParagraph = new Paragraph(this.oDocument.DrawingDocument, this.oDocument);

					extX = drawings[i].ExtX;
					extY = drawings[i].ExtY;

					drawing.getAllFonts(font_map);
					graphicObj = drawing.graphicObject ? drawing.graphicObject.convertToWord(this.oLogicDocument) : drawing.convertToWord(this.oLogicDocument);

					tempParaRun = new ParaRun();
					tempParaRun.Paragraph = null;

					var newParaDrawing = new ParaDrawing();
					//newParaDrawing.Set_DrawingType(drawing_Anchor);
					tempParaRun.Add_ToContent(0, newParaDrawing, false);

					tempParaRun.Content[0].Set_GraphicObject(graphicObj);
					tempParaRun.Content[0].GraphicObj.setParent2(tempParaRun.Content[0]);
					var oGraphicObj = tempParaRun.Content[0].GraphicObj;
					if (oGraphicObj.spPr && oGraphicObj.spPr.xfrm) {
						oGraphicObj.spPr.xfrm.setOffX(0);
						oGraphicObj.spPr.xfrm.setOffY(0);
					}
					tempParaRun.Content[0].CheckWH();

					tempParagraph.Content.splice(tempParagraph.Content.length - 1, 0, tempParaRun);
				}
			}

			fonts = [];
			for (var i in font_map) {
				fonts.push(new CFont(i));
			}

			if (tempParagraph)
				aContent[aContent.length] = tempParagraph;
		} else {
			if (!this.oDocument.bPresentation) {
				fonts = this._convertTableFromExcel(aContentExcel);
				if (PasteElementsId.g_bIsDocumentCopyPaste && this.aContent && this.aContent.length === 1 && 1 === this.aContent[0].Rows && this.aContent[0].Content[0]) {
					var _content = this.aContent[0].Content[0];
					if (_content && _content.Content && 1 === _content.Content.length && _content.Content[0].Content &&
						_content.Content[0].Content.Content[0]) {
						this.aContent[0] = _content.Content[0].Content.Content[0];
					}
				}
			} else {
				fonts = this._convertTableFromExcelForChartTitle(aContentExcel);
			}
			aContent = this.aContent;
		}

		return {content: aContent, fonts: fonts};
	},

	_convertTableFromExcel: function (aContentExcel) {
		var worksheet = aContentExcel.workbook.aWorksheets[0];
		var range;
		var tempActiveRef = aContentExcel.activeRange;
		var activeRange = AscCommonExcel.g_oRangeCache.getAscRange(tempActiveRef);
		var t = this;
		var fonts = [];

		var charToMM = function (mcw) {
			var maxDigitWidth = 7;
			var px = Asc.floor(((256 * mcw + Asc.floor(128 / maxDigitWidth)) / 256) * maxDigitWidth);
			return px * g_dKoef_pix_to_mm;
		};

		var convertBorder = function (border) {
			var res = new CDocumentBorder();
			if (border.w) {
				res.Value = border_Single;
				res.Size = border.w * g_dKoef_pix_to_mm;
			}
			var bc = border.getColorOrDefault();
			res.Color = new CDocumentColor(bc.getR(), bc.getG(), bc.getB());

			return res;
		};

		var addFont = function (fontFamily) {
			if (!t.aContent.fonts) {
				t.aContent.fonts = [];
			}

			if (!t.oFonts) {
				t.oFonts = [];
			}

			if (!t.oFonts[fontFamily]) {
				t.oFonts[fontFamily] = {Name: fontFamily, Index: -1};
				fonts.push(new CFont(fontFamily));
			}
		};

		//grid
		var grid = [];
		var standardWidth = 9;
		for (var i = activeRange.c1; i <= activeRange.c2; i++) {
			if (worksheet.aCols[i]) {
				grid[i - activeRange.c1] = charToMM(worksheet.aCols[i].width);
			} else {
				grid[i - activeRange.c1] = charToMM(standardWidth);
			}
		}

		var table;
		if (editor.WordControl.m_oLogicDocument && editor.WordControl.m_oLogicDocument.Slides) {
			table = this._createNewPresentationTable(grid);
			table.Set_TableStyle(editor.WordControl.m_oLogicDocument.DefaultTableStyleId)
		} else {
			table = new CTable(this.oDocument.DrawingDocument, this.oDocument, true, 0, 0, grid);
		}

		this.aContent.push(table);

		var diffRow = activeRange.r2 - activeRange.r1;
		var diffCol = activeRange.c2 - activeRange.c1;
		for (var i = 0; i <= diffRow; i++) {
			var row = table.private_AddRow(table.Content.length, 0);
			var heightRowPt = worksheet.getRowHeight(i + activeRange.r1);
			row.SetHeight(heightRowPt * g_dKoef_pt_to_mm, linerule_AtLeast);

			for (var j = 0; j <= diffCol; j++) {
				if (!table.Selection.Data) {
					table.Selection.Data = [];
				}

				table.Selection.Data[table.Selection.Data.length] = {Cell: j, Row: i};

				range = worksheet.getCell3(i + activeRange.r1, j + activeRange.c1);
				var oCurCell = row.Add_Cell(row.Get_CellsCount(), row, null, false);

				table.CurCell = oCurCell;

				var isMergedCell = range.hasMerged();
				var gridSpan = 1;
				var vMerge = 1;
				if (isMergedCell) {
					gridSpan = isMergedCell.c2 - isMergedCell.c1 + 1;
					vMerge = isMergedCell.r2 - isMergedCell.r1 + 1;
				}

				//***cell property***
				//set grid
				var sumWidthGrid = 0;
				for (var l = 0; l < gridSpan; l++) {
					sumWidthGrid += grid[j + l];
				}
				oCurCell.Set_W(new CTableMeasurement(tblwidth_Mm, sumWidthGrid));

				//margins
				//left
				oCurCell.Set_Margins({W: 0, Type: tblwidth_Mm}, 3);
				//right
				oCurCell.Set_Margins({W: 0, Type: tblwidth_Mm}, 1);
				//top
				oCurCell.Set_Margins({W: 0, Type: tblwidth_Mm}, 0);
				//bottom
				oCurCell.Set_Margins({W: 0, Type: tblwidth_Mm}, 2);

				//background color
				var background_color = range.getFillColor();
				if (null != background_color) {
					var Shd = new CDocumentShd();
					Shd.Value = c_oAscShdClear;
					Shd.Color = Shd.Fill = new CDocumentColor(background_color.getR(), background_color.getG(), background_color.getB());
					oCurCell.Set_Shd(Shd);
				}

				//borders
				var borders = range.getBorderFull();
				//left
				var border = convertBorder(borders.l);
				if (null != border) {
					oCurCell.Set_Border(border, 3);
				}
				//top
				border = convertBorder(borders.t);
				if (null != border) {
					oCurCell.Set_Border(border, 0);
				}
				//right
				border = convertBorder(borders.r);
				if (null != border) {
					oCurCell.Set_Border(border, 1);
				}
				//bottom
				border = convertBorder(borders.b);
				if (null != border) {
					oCurCell.Set_Border(border, 2);
				}

				//merge
				oCurCell.Set_GridSpan(gridSpan);
				if (vMerge != 1) {
					if (isMergedCell.r1 === i + activeRange.r1) {
						oCurCell.SetVMerge(vmerge_Restart);
					} else {
						oCurCell.SetVMerge(vmerge_Continue);
					}
				}

				var align = range.getAlign();
				var textAngle = align ? align.getAngle() : null;
				if (textAngle === 90) {
					oCurCell.Set_TextDirection(c_oAscCellTextDirection.BTLR);
				} else if (textAngle === -90) {
					oCurCell.Set_TextDirection(c_oAscCellTextDirection.TBRL);
				}

				var vAlign = align ? align.getAlignVertical() : null;
				switch (vAlign) {
					case Asc.c_oAscVAlign.Bottom:
						oCurCell.Set_VAlign(vertalignjc_Bottom);
						break;
					case Asc.c_oAscVAlign.Center:
					case Asc.c_oAscVAlign.Dist:
					case Asc.c_oAscVAlign.Just:
						oCurCell.Set_VAlign(vertalignjc_Center);
						break;
					case Asc.c_oAscVAlign.Top:
						oCurCell.Set_VAlign(vertalignjc_Top);
						break;
				}


				var oCurPar = oCurCell.GetContent().GetElement(0);
				if (!oCurPar || !oCurPar.IsParagraph())
					continue;
				
				if (align) {
					var type = range.getType();
					if (null != align.hor) {
						oCurPar.SetParagraphAlign(align.hor);
					} else if (AscCommon.CellValueType.Number === type) {
						//для пустого текста, даже если тип соответсвующий, не проставляю выравнивание
						if (!range.isEmptyTextString()) {
							oCurPar.SetParagraphAlign(AscCommon.align_Right);
						}
					}
				}

				var hyperLink = range.getHyperlink();
				var oCurHyperlink = null;
				if (hyperLink) {
					oCurHyperlink = new ParaHyperlink();
					oCurHyperlink.SetValue(hyperLink.Hyperlink);
					if (hyperLink.Tooltip) {
						oCurHyperlink.SetToolTip(hyperLink.Tooltip);
					}
				}

				var value2 = range.getValue2();
				for (var n = 0; n < value2.length; n++) {
					const oCurRun = new AscWord.CRun(oCurPar);
					var format = value2[n].format;

					//***text property***
					oCurRun.SetBold(format.getBold());
					var fc = format.getColor();
					if (fc) {
						oCurRun.SetColor(new CDocumentColor(fc.getR(), fc.getG(), fc.getB()));
					}

					//font
					var font_family = format.getName();
					addFont(font_family);
					oCurRun.Pr.FontFamily = font_family;
					var oFontItem = this.oFonts[font_family];
					if (null != oFontItem && null != oFontItem.Name) {
						oCurRun.SetRFontsAscii({Name: oFontItem.Name, Index: oFontItem.Index});
						oCurRun.SetRFontsHAnsi({Name: oFontItem.Name, Index: oFontItem.Index});
						oCurRun.SetRFontsCS({Name: oFontItem.Name, Index: oFontItem.Index});
						oCurRun.SetRFontsEastAsia({Name: oFontItem.Name, Index: oFontItem.Index});
					}

					oCurRun.SetFontSize(format.getSize());
					oCurRun.SetItalic(format.getItalic());
					oCurRun.SetStrikeout(format.getStrikeout());
					oCurRun.SetUnderline(format.getUnderline() !== 2);
					oCurRun.SetVertAlign(format.va);

					//text
					if (true === format.skip || true === format.repeat) {
						oCurRun.AddToContent(-1, new AscWord.CRunSpace(), false);
					} else {
						AscWord.TextToRunElements(value2[n].text, function(runElement) {
							oCurRun.AddToContent(-1, runElement, false);
						});
					}

					//add run
					if (oCurHyperlink) {
						oCurHyperlink.AddToContent(n, oCurRun, false);
					} else {
						oCurPar.AddToContent(n, oCurRun, false);
					}
				}

				if (oCurHyperlink) {
					oCurPar.AddToContent(n, oCurHyperlink, false);
				}

				j = j + gridSpan - 1;
			}
		}

		return fonts;
	},

	_convertTableFromExcelForChartTitle: function (aContentExcel) {
		var worksheet = aContentExcel.workbook.aWorksheets[0];
		var range;
		var tempActiveRef = aContentExcel.activeRange;
		var activeRange = AscCommonExcel.g_oRangeCache.getAscRange(tempActiveRef);
		var t = this;
		var fonts = [];

		var addFont = function (fontFamily) {
			if (!t.aContent.fonts) {
				t.aContent.fonts = [];
			}

			if (!t.oFonts) {
				t.oFonts = [];
			}

			if (!t.oFonts[fontFamily]) {
				t.oFonts[fontFamily] = {Name: fontFamily, Index: -1};
				fonts.push(new CFont(fontFamily));
			}
		};

		var paragraph = new Paragraph(this.oDocument.DrawingDocument, this.oDocument, true);
		this.aContent.push(paragraph);

		var diffRow = activeRange.r2 - activeRange.r1;
		var diffCol = activeRange.c2 - activeRange.c1;
		for (var i = 0; i <= diffRow; i++) {
			for (var j = 0; j <= diffCol; j++) {
				range = worksheet.getCell3(i + activeRange.r1, j + activeRange.c1);
				var isMergedCell = range.hasMerged();
				var gridSpan = 1;
				var vMerge = 1;
				if (isMergedCell) {
					gridSpan = isMergedCell.c2 - isMergedCell.c1 + 1;
					vMerge = isMergedCell.r2 - isMergedCell.r1 + 1;
				}

				var hyperLink = range.getHyperlink();
				var oCurHyperlink = null;
				if (hyperLink) {
					oCurHyperlink = new ParaHyperlink();
					oCurHyperlink.SetParagraph(paragraph);
					oCurHyperlink.Set_Value(hyperLink.Hyperlink);
					if (hyperLink.Tooltip) {
						oCurHyperlink.SetToolTip(hyperLink.Tooltip);
					}
				}

				var value2 = range.getValue2();
				for (var n = 0; n < value2.length; n++) {
					var oCurRun = new ParaRun(paragraph);
					var format = value2[n].format;

					//***text property***
					oCurRun.Pr.Bold = format.getBold();
					var fc = format.getColor();
					if (fc) {
						oCurRun.Set_Unifill(AscFormat.CreateUnfilFromRGB(fc.getR(), fc.getG(), fc.getB()), true);
					}

					//font
					var font_family = format.getName();
					addFont(font_family);
					var oFontItem = this.oFonts[font_family];
					if (null != oFontItem && null != oFontItem.Name) {
						oCurRun.SetRFontsAscii({Name: oFontItem.Name, Index: oFontItem.Index});
						oCurRun.SetRFontsHAnsi({Name: oFontItem.Name, Index: oFontItem.Index});
						oCurRun.SetRFontsCS({Name: oFontItem.Name, Index: oFontItem.Index});
						oCurRun.SetRFontsEastAsia({Name: oFontItem.Name, Index: oFontItem.Index});
					}

					oCurRun.SetFontSize(format.getSize());
					oCurRun.SetItalic(format.getItalic());
					oCurRun.SetStrikeout(format.getStrikeout());
					oCurRun.SetUnderline(format.getUnderline() !== 2);

					//text
					var value = value2[n].text;
					for (var oIterator = value.getUnicodeIterator(); oIterator.check(); oIterator.next()) {
						var nUnicode = oIterator.value();

						var Item;
						if (0x20 !== nUnicode && 0xA0 !== nUnicode && 0x2009 !== nUnicode) {
							Item = new AscWord.CRunText(nUnicode);
						} else {
							Item = new AscWord.CRunSpace();
						}

						//add text
						oCurRun.AddToContent(-1, Item, false);
					}

					if (i !== diffRow || j !== diffCol) {
						oCurRun.Add_ToContent(oIterator.position(), new AscWord.CRunSpace(), false);
					}

					//add run
					if (oCurHyperlink) {
						oCurHyperlink.Add_ToContent(n, oCurRun, false);
					} else {
						paragraph.Internal_Content_Add(paragraph.Content.length - 1, oCurRun, false);
					}
				}

				if (oCurHyperlink) {
					paragraph.Internal_Content_Add(paragraph.Content.length - 1, oCurHyperlink, false);
				}

				j = j + gridSpan - 1;
			}
		}

		return fonts;
	},

	_convertTableToPPTX: function (table, isFromWord) {
		//TODO пересмотреть обработку для вложенных таблиц(можно сделать так, как при копировании из документов в таблицы)
		var oTable = AscFormat.ExecuteNoHistory(function () {
			var allRows = [];
			this.maxTableCell = 0;

			if (isFromWord) {
				table = this._replaceInnerTables(table, allRows, true);
			}

			//convert internal paragraphs
			table.bPresentation = true;
			for (let nRow = 0; nRow < table.Content.length; nRow++) {
				let oRow = table.Content[nRow];
				let aCells = oRow.Content;
				for (let nCell = 0; nCell < aCells.length; nCell++) {
					let oDocContent = aCells[nCell].Content;
					oDocContent.bPresentation = true;
					let aElements = [].concat(oDocContent.Content);
					oDocContent.Content.length = 0;
					let nPos = 0;
					for(let nElement = 0; nElement < aElements.length; ++nElement) {
						let oElement = aElements[nElement];
						if(oElement instanceof AscCommonWord.Paragraph) {
							let oNewParagraph = AscFormat.ConvertParagraphToPPTX(oElement, null, null, true, false);
							oDocContent.Internal_Content_Add(nPos++, oNewParagraph, false);
						}
					}
					if(nPos === 0) {
						let oNewParagraph = new Paragraph(oDocContent.DrawingDocument, oDocContent, true);
						oDocContent.Internal_Content_Add(0, oNewParagraph, true);
					}
				}
			}
			table.SetTableLayout(tbllayout_Fixed);
			return table;
		}, this, []);
		return oTable;
	},

	_replaceInnerTables: function (table, allRows, isRoot) {
		//заменяем внутренние таблички
		for (var i = 0; i < table.Content.length; i++) {
			allRows[allRows.length] = table.Content[i];

			if (this.maxTableCell < table.Content[i].Content.length)
				this.maxTableCell = table.Content[i].Content.length;

			for (var j = 0; j < table.Content[i].Content.length; j++) {
				var cDocumentContent = table.Content[i].Content[j].Content;
				cDocumentContent.bPresentation = true;

				var k = 0;
				for (var n = 0; n < cDocumentContent.Content.length; n++) {
					//если нашли внутреннюю табличку
					if (cDocumentContent.Content[n] instanceof CTable) {
						this._replaceInnerTables(cDocumentContent.Content[n], allRows);
						cDocumentContent.Content.splice(n, 1);
						n--;
					}
				}
			}
		}

		//дополняем пустыми ячейками, строки, где ячеек меньше
		if (isRoot === true) {
			for (var row = 0; row < allRows.length; row++) {
				var cells = allRows[row].Content;
				var cellsLength = 0;
				for (var m = 0; m < cells.length; m++) {
					cellsLength += cells[m].GetGridSpan();
				}
				if (cellsLength < this.maxTableCell) {
					for (var cell = cellsLength; cell < this.maxTableCell; cell++) {
						allRows[row].Add_Cell(allRows[row].Get_CellsCount(), allRows[row], null, false);
					}
				}
			}

			table.Content = allRows;
			table.Rows = allRows.length;
		}

		return table;
	},

	_createNewPresentationTable: function (grid) {
		var presentation = editor.WordControl.m_oLogicDocument;
		var graphicFrame = new CGraphicFrame(presentation.Slides[presentation.CurPage]);
		var table = new CTable(this.oDocument.DrawingDocument, graphicFrame, false, 0, 0, grid, true);
		table.SetTableLayout(tbllayout_Fixed);
		graphicFrame.setGraphicObject(table);
		graphicFrame.setNvSpPr(new AscFormat.UniNvPr());

		return table;
	},

	_getImagesFromExcelShapes: function (aDrawings, aSpTree, aPastedImages, aUrls) {
		//пока только распознаём только графические объекты
		var sImageUrl, nDrawingsCount, oGraphicObj, bDrawings;
		if (Array.isArray(aDrawings)) {
			nDrawingsCount = aDrawings.length;
			bDrawings = true;
		} else if (Array.isArray(aSpTree)) {
			nDrawingsCount = aSpTree.length;
			bDrawings = false;
		} else {
			return;
		}
		for (var i = 0; i < nDrawingsCount; i++) {
			if (bDrawings) {
				oGraphicObj = aDrawings[i].graphicObject;
			} else {
				oGraphicObj = aSpTree[i];
			}
			if (oGraphicObj) {
				if (oGraphicObj.spPr) {
					if (oGraphicObj.spPr.Fill && oGraphicObj.spPr.Fill.fill && typeof oGraphicObj.spPr.Fill.fill.RasterImageId === "string" && oGraphicObj.spPr.Fill.fill.RasterImageId.length > 0) {
						sImageUrl = oGraphicObj.spPr.Fill.fill.RasterImageId;
						aPastedImages[aPastedImages.length] = new AscCommon.CBuilderImages(oGraphicObj.spPr.Fill.fill, sImageUrl, null, oGraphicObj.spPr, null);
						aUrls[aUrls.length] = sImageUrl;
					}
					if (oGraphicObj.spPr.ln && oGraphicObj.spPr.ln.Fill && oGraphicObj.spPr.ln.Fill.fill && typeof oGraphicObj.spPr.ln.Fill.fill.RasterImageId === "string" && oGraphicObj.spPr.ln.Fill.fill.RasterImageId.length > 0) {
						sImageUrl = oGraphicObj.spPr.ln.Fill.fill.RasterImageId;
						aPastedImages[aPastedImages.length] = new AscCommon.CBuilderImages(oGraphicObj.spPr.ln.Fill.fill, sImageUrl, null, oGraphicObj.spPr, oGraphicObj.spPr.ln.Fill.fill.RasterImageId);
						aUrls[aUrls.length] = sImageUrl;
					}
				}
				switch (oGraphicObj.getObjectType()) {
					case AscDFH.historyitem_type_ImageShape: {
						sImageUrl = oGraphicObj.getImageUrl();
						if (typeof sImageUrl === "string" && sImageUrl.length > 0) {
							aPastedImages[aPastedImages.length] = new AscCommon.CBuilderImages(oGraphicObj.blipFill, sImageUrl, oGraphicObj, null, null);
							aUrls[aUrls.length] = sImageUrl;
						}
						break;
					}
					case AscDFH.historyitem_type_Shape: {
						break;
					}
					case AscDFH.historyitem_type_ChartSpace: {
						break;
					}
					case AscDFH.historyitem_type_GroupShape: {
						this._getImagesFromExcelShapes(null, oGraphicObj.spTree, aPastedImages, aUrls);
						break;
					}
				}
			}
		}
	},

	_selectShapesBeforeInsert: function (aNewContent, oDoc) {
		var content, drawingObj, allDrawingObj = [];
		for (var i = 0; i < aNewContent.length; i++) {
			content = aNewContent[i];
			drawingObj = content.GetAllDrawingObjects();

			if (!drawingObj || (drawingObj && !drawingObj.length) || content.GetType() === type_Table) {
				allDrawingObj = null;
				break;
			}

			for (var n = 0; n < drawingObj.length; n++) {
				allDrawingObj[allDrawingObj.length] = drawingObj[n];
			}
		}

		if (allDrawingObj && allDrawingObj.length)
			this.oLogicDocument.SelectDrawings(allDrawingObj, oDoc);
	},

	_readFromBinaryExcel: function (base64) {
		var oBinaryFileReader = new AscCommonExcel.BinaryFileReader(true);
		var tempWorkbook = new AscCommonExcel.Workbook();
		tempWorkbook.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;
		tempWorkbook.theme = this.oDocument.theme ? this.oDocument.theme : this.oLogicDocument.theme;
		if (!tempWorkbook.theme && this.oLogicDocument.Get_Theme)
			tempWorkbook.theme = this.oLogicDocument.Get_Theme();

		Asc.getBinaryOtherTableGVar(tempWorkbook);

		pptx_content_loader.Start_UseFullUrl();
		pptx_content_loader.Reader.ClearConnectedObjects();
		oBinaryFileReader.Read(base64, tempWorkbook);
		pptx_content_loader.Reader.AssignConnectedObjects();

		if (!tempWorkbook.aWorksheets.length) {
			return null;
		}

		var arrImages;
		//если есть срез в контенте - вставляем только картинку
		var _sheet = tempWorkbook.aWorksheets[0];
		var pDrawings;
		if (_sheet && _sheet.aSlicers && _sheet.aSlicers.length) {
			if (tempWorkbook.Core && tempWorkbook.Core.subject) {
				var _str = tempWorkbook.Core.subject;
				var _parseStr = _str.split(";");
				var _width = _parseStr[0];
				var _height = _parseStr[1];
				var _base64 = _str.substring(_width.length + _height.length + 2);
				if (_base64 && _width && _height) {
					//var drawing = CreateImageFromBinary(_base64, parseFloat(_width), parseFloat(_height));
					var drawing = AscFormat.DrawingObjectsController.prototype.createImage(_base64, 0, 0, parseFloat(_width), parseFloat(_height));
					arrImages = [new AscCommon.CBuilderImages(new AscFormat.CBlipFill(), _base64, drawing, null, null)];
					var objectRender = new AscFormat.DrawingObjects();
					var oFlags = {
						from: false,
						to: false,
						pos: false,
						ext: false,
						editAs: window["AscCommon"].c_oAscCellAnchorType.cellanchorTwoCell
					};
					var oNewDrawing = objectRender.createDrawingObject();
					oNewDrawing.graphicObject = drawing;
					pDrawings = [oNewDrawing];
				}
			}
		} else {
			arrImages = pptx_content_loader.End_UseFullUrl();
		}

		return {
			workbook: tempWorkbook,
			activeRange: oBinaryFileReader.InitOpenManager && oBinaryFileReader.InitOpenManager.copyPasteObj && oBinaryFileReader.InitOpenManager.copyPasteObj.activeRange,
			arrImages: arrImages,
			pDrawings: pDrawings
		};
	},

	ReadPresentationText: function (stream, worksheet) {
		var loader = new AscCommon.BinaryPPTYLoader();
		loader.Start_UseFullUrl();
		loader.DrawingDocument = worksheet ? worksheet.getDrawingDocument() : editor.WordControl.m_oLogicDocument.DrawingDocument;
		loader.stream = stream;
		loader.presentation = worksheet ? worksheet.model : editor.WordControl.m_oLogicDocument;

		var cDocumentContent;
		var isExcel = !!worksheet;
		if (!isExcel) {
			var presentation = editor.WordControl.m_oLogicDocument;
			var shape;
			if (presentation.Slides) {
				shape = new CShape(presentation.Slides[presentation.CurPage]);
			} else {
				shape = new CShape(presentation);
			}

			shape.setTxBody(new AscFormat.CTextBody(shape));

			//читаем контент, здесь только параграфы
			cDocumentContent = new AscFormat.CDrawingDocContent(shape.txBody, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, false, false);
		} else {
			cDocumentContent = worksheet;
		}

		var elements = [], paragraph, selectedElement;
		var count = stream.GetULong() / 100000;
		for (var i = 0; i < count; ++i) {
			//loader.stream.Skip2(1); // must be 0
			var selectedAll = stream.GetUChar();
			paragraph = loader.ReadParagraph(cDocumentContent);

			if (isExcel) {
				elements.push(paragraph);
			} else {
				//FONTS
				paragraph.Document_Get_AllFontNames(this.oFonts);

				selectedElement = new AscCommonWord.CSelectedElement();
				selectedElement.Element = paragraph;
				if (selectedAll === 1) {
					selectedElement.SelectedAll = true;
				}
				elements.push(selectedElement);
			}
		}
		return elements;
	},

	ReadPresentationShapes: function (stream) {
		var loader = new AscCommon.BinaryPPTYLoader();
		loader.Start_UseFullUrl();
		loader.ClearConnectedObjects();
		pptx_content_loader.Reader.Start_UseFullUrl();

		loader.stream = stream;
		loader.presentation = editor.WordControl.m_oLogicDocument;
		loader.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;
		var presentation = editor.WordControl.m_oLogicDocument;
		var count = stream.GetULong();
		var arr_shapes = [];
		var arr_transforms = [];
		var cStyle;
		var foundTableStylesIdMap = {};

		for (var i = 0; i < count; ++i) {
			loader.TempMainObject = presentation && presentation.Slides ? presentation.Slides[presentation.CurPage] : presentation;
			var style_index = null;

			//читаем флаг о наличии табличного стиля
			if (!loader.stream.GetBool()) {
				if (loader.stream.GetBool()) {
					loader.stream.Skip2(1);
					cStyle = loader.ReadTableStyle(true);

					loader.stream.GetBool();
					style_index = stream.GetString2();
				}
			}

			var drawing = loader.ReadGraphicObject();

			//для случая, когда копируем 1 таблицу из презентаций, в бинарник заносим ещё одну такую же табличку, но со скомпилированными стилями(для вставки в word/excel)
			if (count === 1 && typeof AscFormat.CGraphicFrame !== "undefined" && drawing instanceof AscFormat.CGraphicFrame) {
				//в презентациях пропускаю чтение ещё раз графического объекта
				if (presentation.Slides) {
					loader.stream.Skip2(1);
					loader.stream.SkipRecord();
				} else {
					drawing = loader.ReadGraphicObject();
				}
			}

			var x = stream.GetULong() / 100000;
			var y = stream.GetULong() / 100000;
			var extX = stream.GetULong() / 100000;
			var extY = stream.GetULong() / 100000;
			var base64 = stream.GetString2();

			if (presentation.Slides)
				arr_shapes[i] = new DrawingCopyObject();
			else
				arr_shapes[i] = {};

			arr_shapes[i].Drawing = drawing;
			arr_shapes[i].X = x;
			arr_shapes[i].Y = y;
			arr_shapes[i].ExtX = extX;
			arr_shapes[i].ExtY = extY;
			if (!presentation.Slides)
				arr_shapes[i].base64 = base64;

			if (style_index != null && arr_shapes[i].Drawing.graphicObject && arr_shapes[i].Drawing.graphicObject.Set_TableStyle) {
				if (!PasteElementsId.g_bIsDocumentCopyPaste) {
					//TODO продумать добавления нового стиля(ReadTableStyle->получуть id нового стиля, сравнить новый стиль со всеми присутвующими.если нет - добавить и сделать Set_TableStyle(id))
					if (foundTableStylesIdMap[style_index]) {
						arr_shapes[i].Drawing.graphicObject.Set_TableStyle(foundTableStylesIdMap[style_index], true);
					} else if (cStyle && presentation.globalTableStyles && presentation.globalTableStyles.Style) {
						var isFoundStyle = false;
						for (var j in presentation.globalTableStyles.Style) {
							//TODO isEqual - сравнивает ещё и имя стиля. для случая, когда одинаковый контент, но имя стиля разное, не подойдет это сравнение
							if (presentation.globalTableStyles.Style[j].isEqual(cStyle)) {
								arr_shapes[i].Drawing.graphicObject.Set_TableStyle(j, true);
								foundTableStylesIdMap[style_index] = j;
								isFoundStyle = true;
								break;
							}
						}

						//в данном случае добавляем новый стиль
						if (!isFoundStyle) {
							//TODO при добавлении нового стиля - падение. пересмотреть!
							/*var newIndexStyle = presentation.globalTableStyles.Add(cStyle);
							presentation.TableStylesIdMap[newIndexStyle] = true;
							arr_shapes[i].Drawing.graphicObject.Set_TableStyle(newIndexStyle, true);
							foundTableStylesIdMap[style_index] = newIndexStyle;*/
						}
					} else if (presentation.TableStylesIdMap[style_index]) {
						arr_shapes[i].Drawing.graphicObject.Set_TableStyle(style_index, true);
					}
				} else if (cStyle) {
					//пока не применяем стили, посольку они отличаются
					//this._applyStylesToTable(arr_shapes[i].Drawing.graphicObject, cStyle);
				}
			}
		}

		var chartImages = pptx_content_loader.Reader.End_UseFullUrl();
		var images = loader.End_UseFullUrl();
		loader.AssignConnectedObjects();
		var allImages = chartImages.concat(images);

		return {arrShapes: arr_shapes, arrImages: allImages, arrTransforms: arr_transforms};
	},

	ReadPresentationSlides: function (stream) {
		var loader = new AscCommon.BinaryPPTYLoader();
		loader.Start_UseFullUrl();
		loader.stream = stream;
		loader.presentation = editor.WordControl.m_oLogicDocument;
		loader.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;
		var presentation = editor.WordControl.m_oLogicDocument;
		var count = stream.GetULong();
		var arr_slides = [];
		var slide;
		for (var i = 0; i < count; ++i) {
			loader.ClearConnectedObjects();
			slide = loader.ReadSlide(0);
			loader.AssignConnectedObjects();
			arr_slides.push(slide);
		}
		return arr_slides;
	},


	ReadSlide: function (stream) {
		var loader = new AscCommon.BinaryPPTYLoader();
		loader.Start_UseFullUrl();
		loader.stream = stream;
		loader.presentation = editor.WordControl.m_oLogicDocument;
		loader.DrawingDocument = editor.WordControl.m_oLogicDocument.DrawingDocument;
		var presentation = editor.WordControl.m_oLogicDocument;
		return loader.ReadSlide(0);
	},

	_Prepeare: function (node, fCallback) {
		var oThis = this;
		if (true === this.bUploadImage || true === this.bUploadFonts || this.pasteTextIntoList) {
			//Пробегаемся по документу собираем список шрифтов и картинок.
			var aPrepeareFonts;
			if (this.pasteTextIntoList) {
				this._Prepeare_recursive(node, true);
				fCallback();
				return;
			} else {
				aPrepeareFonts = this._Prepeare_recursive(node, true, true);
			}

			//TODO пересмотреть все "local" и сделать одинаковые проверки во всех редакторах
			var aImagesToDownload = [];
			var _mapLocal = {};
			var originalSrcArr = [];
			for (var image in this.oImages) {
				var src = this.oImages[image];
				if (undefined !== window["Native"] && undefined !== window["Native"]["GetImageUrl"]) {
					this.oImages[image] = window["Native"]["GetImageUrl"](this.oImages[image]);
				} else if (!g_oDocumentUrls.getImageLocal(src)) {
					if (oThis.rtfImages && oThis.rtfImages[src]) {
						aImagesToDownload.push(oThis.rtfImages[src]);
						originalSrcArr.push(src);
					} else {
						aImagesToDownload.push(src);
						originalSrcArr.push(src);
					}
				}
			}
			if (aImagesToDownload.length > 0) {
				AscCommon.sendImgUrls(oThis.api, aImagesToDownload, function (data) {
					var image_map = {};
					var isError = false;
					for (var i = 0, length = Math.min(data.length, aImagesToDownload.length); i < length; ++i) {
						var elem = data[i];
						var sFrom = originalSrcArr[i] ? originalSrcArr[i] : aImagesToDownload[i];
						if ("error" === elem.url) {
							isError = true;
						}
						if (null != elem.url) {
							var name = g_oDocumentUrls.imagePath2Local(elem.path);
							oThis.oImages[sFrom] = name;
							image_map[i] = name;
						} else {
							image_map[i] = sFrom;
						}
					}
					if (isError && PasteElementsId.g_bIsDocumentCopyPaste) {
						oThis.api.sendEvent("asc_onError", c_oAscError.ID.CanNotPasteImage, c_oAscError.Level.NoCritical);
					}
					fCallback(aPrepeareFonts, image_map);
				}, true);
			} else {
				fCallback(aPrepeareFonts, this.oImages);
			}
		} else {
			fCallback();
		}
	},
	_Prepeare_recursive: function (node, bIgnoreStyle, isCheckFonts) {
		//пробегаемся по всему дереву, собираем все шрифты и картинки
		var nodeName = node.nodeName.toLowerCase();
		var nodeType = node.nodeType;
		if (!bIgnoreStyle) {
			if (Node.TEXT_NODE === nodeType) {
				var computedStyle = this._getComputedStyle(node.parentNode);
				var fontFamily = CheckDefaultFontFamily(this._getStyle(node.parentNode, computedStyle, "font-family"), this.apiEditor);
				if (fontFamily) {
					this.oFonts[fontFamily] = {
						Name: g_fontApplication.GetFontNameDictionary(fontFamily, true),
						Index: -1
					};
				}
			} else {
				var src = node.getAttribute("src");
				if (src && !this._checkSkipMath(node))
					this.oImages[src] = src;
			}
		}
		for (var i = 0, length = node.childNodes.length; i < length; i++) {
			var child = node.childNodes[i];

			//get comment
			var style = child.getAttribute ? child.getAttribute("style") : null;
			if (style && -1 !== style.indexOf("mso-element:comment") && -1 === style.indexOf("mso-element:comment-list")) {
				this._parseMsoElementComment(child);
			}
			if (this.pasteTextIntoList && (nodeName === "li" || nodeName === "ol" || nodeName === "ul" || (style && -1 !== style.indexOf("mso-list")))) {
				this.pasteTextIntoList = null;
			}

			//принудительно добавляю для математики шрифт Cambria Math
			if (child && child.nodeName.toLowerCase() === "#comment" && this.isSupportPasteMathContent(child.nodeValue, true) && !this.pasteInExcel && this.apiEditor["asc_isSupportFeature"]("ooxml")) {
				//TODO пока только в документы разрешаю вставку математики математику
				var mathFont = "Cambria Math";
				this.oFonts[mathFont] = {
					Name: g_fontApplication.GetFontNameDictionary(mathFont, true),
					Index: -1
				};
			}

			var child_nodeType = child.nodeType;
			if (!(Node.ELEMENT_NODE === child_nodeType || Node.TEXT_NODE === child_nodeType))
				continue;
			//попускам элеметы состоящие только из \t,\n,\r
			if (Node.TEXT_NODE === child.nodeType) {
				var value = child.nodeValue;
				if (!value)
					continue;
				value = value.replace(/(\r|\t|\n)/g, '');
				if ("" == value)
					continue;
			}
			this._Prepeare_recursive(child, false);
		}

		if (isCheckFonts) {
			var aPrepeareFonts = [];
			for (var font_family in this.oFonts) {
				//todo подбирать шрифт, хотябы по регистру
				var oFontItem = this.oFonts[font_family];
				//Ищем среди наших шрифтов
				this.oFonts[font_family].Index = -1;
				aPrepeareFonts.push(new CFont(oFontItem.Name));
			}

			return aPrepeareFonts;
		}
	},
	_parseMsoElementComment: function (node) {
		var msoTextNode = node.getElementsByClassName("MsoCommentText");
		if (msoTextNode && msoTextNode[0]) {
			var msoComment = this._getMsoCommentText(msoTextNode[0]);
			var msoCommentId = msoTextNode[0].parentElement.id;
			//в качестве id использую индекс из id у родительского элемента
			//id вида _com_1
			if (msoCommentId) {
				var id = msoCommentId.split("_com_");
				if (id && undefined !== id[1]) {
					this.msoComments[id[1]] = {text: msoComment, start: false};
				}
			}
		}
	},

	_checkSkipMath: function (node) {
		if (!this.pasteInExcel && this.apiEditor["asc_isSupportFeature"]("ooxml")) {
			let parent = node && node.parentNode;
			if (parent && parent.nodeName.toLowerCase() === "span") {
				parent = parent && parent.parentNode;
			}
			if (parent && (parent.nodeName.toLowerCase() === "p" || parent.nodeName.toLowerCase() === "body")) {
				for (let i = 0; i < parent.childNodes.length; i++) {
					let child = parent.childNodes[i];
					if (child && child.nodeName.toLowerCase() === "#comment" && this.isSupportPasteMathContent(child.nodeValue, true)) {
						return true;
					}
				}
			}
		}

		return false;
	},

	_getMsoCommentText: function (node) {
		var res = "";
		var bMsoAnnotation = false;
		var elems = node.childNodes;
		if (elems && elems.length) {
			for (var i = 0; i < elems.length; i++) {
				var child = elems[i];

				//нужно исключить <![if !supportAnnotations]>
				//пока собираем только текст(не форматированный), в дальнейшем, когда будет полная поддержка отображения
				//коментариев, можно преобразовывать уже в структуру и вставлять в комм. форматированный текст
				if (child.nodeName === "#comment") {
					if (child.nodeValue === "[if !supportAnnotations]") {
						bMsoAnnotation = true;
					} else if (bMsoAnnotation && child.nodeValue === "[endif]") {
						bMsoAnnotation = false;
					}
				}
				if (bMsoAnnotation) {
					continue;
				}

				if (Node.TEXT_NODE === child.nodeType) {
					var value = child.nodeValue;
					//пропускаем неразрывный пробел перед комментарием
					if (value === " " && child.parentElement && child.parentElement.getAttribute("style") === "mso-special-character:comment") {
						continue;
					}
					if (!value) {
						continue;
					}
					value = value.replace(/(\r|\t|\n)/g, '');
					if ("" === value) {
						continue;
					}
					res += value;
				} else {
					res += this._getMsoCommentText(child);
				}
			}
		}
		return res;
	},
	_checkFontsOnLoad: function (fonts) {
		if (!fonts)
			return;
		return fonts;
	},
	_IsBlockElem: function (name) {
		if ("p" === name || "div" === name || "ul" === name || "ol" === name || "li" === name || "table" === name || "tbody" === name || "tr" === name || "td" === name || "th" === name ||
			"h1" === name || "h2" === name || "h3" === name || "h4" === name || "h5" === name || "h6" === name || "center" === name || "dd" === name || "dt" === name)
			return true;
		return false;
	},
	_getComputedStyle: function (node) {
		var computedStyle = null;
		if (null != node && Node.ELEMENT_NODE === node.nodeType) {
			var defaultView = node.ownerDocument.defaultView;
			computedStyle = defaultView.getComputedStyle(node, null);
		}
		return computedStyle;
	},
	_getStyle: function (node, computedStyle, styleProp) {
		var getStyle = function () {
			var style = null;
			if (computedStyle) {
				style = computedStyle.getPropertyValue(styleProp);
			}
			if (!style) {
				if (node && node.currentStyle) {
					style = node.currentStyle[styleProp];
				}
			}

			return style;
		};

		var res = getStyle();
		var sNodeName = node.nodeName.toLowerCase();
		if (!res) {
			switch (styleProp) {
				case "font-family": {
					res = node.fontFamily;
					break;
				}
				case "background-color": {
					res = node.style.backgroundColor;
					if ("td" === sNodeName && !res) {
						res = node.bgColor;
					}
					break;
				}
				case "font-weight": {
					res = node.style.fontWeight;
				}
			}
		}

		return res;
	},

	_ParseColor: function (color) {
		if (!color || color.length == 0) {
			return null;
		}
		if ("transparent" === color) {
			return null;
		}

		if ("aqua" === color) {
			return new CDocumentColor(0, 255, 255);
		} else if ("black" === color) {
			return new CDocumentColor(0, 0, 0);
		} else if ("blue" === color) {
			return new CDocumentColor(0, 0, 255);
		} else if ("fuchsia" === color) {
			return new CDocumentColor(255, 0, 255);
		} else if ("gray" === color) {
			return new CDocumentColor(128, 128, 128);
		} else if ("green" === color) {
			return new CDocumentColor(0, 128, 0);
		} else if ("lime" === color) {
			return new CDocumentColor(0, 255, 0);
		} else if ("maroon" === color) {
			return new CDocumentColor(128, 0, 0);
		} else if ("navy" === color) {
			return new CDocumentColor(0, 0, 128);
		} else if ("olive" === color) {
			return new CDocumentColor(128, 128, 0);
		} else if ("purple" === color) {
			return new CDocumentColor(128, 0, 128);
		} else if ("red" === color) {
			return new CDocumentColor(255, 0, 0);
		} else if ("silver" === color) {
			return new CDocumentColor(192, 192, 192);
		} else if ("teal" === color) {
			return new CDocumentColor(0, 128, 128);
		} else if ("white" === color) {
			return new CDocumentColor(255, 255, 255);
		} else if ("yellow" === color) {
			return new CDocumentColor(255, 255, 0);
		} else if ("cyan" === color) {
			return new CDocumentColor(0, 255, 255);
		} else if ("magenta" === color) {
			return new CDocumentColor(255, 0, 255);
		} else if ("darkblue" === color) {
			return new CDocumentColor(0, 0, 139);
		} else if ("darkcyan" === color) {
			return new CDocumentColor(0, 139, 139);
		} else if ("darkGreen" === color) {
			return new CDocumentColor(0, 100, 0);
		} else if ("darkmagenta" === color) {
			return new CDocumentColor(128, 0, 128);
		} else if ("darkRed" === color) {
			return new CDocumentColor(139, 0, 0);
		} else if ("darkyellow" === color) {
			return new CDocumentColor(128, 128, 0);
		} else if ("darkgray" === color) {
			return new CDocumentColor(169, 169, 169);
		} else if ("lightgray" === color) {
			return new CDocumentColor(211, 211, 211);
		} else {
			var r, g, b;
			if (0 === color.indexOf("#")) {
				var hex = color.substring(1);
				if (hex.length === 3) {
					hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
				}
				if (hex.length === 6) {
					r = parseInt("0x" + hex.substring(0, 2));
					g = parseInt("0x" + hex.substring(2, 4));
					b = parseInt("0x" + hex.substring(4, 6));
					return new CDocumentColor(r, g, b);
				}
			}
			if (0 === color.indexOf("rgb")) {
				var nStart = color.indexOf('(');
				var nEnd = color.indexOf(')');
				if (-1 !== nStart && -1 !== nEnd && nStart < nEnd) {
					var temp = color.substring(nStart + 1, nEnd);
					var aParems = temp.split(',');
					if (aParems.length >= 3) {
						if (aParems.length >= 4) {
							var oA = AscCommon.valueToMmType(aParems[3]);
							if (0 == oA.val)//полностью прозрачный
							{
								return null;
							}
						}
						var oR = AscCommon.valueToMmType(aParems[0]);
						var oG = AscCommon.valueToMmType(aParems[1]);
						var oB = AscCommon.valueToMmType(aParems[2]);
						if (oR && "%" === oR.type) {
							r = parseInt(255 * oR.val / 100);
						} else {
							r = oR.val;
						}
						if (oG && "%" === oG.type) {
							g = parseInt(255 * oG.val / 100);
						} else {
							g = oG.val;
						}
						if (oB && "%" === oB.type) {
							b = parseInt(255 * oB.val / 100);
						} else {
							b = oB.val;
						}
						return new CDocumentColor(r, g, b);
					}
				}
			}
		}
		return null;
	}, _isEmptyProperty: function (prop) {
		var bIsEmpty = true;
		for (var i in prop) {
			if (null != prop[i]) {
				bIsEmpty = false;
				break;
			}
		}
		return bIsEmpty;
	},
	_set_pPr: function (node, Para, pNoHtmlPr) {
		//Пробегаемся вверх по дереву в поисках блочного элемента
		var t = this;
		var sNodeName = node.nodeName.toLowerCase();
		if (node !== this.oRootNode) {
			while (false === this._IsBlockElem(sNodeName)) {
				if (this.oRootNode !== node.parentNode) {
					node = node.parentNode;
					sNodeName = node.nodeName.toLowerCase();
				} else {
					break;
				}
			}
		}

		var _applyTextAlign = function () {
			let text_align;
			if (node.style && node.align && !node.style.textAlign) {
				//some editors(LO) put old attr -> align, and skip text-align
				text_align = node.align;
			} else {
				text_align = t._getStyle(node, computedStyle, "text-align");
			}
			if (text_align) {
				//Может приходить -webkit-right
				let Jc = null;
				if (-1 !== text_align.indexOf('center')) {
					Jc = align_Center;
				} else if (-1 !== text_align.indexOf('right')) {
					Jc = align_Right;
				} else if (-1 !== text_align.indexOf('justify')) {
					Jc = align_Justify;
				}
				if (null != Jc) {
					Para.Set_Align(Jc, false);
				}
			}
		};

		var computedStyle = this._getComputedStyle(node);
		if ("td" === sNodeName || "th" === sNodeName) {
			_applyTextAlign();

			//для случая <td>br<span></span></td> без текста в ячейке
			var oNewSpacing = new CParaSpacing();
			oNewSpacing.Set_FromObject({After: 0, Before: 0, Line: Asc.linerule_Auto});
			Para.Set_Spacing(oNewSpacing);
			return;
		}
		var oDocument = this.oDocument;

		//Heading
		//Ранее применялся весь заголовок - Para.Style_Add(oDocument.Styles.Get_Default_Heading(pNoHtmlPr.hLevel));
		if (null != pNoHtmlPr.hLevel && oDocument.Styles) {
			//Para.SetOutlineLvl(pNoHtmlPr.hLevel);
			Para.Style_Add(oDocument.Styles.Get_Default_Heading(pNoHtmlPr.hLevel));
		}

		var pPr = Para.Pr;

		//Borders
		var oNewBorder = {Left: null, Top: null, Right: null, Bottom: null, Between: null};
		var sBorder = pNoHtmlPr["mso-border-alt"];
		if (null != sBorder) {
			var oNewBrd = this._ExecuteParagraphBorder(sBorder);
			oNewBorder.Left = oNewBrd;
			oNewBorder.Top = oNewBrd.Copy();
			oNewBorder.Right = oNewBrd.Copy();
			oNewBorder.Bottom = oNewBrd.Copy();
		} else {
			sBorder = pNoHtmlPr["mso-border-left-alt"];
			if (null != sBorder) {
				var oNewBrd = this._ExecuteParagraphBorder(sBorder);
				oNewBorder.Left = oNewBrd;
			}
			sBorder = pNoHtmlPr["mso-border-top-alt"];
			if (null != sBorder) {
				var oNewBrd = this._ExecuteParagraphBorder(sBorder);
				oNewBorder.Top = oNewBrd;
			}
			sBorder = pNoHtmlPr["mso-border-right-alt"];
			if (null != sBorder) {
				var oNewBrd = this._ExecuteParagraphBorder(sBorder);
				oNewBorder.Right = oNewBrd;
			}
			sBorder = pNoHtmlPr["mso-border-bottom-alt"];
			if (null != sBorder) {
				var oNewBrd = this._ExecuteParagraphBorder(sBorder);
				oNewBorder.Bottom = oNewBrd;
			}
		}
		sBorder = pNoHtmlPr["mso-border-between"];
		if (null != sBorder) {
			var oNewBrd = this._ExecuteParagraphBorder(sBorder);
			oNewBorder.Between = oNewBrd;
		}

		if (computedStyle) {
			var font_family = CheckDefaultFontFamily(this._getStyle(node, computedStyle, "font-family"), this.apiEditor);
			if (font_family && "" != font_family) {
				var oFontItem = this.oFonts[font_family];
				if (null != oFontItem && null != oFontItem.Name && Para.TextPr && Para.TextPr.Value && Para.TextPr.Value.RFonts) {
					Para.TextPr.Value.RFonts.Ascii = {Name: oFontItem.Name, Index: oFontItem.Index};
					Para.TextPr.Value.RFonts.HAnsi = {Name: oFontItem.Name, Index: oFontItem.Index};
					Para.TextPr.Value.RFonts.CS = {Name: oFontItem.Name, Index: oFontItem.Index};
					Para.TextPr.Value.RFonts.EastAsia = {Name: oFontItem.Name, Index: oFontItem.Index};
				}
			}

			var font_size = node.style ? node.style.fontSize : null;
			if (!font_size) {
				font_size = this._getStyle(node, computedStyle, "font-size");
			}
			font_size = CheckDefaultFontSize(font_size, this.apiEditor);
			if (font_size && Para.TextPr && Para.TextPr.Value) {
				var obj = AscCommon.valueToMmType(font_size);
				if (obj && "%" !== obj.type && "none" !== obj.type) {
					font_size = obj.val;
					//Если браузер не поддерживает нецелые пикселы отсекаем половинные шрифты, они появляются при вставке 8, 11, 14, 20, 26pt
					if ("px" === obj.type && false === this.bIsDoublePx) {
						font_size = Math.round(font_size * g_dKoef_mm_to_pt);
					} else {
						font_size = Math.round(2 * font_size * g_dKoef_mm_to_pt) / 2;
					}//половинные значения допустимы.

					//TODO use constant
					if (font_size > 300) {
						font_size = 300;
					} else if (font_size === 0) {
						font_size = 1;
					}

					Para.TextPr.Value.FontSize = font_size;
				}
			}

			//Ind
			var Ind = new CParaInd();
			var margin_left = this._getStyle(node, computedStyle, "margin-left");

			//TODO перепроверить правку с pageColumn
			var curContent = this.oLogicDocument.Content[this.oLogicDocument.CurPos.ContentPos];
			var curIndexColumn = curContent && curContent.Get_CurrentColumn ? curContent.Get_CurrentColumn(this.oLogicDocument.CurPage) : null;
			var curPage = this.oLogicDocument.Pages[this.oLogicDocument.CurPage];
			var pageColumn = null !== curIndexColumn && curPage && curPage.Sections && curPage.Sections[0] && curPage.Sections[0].Columns ?
				curPage.Sections[0].Columns[curIndexColumn] : null;
			if (margin_left && null != (margin_left = AscCommon.valueToMm(margin_left))) {
				if (!pageColumn || (pageColumn && pageColumn.X + margin_left < pageColumn.XLimit)) {
					Ind.Left = margin_left;
				}
			}
			var margin_right = this._getStyle(node, computedStyle, "margin-right");
			if (margin_right && null != (margin_right = AscCommon.valueToMm(margin_right))) {
				if (!pageColumn || (pageColumn && pageColumn.XLimit - margin_right > pageColumn.X)) {
					Ind.Right = margin_right;
				}
			}

			//scale
			// if(null != pPr.Ind.Left && true == this.bUseScaleKoef)
			// pPr.Ind.Left = pPr.Ind.Left * this.dScaleKoef;
			// if(null != pPr.Ind.Right && true == this.bUseScaleKoef)
			// pPr.Ind.Right = pPr.Ind.Right * this.dScaleKoef;
			//Проверка чтобы правый margin не заходил за левый или не приближался ближе чем на линейке
			if (null != Ind.Left && null != Ind.Right) {
				//30 ограничение как и на линейке
				var dif = Page_Width - X_Left_Margin - X_Right_Margin - Ind.Left - Ind.Right;
				if (dif < 30) {
					Ind.Right = Page_Width - X_Left_Margin - X_Right_Margin - Ind.Left - 30;
				}
			}
			var text_indent = this._getStyle(node, computedStyle, "text-indent");
			if (text_indent && null != (text_indent = AscCommon.valueToMm(text_indent))) {
				Ind.FirstLine = text_indent;
			}
			// if(null != pPr.Ind.FirstLine && true == this.bUseScaleKoef)
			// pPr.Ind.FirstLine = pPr.Ind.FirstLine * this.dScaleKoef;
			if (false === this._isEmptyProperty(Ind) && !pNoHtmlPr['mso-list']) {
				Para.Set_Ind(Ind);
			}

			//Jc
			_applyTextAlign();

			//Spacing
			//use not computedStyle -> html often comes with the font size set by the parent, the browser calculate
			//the font to margin-top/margin-bottom
			var Spacing = new CParaSpacing();
			var margin_top = node.style.getPropertyValue("margin-top")/*this._getStyle(node, computedStyle, "margin-top")*/;
			if (margin_top && null != (margin_top = AscCommon.valueToMm(margin_top)) && margin_top >= 0) {
				Spacing.Before = margin_top;
			}
			var margin_bottom =  node.style.getPropertyValue("margin-bottom")/*this._getStyle(node, computedStyle, "margin-bottom")*/;
			if (margin_bottom && null != (margin_bottom = AscCommon.valueToMm(margin_bottom)) && margin_bottom >= 0) {
				Spacing.After = margin_bottom;
			}
			//line height
			//computedStyle возвращает значение в px. мне нужны %(ms записывает именно % в html)
			var line_height = node.style && node.style.lineHeight ? node.style.lineHeight : this._getStyle(node, computedStyle, "line-height");
			if (line_height) {
				var oLineHeight = AscCommon.valueToMmType(line_height);
				if (oLineHeight && ("%" === oLineHeight.type || "none" === oLineHeight.type)) {
					Spacing.Line = oLineHeight.val;
				} else if (line_height && null != (line_height = AscCommon.valueToMm(line_height)) && line_height >= 0) {
					Spacing.Line = line_height;
					Spacing.LineRule = Asc.linerule_AtLeast;
				}
			}
			if (false === this._isEmptyProperty(Spacing)) {
				Para.Set_Spacing(Spacing);
			}
			//Shd
			//background-color не наследуется остальные свойства, надо смотреть родительские элементы
			var background_color = null;
			var oTempNode = node;
			while (true) {
				var tempComputedStyle = this._getComputedStyle(oTempNode);
				if (null == tempComputedStyle) {
					break;
				}
				background_color = this._getStyle(oTempNode, tempComputedStyle, "background-color");
				if (null != background_color && (background_color = this._ParseColor(background_color))) {
					break;
				}

				oTempNode = oTempNode.parentNode;
				let _nodeName = oTempNode && oTempNode.nodeName && oTempNode.nodeName.toLowerCase();
				if (!oTempNode || this.oRootNode === oTempNode || "body" === _nodeName || true === this._IsBlockElem(_nodeName)) {
					break;
				}
			}
			if (PasteElementsId.g_bIsDocumentCopyPaste) {
				if (background_color) {
					var Shd = new CDocumentShd();
					Shd.Value = c_oAscShdClear;
					Shd.Color = background_color;
					Shd.Fill = background_color;
					Para.Set_Shd(Shd);
				}
			}

			if (null == oNewBorder.Left) {
				oNewBorder.Left = this._ExecuteBorder(computedStyle, node, "left", "Left", false);
			}
			if (null == oNewBorder.Top) {
				oNewBorder.Top = this._ExecuteBorder(computedStyle, node, "top", "Top", false);
			}
			if (null == oNewBorder.Right) {
				oNewBorder.Right = this._ExecuteBorder(computedStyle, node, "left", "Left", false);
			}
			if (null == oNewBorder.Bottom) {
				oNewBorder.Bottom = this._ExecuteBorder(computedStyle, node, "bottom", "Bottom", false);
			}
		}
		if (false === this._isEmptyProperty(oNewBorder)) {
			Para.Set_Borders(oNewBorder);
		}

		//KeepLines , WidowControl
		var pagination = pNoHtmlPr["mso-pagination"];
		if (pagination) {
			//todo WidowControl
			if ("none" === pagination) {
				;
			}//pPr.WidowControl = !Def_pPr.WidowControl;
			else if (-1 !== pagination.indexOf("widow-orphan") && -1 !== pagination.indexOf("lines-together")) {
				Para.Set_KeepLines(true);
			} else if (-1 !== pagination.indexOf("none") && -1 !== pagination.indexOf("lines-together")) {
				;//pPr.WidowControl = !Def_pPr.WidowControl;
				Para.Set_KeepLines(true);
			}
		}
		//todo KeepNext
		if ("avoid" === pNoHtmlPr["page-break-after"]) {
			;
		}//pPr.KeepNext = !Def_pPr.KeepNext;
		//PageBreakBefore
		if ("always" === pNoHtmlPr["page-break-before"]) {
			Para.Set_PageBreakBefore(true);
		}
		//Tabs
		var tab_stops = pNoHtmlPr["tab-stops"];
		if (tab_stops && "" != pNoHtmlPr["tab-stops"]) {
			var aTabs = tab_stops.split(' ');
			var nTabLen = aTabs.length;
			if (nTabLen > 0) {
				var Tabs = new CParaTabs();
				for (var i = 0; i < nTabLen; i++) {
					var val = AscCommon.valueToMm(aTabs[i]);
					if (val) {
						Tabs.Add(new CParaTab(tab_Left, val));
					}
				}
				Para.Set_Tabs(Tabs);
			}
		}

		//*****num*****
		if (PasteElementsId.g_bIsDocumentCopyPaste) {
			if (true === pNoHtmlPr.bNum) {
				var setListTextPr = function (oNum) {
					//текстовые настройки списка берем по настройкам первого текстового элемента
					var oFirstTextChild = node;
					while (true) {
						var bContinue = false;
						for (var i = 0, length = oFirstTextChild.childNodes.length; i < length; i++) {
							var child = oFirstTextChild.childNodes[i];
							var nodeType = child.nodeType;

							if (!(Node.ELEMENT_NODE === nodeType || Node.TEXT_NODE === nodeType)) {
								continue;
							}
							//попускам элеметы состоящие только из \t,\n,\r
							if (Node.TEXT_NODE === child.nodeType) {
								var value = child.nodeValue;
								if (!value) {
									continue;
								}
								value = value.replace(/(\r|\t|\n)/g, '');
								if ("" === value) {
									continue;
								}
							}
							if (Node.ELEMENT_NODE === nodeType) {
								oFirstTextChild = child;
								bContinue = true;
								break;
							}
						}
						if (false === bContinue) {
							break;
						}
					}
					if (node != oFirstTextChild) {
						if (!t.bIsPlainText) {
							var oLvl = oNum.GetLvl(0);
							var oTextPr = t._read_rPr(oFirstTextChild);
							if (Asc.c_oAscNumberingFormat.Bullet === num) {
								oTextPr.RFonts = oLvl.GetTextPr().RFonts.Copy();
							}

							//TODO убираю пока при всатвке извне underline/bold/italic у стиля маркера
							oTextPr.Bold = oTextPr.Underline = oTextPr.Italic = undefined;
							if (oFirstTextChild.nodeName.toLowerCase() === "a" && oTextPr.Color) {
								oTextPr.Color.Set(0, 0, 0);
							}

							//получаем настройки из node
							oNum.ApplyTextPr(0, oTextPr);
						}
					}
				};


				if (pNoHtmlPr['mso-list']) {
					var level = 0;
					var listId = null;
					var startIndex;
					var listName;
					if (-1 !== (startIndex = pNoHtmlPr['mso-list'].indexOf("level"))) {
						level = parseInt(pNoHtmlPr['mso-list'].substr(startIndex + 5, 1)) - 1;
					}
					if (-1 !== (startIndex = pNoHtmlPr['mso-list'].indexOf("lfo"))) {
						listId = pNoHtmlPr['mso-list'].substr(startIndex, 4);
					}
					if (0 === pNoHtmlPr['mso-list'].indexOf("l")) {
						listName = pNoHtmlPr['mso-list'].split(" ");
						listName = (listName && listName[0]) ? listName[0] : null;
					}

					var NumId = null;
					if (listId && this.msoListMap[listId])//find list id into map
					{
						NumId = this.msoListMap[listId];
					}

					//get listId and level from mso-list property
					var msoListIgnoreSymbol = this._getMsoListSymbol(node);
					if (!msoListIgnoreSymbol) {
						msoListIgnoreSymbol = "ol" === node.parentElement.nodeName.toLowerCase() ? "1." : ".";
					}

					if (null == NumId && this.pasteInPresentationShape !== true && !this.oMsoHeadStylesListMap[listId]) {
						this.oMsoHeadStylesListMap[listId] = this._findElemFromMsoHeadStyle("@list", listName);
						var _msoNum = this._tryGenerateNumberingFromMsoStyle(this.oMsoHeadStylesListMap[listId], node);
						if (_msoNum) {
							NumId = _msoNum.GetId();
						}
					}

					var listObj = this._getTypeMsoListSymbol(msoListIgnoreSymbol, (null === NumId));
					var num = listObj.type;
					var startPos = listObj.startPos;

					if (null == NumId && this.pasteInPresentationShape !== true)//create new NumId
					{
						// Создаем нумерацию
						var oNum = this.oLogicDocument.GetNumbering().CreateNum();
						NumId = oNum.GetId();

						if (Asc.c_oAscNumberingFormat.Bullet === num) {
							oNum.CreateDefault(c_oAscMultiLevelNumbering.Bullet);
							var LvlText = String.fromCharCode(0x00B7);
							var NumTextPr = new CTextPr();
							NumTextPr.RFonts.SetAll("Symbol", -1);

							switch (type) {
								case "disc": {
									NumTextPr.RFonts.SetAll("Symbol", -1);
									LvlText = String.fromCharCode(0x00B7);
									break;
								}
								case "circle": {
									NumTextPr.RFonts.SetAll("Courier New", -1);
									LvlText = "o";
									break;
								}
								case "square": {
									NumTextPr.RFonts.SetAll("Wingdings", -1);
									LvlText = String.fromCharCode(0x00A7);
									break;
								}
							}
						} else {
							oNum.CreateDefault(c_oAscMultiLevelNumbering.Numbered);
						}

						switch (num) {
							case Asc.c_oAscNumberingFormat.Bullet     :
								oNum.SetLvlByType(level, c_oAscNumberingLevel.Bullet, LvlText, NumTextPr);
								break;
							case Asc.c_oAscNumberingFormat.Decimal    :
								oNum.SetLvlByType(level, c_oAscNumberingLevel.DecimalDot_Left);
								break;
							case Asc.c_oAscNumberingFormat.LowerRoman :
								oNum.SetLvlByType(level, c_oAscNumberingLevel.LowerRomanDot_Right);
								break;
							case Asc.c_oAscNumberingFormat.UpperRoman :
								oNum.SetLvlByType(level, c_oAscNumberingLevel.UpperRomanDot_Right);
								break;
							case Asc.c_oAscNumberingFormat.LowerLetter:
								oNum.SetLvlByType(level, c_oAscNumberingLevel.LowerLetterDot_Left);
								break;
							case Asc.c_oAscNumberingFormat.UpperLetter:
								oNum.SetLvlByType(level, c_oAscNumberingLevel.UpperLetterDot_Left);
								break;
						}

						//проставляем начальную позицию
						if (null !== startPos) {
							oNum.SetLvlStart(level, startPos);
						}

						//setListTextPr(oNum);
					}

					//put into map listId
					if (!this.msoListMap[listId]) {
						this.msoListMap[listId] = NumId;
					}

					if (this.pasteInPresentationShape !== true && Para.bFromDocument === true) {
						Para.SetNumPr(NumId, level);
					}
				} else {
					var num = Asc.c_oAscNumberingFormat.Bullet;
					if (null != pNoHtmlPr.numType) {
						num = pNoHtmlPr.numType;
					}
					var type = pNoHtmlPr["list-style-type"];

					if (type) {
						switch (type) {
							case "disc"       :
								num = Asc.c_oAscNumberingFormat.Bullet;
								break;
							case "decimal"    :
								num = Asc.c_oAscNumberingFormat.Decimal;
								break;
							case "lower-roman":
								num = Asc.c_oAscNumberingFormat.LowerRoman;
								break;
							case "upper-roman":
								num = Asc.c_oAscNumberingFormat.UpperRoman;
								break;
							case "lower-alpha":
								num = Asc.c_oAscNumberingFormat.LowerLetter;
								break;
							case "upper-alpha":
								num = Asc.c_oAscNumberingFormat.UpperLetter;
								break;
						}
					}
					//Часть кода скопирована из Document.Set_ParagraphNumbering

					//Смотрим передыдущий параграф, если тип списка совпадает, то берем тип списка из предыдущего параграфа
					if (this.aContent.length > 1) {
						var prevElem = this.aContent[this.aContent.length - 2];
						if (null != prevElem && type_Paragraph === prevElem.GetType()) {
							var PrevNumPr = prevElem.GetNumPr();
							if (null != PrevNumPr && true === this.oLogicDocument.Numbering.CheckFormat(PrevNumPr.NumId, PrevNumPr.Lvl, num)) {
								NumId = PrevNumPr.NumId;
							}
						}
					}
					if (null == NumId && this.pasteInPresentationShape !== true) {
						// Создаем нумерацию
						var oNum = this.oLogicDocument.GetNumbering().CreateNum();
						NumId = oNum.GetId();
						if (Asc.c_oAscNumberingFormat.Bullet === num) {
							oNum.CreateDefault(c_oAscMultiLevelNumbering.Bullet);
							var LvlText = String.fromCharCode(0x00B7);
							var NumTextPr = new CTextPr();
							NumTextPr.RFonts.SetAll("Symbol", -1);

							switch (type) {
								case "disc": {
									NumTextPr.RFonts.SetAll("Symbol", -1);
									LvlText = String.fromCharCode(0x00B7);
									break;
								}
								case "circle": {
									NumTextPr.RFonts.SetAll("Courier New", -1);
									LvlText = "o";
									break;
								}
								case "square": {
									NumTextPr.RFonts.SetAll("Wingdings", -1);
									LvlText = String.fromCharCode(0x00A7);
									break;
								}
							}
						} else {
							oNum.CreateDefault(c_oAscMultiLevelNumbering.Numbered);
						}

						for (var iLvl = 0; iLvl <= 8; iLvl++) {
							switch (num) {
								case Asc.c_oAscNumberingFormat.Bullet     :
									oNum.SetLvlByType(iLvl, c_oAscNumberingLevel.Bullet, LvlText, NumTextPr);
									break;
								case Asc.c_oAscNumberingFormat.Decimal    :
									oNum.SetLvlByType(iLvl, c_oAscNumberingLevel.DecimalDot_Right);
									break;
								case Asc.c_oAscNumberingFormat.LowerRoman :
									oNum.SetLvlByType(iLvl, c_oAscNumberingLevel.LowerRomanDot_Right);
									break;
								case Asc.c_oAscNumberingFormat.UpperRoman :
									oNum.SetLvlByType(iLvl, c_oAscNumberingLevel.UpperRomanDot_Right);
									break;
								case Asc.c_oAscNumberingFormat.LowerLetter:
									oNum.SetLvlByType(iLvl, c_oAscNumberingLevel.LowerLetterDot_Left);
									break;
								case Asc.c_oAscNumberingFormat.UpperLetter:
									oNum.SetLvlByType(iLvl, c_oAscNumberingLevel.UpperLetterDot_Left);
									break;
							}
						}

						setListTextPr(oNum);
					}

					if (this.pasteInPresentationShape !== true && Para.bFromDocument === true) {
						Para.ApplyNumPr(NumId, 0);
					}
				}
			} else {
				var numPr = Para.GetNumPr();
				if (numPr) {
					Para.RemoveNumPr();
				}
			}
		} else {
			if (true === pNoHtmlPr.bNum) {
				var num = AscFormat.numbering_presentationnumfrmt_Char;
				if (null != pNoHtmlPr.numType) {
					num = pNoHtmlPr.numType;
				}
				var type = pNoHtmlPr["list-style-type"];
				var oBullet = null;
				if (type) {
					switch (type) {
						case "disc": {
							oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 0, SubType: 1});
							break;
						}
						case "decimal": {
							oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 1, SubType: 0});
							break;
						}

						case "lower-roman": {
							oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 1, SubType: 7});
							break;
						}
						case "upper-roman": {
							oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 1, SubType: 3});
							break;
						}
						case "lower-alpha": {
							oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 1, SubType: 6});
							break;
						}
						case "upper-alpha": {
							oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 1, SubType: 4});
							break;
						}
						default: {
							oBullet = AscFormat.fGetPresentationBulletByNumInfo({Type: 0, SubType: 1});
							break;
						}
					}
				}
				Para.Add_PresentationNumbering(oBullet);
			} else {
				Para.Remove_PresentationNumbering();
			}
		}
		Para.CompiledPr.NeedRecalc = true;
	},
	_commit_rPr: function (node, bUseOnlyInherit) {
		if (!this.bIsPlainText) {
			var rPr = this._read_rPr(node, bUseOnlyInherit);

			//заглушка для вставки в excel внутрь шейпа
			var tempRpr;
			var bSaveExcelFormat = window['AscCommon'].g_clipboardBase.bSaveFormat;
			if (this.pasteInExcel === true && !bSaveExcelFormat && this.oDocument && this.oDocument.Parent &&
				this.oDocument.Parent.parent &&
				this.oDocument.Parent.parent.getObjectType() === AscDFH.historyitem_type_Shape) {
				tempRpr = new CTextPr();
				tempRpr.Underline = rPr.Underline;
				tempRpr.Bold = rPr.Bold;
				tempRpr.Italic = rPr.Italic;

				rPr = tempRpr;
			}

			//Если текстовые настройки поменялись добавляем элемент
			if (!this.oCur_rPr.Is_Equal(rPr)) {
				this._Set_Run_Pr(rPr);
				this.oCur_rPr = rPr;
			}
		}
	},
	_read_rPr: function (node, bUseOnlyInherit) {
		var oDocument = this.oDocument;
		var rPr = new CTextPr();
		if (false == PasteElementsId.g_bIsDocumentCopyPaste) {
			rPr.Set_FromObject({
				Bold: false,
				Italic: false,
				Underline: false,
				Strikeout: false,
				RFonts:
					{
						Ascii: {
							Name: "Arial",
							Index: -1
						},
						EastAsia: {
							Name: "Arial",
							Index: -1
						},
						HAnsi: {
							Name: "Arial",
							Index: -1
						},
						CS: {
							Name: "Arial",
							Index: -1
						}
					},
				FontSize: 11,
				Color:
					{
						r: 0,
						g: 0,
						b: 0
					},
				VertAlign: AscCommon.vertalign_Baseline,
				HighLight: highlight_None
			});
		}
		var computedStyle = this._getComputedStyle(node);
		if (computedStyle) {
			var font_family = CheckDefaultFontFamily(this._getStyle(node, computedStyle, "font-family"), this.apiEditor);
			if (font_family && "" != font_family) {
				var oFontItem = this.oFonts[font_family];
				if (null != oFontItem && null != oFontItem.Name) {
					rPr.RFonts.Ascii = {Name: oFontItem.Name, Index: oFontItem.Index};
					rPr.RFonts.HAnsi = {Name: oFontItem.Name, Index: oFontItem.Index};
					rPr.RFonts.CS = {Name: oFontItem.Name, Index: oFontItem.Index};
					rPr.RFonts.EastAsia = {Name: oFontItem.Name, Index: oFontItem.Index};
				}
			}
			var font_size = node.style ? node.style.fontSize : null;
			if (!font_size)
				font_size = this._getStyle(node, computedStyle, "font-size");
			font_size = CheckDefaultFontSize(font_size, this.apiEditor);
			if (font_size) {
				var obj = AscCommon.valueToMmType(font_size);
				if (obj && "%" !== obj.type && "none" !== obj.type) {
					font_size = obj.val;
					//Если браузер не поддерживает нецелые пикселы отсекаем половинные шрифты, они появляются при вставке 8, 11, 14, 20, 26pt
					if ("px" === obj.type && false === this.bIsDoublePx)
						font_size = Math.round(font_size * g_dKoef_mm_to_pt);
					else
						font_size = Math.round(2 * font_size * g_dKoef_mm_to_pt) / 2;//половинные значения допустимы.

					//TODO use constant
					if (font_size > 300)
						font_size = 300;
					else if (font_size === 0)
						font_size = 1;

					rPr.FontSize = font_size;
				}
			}
			var font_weight = this._getStyle(node, computedStyle, "font-weight");
			if (font_weight) {
				if ("bold" === font_weight || "bolder" === font_weight || 400 < font_weight)
					rPr.Bold = true;
			}
			var font_style = this._getStyle(node, computedStyle, "font-style");
			if ("italic" === font_style)
				rPr.Italic = true;
			var color = this._getStyle(node, computedStyle, "color");
			if (color && (color = this._ParseColor(color))) {
				if (PasteElementsId.g_bIsDocumentCopyPaste) {
					rPr.Color = color;
				} else {
					if (color) {
						rPr.Unifill = AscFormat.CreateUnfilFromRGB(color.r, color.g, color.b);
					}
				}
			}

			var spacing = this._getStyle(node, computedStyle, "letter-spacing");
			if (spacing && null != (spacing = AscCommon.valueToMm(spacing)))
				rPr.Spacing = spacing;

			//Провяем те свойства, которые не наследуется, надо смотреть родительские элементы
			var background_color = null;
			var underline = null;
			var Strikeout = null;
			var vertical_align = null;
			var oTempNode = node;
			while (true !== bUseOnlyInherit && true) {
				var tempComputedStyle = this._getComputedStyle(oTempNode);
				if (null == tempComputedStyle)
					break;
				if (null == underline || null == Strikeout) {
					var text_decoration = this._getStyle(node, tempComputedStyle, "text-decoration");
					if (text_decoration) {
						if (null == underline) {
							if (-1 !== text_decoration.indexOf("underline")) {
								underline = true;
							} else if (-1 !== text_decoration.indexOf("none") && node.parentElement && node.parentElement.nodeName.toLowerCase() === "a") {
								underline = false;
							} else if (node.style["text-decoration"] === "none") {
								//for next situation: if element style has text-decoration = "none"
								//computed style gives same settings for all elements(-1!=text_decoration.indexOf("none"))
								//we ignore computed style "none" -> use parent underline style
								underline = false;
							}
						}

						if (null == Strikeout && -1 !== text_decoration.indexOf("line-through"))
							Strikeout = true;
					}
				}
				if (null == background_color) {
					background_color = this._getStyle(node, tempComputedStyle, "background-color");
					if (background_color)
						background_color = this._ParseColor(background_color);
					else
						background_color = null;
				}
				if (null == vertical_align || "baseline" === vertical_align) {
					vertical_align = this._getStyle(node, tempComputedStyle, "vertical-align");
					if (!vertical_align)
						vertical_align = null;
				}
				if (vertical_align && background_color && Strikeout && underline)
					break;
				oTempNode = oTempNode.parentNode;
				let _nodeName = oTempNode && oTempNode.nodeName && oTempNode.nodeName.toLowerCase();
				if (!oTempNode || this.oRootNode === oTempNode || "body" === _nodeName || true === this._IsBlockElem(_nodeName)) {
					break;
				}
			}
			if (PasteElementsId.g_bIsDocumentCopyPaste) {
				if (background_color)
					rPr.HighLight = background_color;
			} else
				delete rPr.HighLight;
			if (null != underline)
				rPr.Underline = underline;
			if (null != Strikeout)
				rPr.Strikeout = Strikeout;
			switch (vertical_align) {
				case "sub":
					rPr.VertAlign = AscCommon.vertalign_SubScript;
					break;
				case "super":
					rPr.VertAlign = AscCommon.vertalign_SuperScript;
					break;
			}
		}
		return rPr;
	},
	_parseCss: function (sStyles, pPr) {
		var aStyles = sStyles.split(';');
		if (aStyles) {
			for (var i = 0, length = aStyles.length; i < length; i++) {
				var style = aStyles[i];
				var aPair = style.split(':');
				if (aPair && aPair.length > 1) {
					var prop_name = trimString(aPair[0]);
					var prop_value = trimString(aPair[1]);
					if (null != this.MsoStyles[prop_name])
						pPr[prop_name] = prop_value;
				}
			}
		}
	},
	_PrepareContent: function (indent) {
		//Не допускам чтобы контент заканчивался на таблицу, иначе тяжело вставить параграф после
		if (this.aContent.length > 0) {
			var last = this.aContent[this.aContent.length - 1];
			if (type_Table === last.GetType()) {
				this._Add_NewParagraph();
			} else if (indent && type_Paragraph === last.GetType()) {
				//при копировании внутри мс indent ячейки записывается в FirstLine (excel->word->xlsx->w:ind->w:firstLine)
				if (last.Pr && last.Pr.Ind) {
					last.Pr.Ind.FirstLine = indent * koef_mm_to_indent;
				}
			}
		}
	},

	_getMsoListSymbol: function (node) {
		var res = null;
		var nodeList = this._getMsoListIgnore(node);
		if (nodeList) {
			var value = nodeList.innerText;
			if (value) {
				for (var pos = value.getUnicodeIterator(); pos.check(); pos.next()) {
					var nUnicode = pos.value();

					if (null !== nUnicode) {
						if (0x20 !== nUnicode && 0xA0 !== nUnicode && 0x2009 !== nUnicode) {
							if (!res) {
								res = "";
							}
							res += value.charAt(pos.position());
						}
					}
				}
			}
		}
		return res;
	},

	_getMsoListIgnore: function (node) {
		if (!node || (node && !node.children)) {
			return null;
		}

		for (var i = 0; i < node.children.length; i++) {
			var child = node.children[i];
			var style = child.getAttribute("style");
			if (style) {
				var pNoHtml = {};
				this._parseCss(style, pNoHtml);
				if ("Ignore" === pNoHtml["mso-list"]) {
					return child;
				}
			}

			if (child.children && child.children.length) {
				return this._getMsoListIgnore(child);
			}
		}
	},
	_getTypeMsoListSymbol: function (str, getStartPosition) {
		var symbolsArr = [
			"ivxlcdm",
			"IVXLCDM",
			"abcdefghijklmnopqrstuvwxyz",
			"ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		];

		//TODO пересмотреть функцию перевода из римских чисел
		var romanToIndex = function (text) {
			var arab_number = [1, 4, 5, 9, 10, 40, 50, 90, 100, 400, 500, 900, 1000, 4000, 5000, 9000, 10000];
			var rom_number = ["I", "IV", "V", "IX", "X", "XL", "L", "XC", "C", "CD", "D", "CM", "M", "M&#8577;", "&#8577;", "&#8577;&#8578;", "&#8578;"];

			var text = text.toUpperCase();
			var result = 0;
			var pos = 0;
			var i = arab_number.length - 1;
			while (i >= 0 && pos < text.length) {
				if (text.substr(pos, rom_number[i].length) === rom_number[i]) {
					result += arab_number[i];
					pos += rom_number[i].length;
				} else {
					i--;
				}
			}
			return result;
		};

		var latinToIndex = function (text) {
			var text = text.toUpperCase();
			var index = 0;
			for (var i = 0; i < text.length; i++) {
				index += symbolsArr[3].indexOf(text[i]) + 1;
			}
			return index;
		};

		var getFullListIndex = function (indexStr, str) {
			var fullListIndex = "";
			for (var i = 0; i < str.length; i++) {
				var symbol = str[i];
				if (-1 !== indexStr.indexOf(symbol)) {
					fullListIndex += symbol;
				} else {
					break;
				}
			}
			return fullListIndex;
		};

		//TODO пока делаю так, пересмотреть регулярные выражения
		var resType = Asc.c_oAscNumberingFormat.Bullet;
		var number = parseInt(str);
		var startPos = null, fullListIndex;
		if (!isNaN(number)) {
			resType = Asc.c_oAscNumberingFormat.Decimal;
			startPos = number;
		} else if (1 === str.length && -1 !== str.indexOf("o")) {
			resType = Asc.c_oAscNumberingFormat.Bullet;
		} else {
			//1)смотрим на первый символ в строке
			//2)ищем все символы, соответсвующие данному типу
			//3)находим порядковый номер этих символов
			var firstSymbol = str[0];
			if (-1 !== symbolsArr[0].indexOf(firstSymbol)) {
				if (getStartPosition) {
					fullListIndex = getFullListIndex(symbolsArr[0], str);
					startPos = romanToIndex(fullListIndex);
				}

				resType = Asc.c_oAscNumberingFormat.LowerRoman;
			} else if (-1 !== symbolsArr[1].indexOf(firstSymbol)) {
				if (getStartPosition) {
					fullListIndex = getFullListIndex(symbolsArr[1], str);
					startPos = romanToIndex(fullListIndex);
				}

				resType = Asc.c_oAscNumberingFormat.UpperRoman;
			} else if (-1 !== symbolsArr[2].indexOf(firstSymbol)) {
				if (getStartPosition) {
					fullListIndex = getFullListIndex(symbolsArr[2], str);
					startPos = latinToIndex(fullListIndex);
				}

				resType = Asc.c_oAscNumberingFormat.LowerLetter;
			} else if (-1 !== symbolsArr[3].indexOf(firstSymbol)) {
				if (getStartPosition) {
					fullListIndex = getFullListIndex(symbolsArr[3], str);
					startPos = latinToIndex(fullListIndex);
				}

				resType = Asc.c_oAscNumberingFormat.UpperLetter;
			}
		}

		return {type: resType, startPos: startPos};
	},
	_tryGenerateNumberingFromMsoStyle: function (aNumbering, node) {
		//TODO mso-level-style-link - пока не парсил ссылку на стиль
		/*@list l0:level1
		{mso-level-style-link:"Р—Р°РіРѕР»РѕРІРѕРє 1";
			mso-level-text:%1;
			mso-level-tab-stop:none;
			mso-level-number-position:left;
			margin-left:.3in;
			text-indent:-.3in;}*/
		//лежит в таком виде:
		/*{mso-style-name:"Р—Р°РіРѕР»РѕРІРѕРє 1";
			mso-style-unhide:no;
			margin-top:0in;
			margin-right:0in;
			margin-bottom:8.0pt;
			margin-left:.3in;
			text-indent:-.3in;.....*/

		// Создаем нумерацию
		var res = null;
		if (aNumbering && aNumbering[1]) {

			var correctText = function (_str) {
				//лежит строка вида - ""\(\0022\\\0027%1sdfdf\0022J\\J\)""

				if (!_str) {
					return "";
				}

				var res = "";
				var isStartSpecSym = null;
				for (var i = 0; i < _str.length; i++) {
					if (_str[i] === "\\" && null === isStartSpecSym) {
						isStartSpecSym = "";
						continue;
					}
					if (_str[i] === '"') {
						continue;
					}
					if (isStartSpecSym !== null) {
						if (window['AscCommon'].isNumber(_str[i])) {
							if (isStartSpecSym.length + 1 === 4) {
								isStartSpecSym += _str[i];
								res += String.fromCharCode("0x" + isStartSpecSym);
								isStartSpecSym = null;
							} else {
								isStartSpecSym += _str[i];
							}
						} else {
							res += _str[i];
							isStartSpecSym = null;
						}
					} else {
						res += _str[i];
					}
				}
				return res;
			};

			var getPropValue = function (val, oStyleLink, oStyle) {
				var res = oStyleLink && oStyleLink[val];
				if (!res) {
					res = oStyle && oStyle[val];
				}
				if (res) {
					res = AscCommon.valueToMmType(res);
				}
				return res ? res.val : null;
			};

			res = this.oLogicDocument.GetNumbering().CreateNum();
			for (var i = 1; i < aNumbering.length; i++) {
				var curNumbering = aNumbering[i];
				if (!curNumbering) {
					continue;
				}

				var sType = curNumbering["mso-level-number-format"];

				var nType = Asc.c_oAscNumberingFormat.None;
				if ("none" === sType) {
					nType = Asc.c_oAscNumberingFormat.None;
				} else if ("bullet" === sType || "image" === sType) {
					nType = Asc.c_oAscNumberingFormat.Bullet;
				} else if ("decimal" === sType) {
					nType = Asc.c_oAscNumberingFormat.Decimal;
				} else if ("roman-lower" === sType) {
					nType = Asc.c_oAscNumberingFormat.LowerRoman;
				} else if ("roman-upper" === sType) {
					nType = Asc.c_oAscNumberingFormat.UpperRoman;
				} else if ("letter-lower" === sType || "alpha-lower" === sType) {
					nType = Asc.c_oAscNumberingFormat.LowerLetter;
				} else if ("letter-upper" === sType || "alpha-upper" === sType) {
					nType = Asc.c_oAscNumberingFormat.UpperLetter;
				} else if ("decimal-zero" === sType) {
					nType = Asc.c_oAscNumberingFormat.DecimalZero;
				} else {
					nType = Asc.c_oAscNumberingFormat.Decimal;
				}

				var sAlign = curNumbering["mso-level-number-position"];
				var nAlign = align_Left;
				if ("left" === sAlign) {
					nAlign = align_Left;
				} else if ("right" === sAlign) {
					nAlign = align_Right;
				} else if ("center" === sAlign) {
					nAlign = align_Center;
				}

				var styleLink = curNumbering["mso-level-style-link"];
				var msoLinkStyles = this._findMsoStyleFromMsoHeadStyle(styleLink);
				var sTextFormatString = curNumbering["mso-level-text"];
				if (nType === Asc.c_oAscNumberingFormat.Bullet) {
					var LvlText = String.fromCharCode(0x00B7);
					var NumTextPr = new CTextPr();

					var fontFamily = curNumbering["font-family"];
					if (fontFamily && -1 !== fontFamily.indexOf("Symbol")) {
						NumTextPr.RFonts.SetAll("Symbol", -1);
						LvlText = sTextFormatString ? sTextFormatString : String.fromCharCode(0x00B7);
					} else if (fontFamily && -1 !== fontFamily.indexOf("Courier New")) {
						NumTextPr.RFonts.SetAll("Courier New", -1);
						LvlText = sTextFormatString ? sTextFormatString : "o";
					} else if (fontFamily && -1 !== fontFamily.indexOf("Wingdings")) {
						NumTextPr.RFonts.SetAll("Wingdings", -1);
						LvlText = sTextFormatString ? sTextFormatString : String.fromCharCode(0x00A7);
					} else {
						NumTextPr.RFonts.SetAll("Symbol", -1);
						LvlText = sTextFormatString ? sTextFormatString : String.fromCharCode(0x00B7);
					}

					res.SetLvlByType(i-1, c_oAscNumberingLevel.Bullet, LvlText, NumTextPr);
				} else {
					if (sTextFormatString) {
						res.SetLvlByFormat(i-1, nType, correctText(sTextFormatString), nAlign);
					} else {
						switch (nType) {
							case Asc.c_oAscNumberingFormat.Decimal    :
								res.SetLvlByType(i-1, c_oAscNumberingLevel.DecimalDot_Left);
								break;
							case Asc.c_oAscNumberingFormat.LowerRoman :
								res.SetLvlByType(i-1, c_oAscNumberingLevel.LowerRomanDot_Right);
								break;
							case Asc.c_oAscNumberingFormat.UpperRoman :
								res.SetLvlByType(i-1, c_oAscNumberingLevel.UpperRomanDot_Right);
								break;
							case Asc.c_oAscNumberingFormat.LowerLetter:
								res.SetLvlByType(i-1, c_oAscNumberingLevel.LowerLetterDot_Left);
								break;
							case Asc.c_oAscNumberingFormat.UpperLetter:
								res.SetLvlByType(i-1, c_oAscNumberingLevel.UpperLetterDot_Left);
								break;
						}
					}
					var startPos = curNumbering["mso-level-start-at"];
					if (null != startPos) {
						res.SetLvlStart(i-1, startPos);
					}
				}
				//res.GetAbstractNum().Lvl[i-1].ParaPr
				var marginLeft = getPropValue("margin-left", msoLinkStyles, curNumbering);
				var firstLine = getPropValue("text-indent", msoLinkStyles, curNumbering);
				if (marginLeft !== null) {
					var newParaPr = new CParaPr();
					if (marginLeft !== null) {
						newParaPr.Ind.Left = marginLeft;
					}
					if (firstLine !== null) {
						newParaPr.Ind.FirstLine = firstLine;
					}
					res.SetParaPr(i-1, newParaPr);
				}

				//text properties
				if (nType !== Asc.c_oAscNumberingFormat.Bullet) {
					var rPr = this._read_rPr_mso_numbering(curNumbering, msoLinkStyles, node);
					if (rPr) {
						res.SetTextPr(i-1, rPr);
					}
				}
			}
		}
		return res;
	},
	_read_rPr_mso_numbering: function (numberingProps, msoLinkStyles, node) {

		//пример того, что должно лежать в numberingProps:
		/*	mso-list-template-ids:-1656974294;}
			@list l0:level1
			{mso-level-style-link:"Heading 11";
			mso-level-suffix:space;
			mso-level-text:"ARTICLE %1 -";
			mso-level-tab-stop:none;
			mso-level-number-position:left;
			margin-left:229.5pt;
			text-indent:0in;
			color:red;
			mso-style-textfill-fill-color:red;
			mso-style-textfill-fill-alpha:100.0%;
			mso-ansi-font-style:italic;
			mso-bidi-font-style:italic;
			text-underline:#000000 single;
			text-decoration:line-through underline;
			vertical-align:super;*/

		//пример того, что должно лежать в msoLinkStyles:
		//ссылка mso-level-style-link:"Heading 11"
		/*{mso-style-name:"Heading 11";
			mso-style-update:auto;
			mso-style-unhide:no;
			mso-style-link:"Heading 1 Char";
			mso-style-next:"Heading 21";
			margin-top:12.0pt;
			margin-right:0in;
			margin-bottom:0in;
			margin-left:0in;
			text-align:center;
			text-indent:0in;
			mso-pagination:widow-orphan;
			page-break-after:avoid;
			mso-list:l0 level1 lfo1;
			font-size:10.0pt;
			font-family:"Arial",sans-serif;
			mso-fareast-font-family:"Times New Roman";
			text-transform:uppercase;
			mso-ansi-language:EN-CA;
			font-weight:bold;
			mso-bidi-font-weight:normal;}*/

		var getProperty = function (name) {
			return (numberingProps && numberingProps[name]) || (msoLinkStyles && msoLinkStyles[name]);
		};

		var rPr = new CTextPr();

		var font_family = getProperty("font-family");
		font_family = font_family && font_family.split(",");
		if (font_family && font_family[0] && "" != font_family[0]) {
			var oFontItem = this.oFonts[font_family[0]];
			if (null != oFontItem && null != oFontItem.Name) {
				rPr.RFonts.Ascii = {Name: oFontItem.Name, Index: oFontItem.Index};
				rPr.RFonts.HAnsi = {Name: oFontItem.Name, Index: oFontItem.Index};
				rPr.RFonts.CS = {Name: oFontItem.Name, Index: oFontItem.Index};
				rPr.RFonts.EastAsia = {Name: oFontItem.Name, Index: oFontItem.Index};
			}
		}

		var font_size = getProperty("font-size");
		if (font_size) {
			font_size = CheckDefaultFontSize(font_size, this.apiEditor);
			if (font_size) {
				var obj = AscCommon.valueToMmType(font_size);
				if (obj && "%" !== obj.type && "none" !== obj.type) {
					font_size = obj.val;
					//Если браузер не поддерживает нецелые пикселы отсекаем половинные шрифты, они появляются при вставке 8, 11, 14, 20, 26pt
					if ("px" === obj.type && false === this.bIsDoublePx)
						font_size = Math.round(font_size * g_dKoef_mm_to_pt);
					else
						font_size = Math.round(2 * font_size * g_dKoef_mm_to_pt) / 2;//половинные значения допустимы.

					//TODO use constant
					if (font_size > 300)
						font_size = 300;
					else if (font_size === 0)
						font_size = 1;

					rPr.FontSize = font_size;
				}
			}
		}

		var font_weight = getProperty("font-weight");
		if (font_weight) {
			if ("bold" === font_weight || "bolder" === font_weight || 400 < font_weight)
				rPr.Bold = true;
		}
		var font_style = getProperty("mso-ansi-font-style");
		if ("italic" === font_style)
			rPr.Italic = true;

		var color = getProperty("color");
		if (color && (color = this._ParseColor(color))) {
			if (PasteElementsId.g_bIsDocumentCopyPaste) {
				rPr.Color = color;
			} else {
				if (color) {
					rPr.Unifill = AscFormat.CreateUnfilFromRGB(color.r, color.g, color.b);
				}
			}
		}

		var spacing = getProperty("letter-spacing");
		if (spacing && null != (spacing = AscCommon.valueToMm(spacing)))
			rPr.Spacing = spacing;

		//Провяем те свойства, которые не наследуется, надо смотреть родительские элементы
		var background_color = null;
		var underline = null;
		var Strikeout = null;
		var vertical_align = null;

		var text_decoration = getProperty("text-decoration");
		if (text_decoration) {
			if (-1 !== text_decoration.indexOf("underline")) {
				underline = true;
			} else if (-1 !== text_decoration.indexOf("none") && node && node.parentElement && node.parentElement.nodeName && node.parentElement.nodeName.toLowerCase() === "a") {
				underline = false;
			}

			if (-1 !== text_decoration.indexOf("line-through")) {
				Strikeout = true;
			}
		}

		background_color = getProperty("background-color");
		if (background_color) {
			background_color = this._ParseColor(background_color);
		}

		vertical_align = getProperty("vertical-align");
		if (!vertical_align) {
			vertical_align = null;
		}

		if (null != underline) {
			rPr.Underline = underline;
		}
		if (null != Strikeout) {
			rPr.Strikeout = Strikeout;
		}
		switch (vertical_align) {
			case "sub":
				rPr.VertAlign = AscCommon.vertalign_SubScript;
				break;
			case "super":
				rPr.VertAlign = AscCommon.vertalign_SuperScript;
				break;
		}

		return rPr;
	},
	_findMsoHeadStyle: function (html) {
		var res;
		var headTag = html && html.getElementsByTagName( "head" );

		if (headTag && headTag[0]) {
			for (var i = 0; i < headTag[0].children.length; ++i) {
				if ("style" === headTag[0].children[i].nodeName.toLowerCase()) {
					if (!res) {
						res = [];
					}
					res.push(headTag[0].children[i].innerText);
				}
			}
		}

		return res;
	},
	_findElemFromMsoHeadStyle: function (prefixName, name) {
		//первый элемент массива - общая инфомарция о списках, остальные - инфомарция об уровнях по порядку
		/*@list l0
		{mso-list-id:1405642429;
			mso-list-type:hybrid;
			mso-list-template-ids:1269742048 682115302 67698713 67698715 67698703 67698713 67698715 67698703 67698713 67698715;}
		@list l0:level1
		{mso-level-number-format:roman-lower;
			mso-level-text:1iii2%1eeiiiFFF;
			mso-level-tab-stop:none;
			mso-level-number-position:left;
			text-indent:-.25in;}
		@list l0:level2
		{mso-level-number-format:alpha-lower;
			mso-level-tab-stop:none;
			mso-level-number-position:left;
			text-indent:-.25in;}*/

		var res = null;
		var fullName = prefixName +" " + name;
		if (this.aMsoHeadStylesStr) {
			var _generateIndexByPrefix = function (_str) {
				if (0 === _str.indexOf(fullName)) {
					_index = 0;
					var _fullStr = fullName + ":level" ;
					if (-1 !== _str.indexOf(_fullStr)) {
						var level = _str.split(_fullStr);
						if (level && level[1]) {
							_index = parseInt(level[1]);
							if (isNaN(_index)) {
								_index = undefined;
							}
						}
					}
				} else {
					_index = undefined;
				}
			};

			var _pushStr = function (key, val) {
				if (val === null || key === null || _index === undefined) {
					return;
				}
				if (!res) {
					res = [];
				}
				if (!res[_index]) {
					res[_index] = {};
				}

				res[_index][key] = val;
			};
			var _index, prefix = "";
			for (var i = 0; i < this.aMsoHeadStylesStr.length; i++) {
				var pos = this.aMsoHeadStylesStr[i].indexOf(prefixName +" " + name);
				if (pos !== -1) {
					var startObjStr = null;
					var startKeyStr = null;
					var startValStr = null;
					for (var j = pos; j < this.aMsoHeadStylesStr[i].length; j++) {
						var sym = this.aMsoHeadStylesStr[i][j];
						if (sym === "\n" || sym === "\t") {
							continue;
						}
						if (!startObjStr) {
							if (sym === "{") {
								startObjStr = true;
								_generateIndexByPrefix(prefix);
								prefix = "";
							} else {
								prefix += sym;
							}

							startKeyStr = "";
							startValStr = null;
						} else {
							if (sym === "}") {
								_pushStr(startKeyStr, startValStr);
								startObjStr = null;
							} else if (sym === ":") {
								startValStr = "";
							} else if (sym === ";") {
								_pushStr(startKeyStr, startValStr);
								startValStr = null;
								startKeyStr = "";
							} else if (startValStr !== null) {
								startValStr += sym;
							} else if (startKeyStr !== null) {
								startKeyStr += sym;
							}
						}
					}
				}
			}
		}
		return res;
	},
	_findMsoStyleFromMsoHeadStyle: function (name) {

		var res = null;
		if (this.aMsoHeadStylesStr) {

			var _pushStr = function (key, val) {
				if (val === null || key === null) {
					return;
				}
				if (!res) {
					res = [];
				}
				res[key] = val;
			};

			var searchStr = "mso-style-name:" + name;
			for (var i = 0; i < this.aMsoHeadStylesStr.length; i++) {
				var startPos = this.aMsoHeadStylesStr[i].indexOf(searchStr);

				if (startPos !== -1) {
					var startKeyStr = "";
					var startValStr = null;
					for (var j = startPos; j < this.aMsoHeadStylesStr[i].length; j++) {
						var sym = this.aMsoHeadStylesStr[i][j];
						if (sym === "\n" || sym === "\t") {
							continue;
						}
						if (sym === "}") {
							_pushStr(startKeyStr, startValStr);
							break;
						} else if (sym === ":") {
							startValStr = "";
						} else if (sym === ";") {
							_pushStr(startKeyStr, startValStr);
							startValStr = null;
							startKeyStr = "";
						} else if (startValStr !== null) {
							startValStr += sym;
						} else if (startKeyStr !== null) {
							startKeyStr += sym;
						}
					}
				}

			}
		}
		return res;
	},
	_AddNextPrevToContent: function (oDoc) {
		var prev = null;
		for (var i = 0, length = this.aContent.length; i < length; ++i) {
			var cur = this.aContent[i];
			cur.Set_DocumentPrev(prev);
			cur.Parent = oDoc;
			if (prev)
				prev.Set_DocumentNext(cur);
			prev = cur;
		}
	},
	_Set_Run_Pr: function (oPr) {
		this._CommitRunToParagraph(false);
		if (null != this.oCurRun) {
			this.oCurRun.Set_Pr(oPr);
		}
	},
	_CommitRunToParagraph: function (bCreateNew) {
		if (bCreateNew || this.oCurRun.Content.length > 0) {
			this.oCurRun = new ParaRun(this.oCurPar);
			this.oCurRunContentPos = 0;
		}
	},
	_CommitElemToParagraph: function (elem) {
		if (null != this.oCurHyperlink) {
			this.oCurHyperlink.Add_ToContent(this.oCurHyperlinkContentPos, elem, false);
			this.oCurHyperlinkContentPos++;
		} else {
			this.oCurPar.Internal_Content_Add(this.oCurParContentPos, elem, false);
			this.oCurParContentPos++;
		}
	},
	_AddToParagraph: function (elem) {
		if (null != this.oCurRun) {
			if (para_Hyperlink === elem.Type) {
				this._CommitRunToParagraph(true);
				this._CommitElemToParagraph(elem);
			} else {
				if (this.oCurRun.Content.length === Asc.c_dMaxParaRunContentLength) {
					//создаём новый paraRun и выставляем ему настройки предыдущего
					//сделано для того, чтобы избежать большого количества данных в paraRun
					if (this.oCurRun && this.oCurRun.Pr && this.oCurRun.Pr.Copy) {
						this._Set_Run_Pr(this.oCurRun.Pr.Copy());
					}
				}

				this.oCurRun.Add_ToContent(this.oCurRunContentPos, elem, false);
				this.oCurRunContentPos++;
				if (1 === this.oCurRun.Content.length)
					this._CommitElemToParagraph(this.oCurRun);
			}
		}
	},
	_Add_NewParagraph: function () {
		this.oCurPar = new Paragraph(this.oDocument.DrawingDocument, this.oDocument, this.oDocument.bPresentation === true);
		this.oCurParContentPos = this.oCurPar.CurPos.ContentPos;
		this.oCurRun = new ParaRun(this.oCurPar);
		this.oCurRunContentPos = 0;
		this.aContent.push(this.oCurPar);
		//сбрасываем настройки теста
		this.oCur_rPr = new CTextPr();
	},
	_Execute_AddParagraph: function (node, pPr) {
		this._Add_NewParagraph();
		//Устанавливаем стили параграфа
		this._set_pPr(node, this.oCurPar, pPr);
	},
	_Decide_AddParagraph: function (node, pPr, bParagraphAdded, bCommitBr) {
		//Игнорируем пустые параграфы(как браузеры, как MS), добавляем параграф только когда придет текст
		if (true == bParagraphAdded) {
			if (false != bCommitBr)
				this._Commit_Br(2, node, pPr);//word игнорируем 2 последних br
			this._Execute_AddParagraph(node, pPr);
		} else if (false != bCommitBr)
			this._Commit_Br(0, node, pPr);
		return false;
	},
	_Commit_Br: function (nIgnore, node, pPr) {
		for (var i = 0, length = this.nBrCount - nIgnore; i < length; i++) {
			if ("always" === pPr["mso-column-break-before"])
				this._AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));
			else {
				if (this.bInBlock)
					this._AddToParagraph(new AscWord.CRunBreak(AscWord.break_Line));
				else
					this._Execute_AddParagraph(node, pPr);
			}
		}
		this.nBrCount = 0;
	},

	_StartExecuteTable: function (node, pPr, arrShapes, arrImages, arrTables) {
		var oDocument = this.oDocument;
		var tableNode = node, newNode, headNode;
		var bPresentation = !PasteElementsId.g_bIsDocumentCopyPaste;

		//Ищем если есть tbody
		var i, length, j, length2;
		for (i = 0, length = node.childNodes.length; i < length; ++i) {
			var nodeName = node.childNodes[i].nodeName.toLowerCase();
			if ("tbody" === nodeName) {
				if (!newNode) {
					newNode = node.childNodes[i];
					if (headNode) {
						for (j = 0; j < headNode.childNodes.length; j++) {
							newNode.insertBefore(headNode.childNodes[0], newNode.childNodes[0]);
						}
						pPr.repeatHeaderRow = true;
					}
				} else {
					var lengthChild = node.childNodes[i].childNodes.length;
					for (j = 0; j < lengthChild; j++) {
						newNode.appendChild(node.childNodes[i].childNodes[0]);
					}
				}
			} else if ("thead" === nodeName) {
				headNode = node.childNodes[i];
			}
		}

		if (newNode) {
			node = newNode;
			tableNode = newNode;
		} else if (headNode) {
			node = headNode;
			//tableNode = headNode;
			//pPr.repeatHeaderRow = true;
		}

		//валидация талиц. В таблице не может быть строк состоящих из вертикально замерженых ячеек.
		var nRowCount = 0;
		var nMinColCount = 0;
		var nMaxColCount = 0;
		var aColsCountByRow = [];
		var oRowSums = {};
		oRowSums[0] = 0;
		var dMaxSum = 0;
		var nCurColWidth = 0;
		var nCurSum = 0;
		var nAllSum = 0;
		var oRowSpans = {};
		var columnSize;
		if (!bPresentation && ((!window["Asc"] || !(Asc["editor"] && Asc["editor"].wb))) && this.oLogicDocument) {
			columnSize = this.oLogicDocument.GetColumnSize();
		}
		var fParseSpans = function () {
			var spans = oRowSpans[nCurColWidth];
			while (null != spans && spans.row > 0) {
				spans.row--;
				nCurColWidth += spans.col;
				nCurSum += spans.width;
				spans = oRowSpans[nCurColWidth];
			}
		};

		var tc, tcName, nCurRowSpan;
		for (i = 0, length = node.childNodes.length; i < length; ++i) {
			var tr = node.childNodes[i];
			if ("tr" === tr.nodeName.toLowerCase()) {
				nCurSum = 0;
				nCurColWidth = 0;
				var minRowSpanIndex = null;
				var nMinRowSpanCount = null;//минимальный rowspan ячеек строки
				for (j = 0, length2 = tr.childNodes.length; j < length2; ++j) {
					tc = tr.childNodes[j];
					tcName = tc.nodeName.toLowerCase();
					if ("td" === tcName || "th" === tcName) {
						fParseSpans();

						var dWidth = null;
						var computedStyle = this._getComputedStyle(tc);
						var computedWidth = this._getStyle(tc, computedStyle, "width");
						if (null != computedWidth && null != (computedWidth = AscCommon.valueToMm(computedWidth)))
							dWidth = computedWidth;

						if (null == dWidth)
							dWidth = tc.clientWidth * g_dKoef_pix_to_mm;

						var nColSpan = tc.getAttribute("colspan");
						if (null != nColSpan)
							nColSpan = nColSpan - 0;
						else
							nColSpan = 1;
						nCurRowSpan = tc.getAttribute("rowspan");
						if (null != nCurRowSpan) {
							nCurRowSpan = nCurRowSpan - 0;
							if (null == nMinRowSpanCount) {
								nMinRowSpanCount = nCurRowSpan;
								minRowSpanIndex = j;
							} else if (nMinRowSpanCount > nCurRowSpan) {
								nMinRowSpanCount = nCurRowSpan;
								minRowSpanIndex = j;
							}

							if (nCurRowSpan > 1)
								oRowSpans[nCurColWidth] = {row: nCurRowSpan - 1, col: nColSpan, width: dWidth};
						} else {
							nMinRowSpanCount = 0;
							minRowSpanIndex = j;
						}

						nCurSum += dWidth;
						if (null == oRowSums[nCurColWidth + nColSpan]) {
							oRowSums[nCurColWidth + nColSpan] = nCurSum;
						} else if (null != oRowSums[nCurColWidth + nColSpan - 1] && oRowSums[nCurColWidth + nColSpan - 1] >= oRowSums[nCurColWidth + nColSpan] && dWidth !== 0) {
							oRowSums[nCurColWidth + nColSpan] += nCurSum;
						}
						nCurColWidth += nColSpan;
					}
				}
				nAllSum += nCurSum;
				fParseSpans();
				//Удаляем лишние rowspan
				if (nMinRowSpanCount > 1) {
					for (j = 0, length2 = tr.childNodes.length; j < length2; ++j) {
						tc = tr.childNodes[j];
						tcName = tc.nodeName.toLowerCase();
						if (minRowSpanIndex !== j && ("td" === tcName || "th" === tcName)) {
							nCurRowSpan = tc.getAttribute("rowspan");
							if (null != nCurRowSpan)
								tc.setAttribute("rowspan", nCurRowSpan - nMinRowSpanCount);
						}
					}
				}
				if (dMaxSum < nCurSum)
					dMaxSum = nCurSum;
				//удаляем пустые tr
				if (0 === nCurColWidth) {
					node.removeChild(tr);
					length--;
					i--;
				} else {
					if (0 === nMinColCount || nMinColCount > nCurColWidth) {
						nMinColCount = nCurColWidth;
					}
					if (nMaxColCount < nCurColWidth) {
						nMaxColCount = nCurColWidth;
					}
					nRowCount++;
					aColsCountByRow.push(nCurColWidth);
				}
			}
		}
		if (nMaxColCount !== nMinColCount) {
			for (i = 0, length = aColsCountByRow.length; i < length; ++i) {
				aColsCountByRow[i] = nMaxColCount - aColsCountByRow[i];
			}
		}
		if (nRowCount > 0 && nMaxColCount > 0) {
			var bUseScaleKoef = this.bUseScaleKoef;
			var dScaleKoef = this.dScaleKoef;
			if (dMaxSum * dScaleKoef > this.dMaxWidth) {
				dScaleKoef = dScaleKoef * this.dMaxWidth / dMaxSum;
				bUseScaleKoef = true;
			}
			//строим Grid
			var aGrid = [];
			var nPrevIndex = null;
			var nPrevVal = 0;
			for (i in oRowSums) {
				var nCurIndex = i - 0;
				var nCurVal = oRowSums[i];
				var nCurWidth = nCurVal - nPrevVal;
				if (bUseScaleKoef)
					nCurWidth *= dScaleKoef;
				if (null != nPrevIndex) {
					var nDif = nCurIndex - nPrevIndex;
					if (1 === nDif) {
						if (!nCurWidth && !nAllSum && columnSize) {
							aGrid.push(columnSize.W / nMaxColCount);
						} else {
							aGrid.push(nCurWidth);
						}
					} else {
						var nPartVal = nCurWidth / nDif;
						for (j = 0; j < nDif; ++j)
							aGrid.push(nPartVal);
					}
				}
				nPrevVal = nCurVal;
				nPrevIndex = nCurIndex;
			}

			var table;
			if (bPresentation) {
				table = this._createNewPresentationTable(aGrid);
				var graphicFrame = table.Parent;
				table.Set_TableStyle(0);
				arrTables.push(graphicFrame);

				//TODO пересмотреть!!!
				//graphicFrame.setXfrm(dd.GetMMPerDot(node["offsetLeft"]), dd.GetMMPerDot(node["offsetTop"]), dd.GetMMPerDot(node["offsetWidth"]), dd.GetMMPerDot(node["offsetHeight"]), null, null, null);
			} else {
				table = new CTable(oDocument.DrawingDocument, oDocument, true, 0, 0, aGrid);
			}


			//считаем aSumGrid
			var aSumGrid = [];
			aSumGrid[-1] = 0;
			var nSum = 0;
			for (i = 0, length = aGrid.length; i < length; ++i) {
				nSum += aGrid[i];
				aSumGrid[i] = nSum;
			}
			//набиваем content
			this._ExecuteTable(tableNode, node, table, aSumGrid, nMaxColCount !== nMinColCount ? aColsCountByRow : null, pPr, bUseScaleKoef, dScaleKoef, arrShapes, arrImages, arrTables);
			table.MoveCursorToStartPos();

			if (!bPresentation) {
				this.aContent.push(table);
			}
		}
	},

	_ExecuteBlockLevelStd : function (node, pPr) {

		//1. plain text (CInlineLevelSdt)
// 		<body lang=EN-US style='tab-interval:.5in;word-wrap:break-word'>
// 		<!--StartFragment-->
// 		<span>
// 		<w:Sdt DocPart="AD334DF7FD99465099BBBD9F3C3AE5AE" ID="-818574957">
// 			<span style='mso-spacerun:yes'>В </span>
// 			Plaein text
// 		</w:Sdt>
// </span>
// 		<!--EndFragment-->
// 		</body>


		//2. Rich text (CBlockLevelSdt)
		//
		// <body lang=EN-US style='tab-interval:.5in;word-wrap:break-word'>
		// <!--StartFragment-->
		// <w:Sdt DocPart="5AAA70241BB24C68A16AC32AB3233FFC" ID="819472493">
		// 	<p className=MsoNormal>Rich plain textf
		// 		<o:p/>
		// 		<w:sdtPr/>
		// 	</p>
		// </w:Sdt>
		// <!--EndFragment-->
		// </body>

		//3.picture(CInlineLevelSdt)

		// <p class=MsoNormal>
		// 	<w:Sdt ShowingPlcHdr="t" DocPart="6BE2DA46972A420788E197A3986DD1A7" DisplayAsPicture="t" Text="t" ID="575168600">
		// 	</span>
		// </w:Sdt>
		// <o:p/>
		// </p>

		//4. Choose (CInlineLevelSdt)

// 		<body lang=EN-US style='tab-interval:.5in;word-wrap:break-word'>
// 		<!--StartFragment-->
// 		<span>
// 		<w:Sdt ShowingPlcHdr="t" DocPart="F7011D2027BE4F6C919D5BB7D95A0DC0" DropDown="t" ID="1536853350">
// 			<w:ListItem ListValue="Choose an item" DataValue=""/>Choose an item
// 		</w:Sdt>
// 	</span>
// 		<!--EndFragment-->
// 		</body>


		//5. Date (CInlineLevelSdt)

// 		<body lang=EN-US style='tab-interval:.5in;word-wrap:break-word'>
// 		<!--StartFragment-->
// 		<span>
// 		<w:Sdt DocPart="0D4FD865761947FCBA0D9229E17016DB" Calendar="t" MapToDateTime="t" CalendarType="Gregorian" Date="2022-10-24T20:27:00Z" DateFormat="dd.MM.yyyy" Lang="EN-US" ID="-291673853">24.10.2022</w:Sdt>
// 	</span>
// 		<!--EndFragment-->
// 		</body>


		//6. check box (CInlineLevelSdt)
//
// 		<body lang=EN-US style='tab-interval:.5in;word-wrap:break-word'>
// 		<!--StartFragment-->
// 		<span>
// 		<w:Sdt CheckBox="t" CheckBoxIsChecked="t" CheckBoxValueChecked="в’" CheckBoxValueUnchecked="вђ" CheckBoxFontChecked="MS Gothic" CheckBoxFontUnchecked="MS Gothic" ID="-213812918">
// 			<span style='font-family:
//  "MS Gothic"'>в’</span>
// 		</w:Sdt>
// </span>
// 		<!--EndFragment-->
// 		</body>




		//w:Sdt -> ориентируемся по первому внутреннему тегу
		//если параграф - то оборачиваем в CBlockLevelSdt, если текст с настройками - в CInlineLevelSdt
		let isBlockLevelSdt = node.getElementsByTagName("p").length > 0;
		let levelSdt = isBlockLevelSdt ? new CBlockLevelSdt(this.oLogicDocument, this.oDocument) : new CInlineLevelSdt();

		//ms в буфер записывает только lock контента
		let checkBox, dropdown, comboBox;
		if (node && node.attributes) {
			let contentLocked = node.attributes["contentlocked"];
			if (contentLocked /*&& contentLocked.value === "t"*/) {
				levelSdt.SetContentControlLock(c_oAscSdtLockType.SdtContentLocked);
			}
			//далее тег и титульник, цвета нет
			let alias = node.attributes["title"];
			if (alias && alias.value) {
				levelSdt.SetAlias(alias.value);
			}
			let tag = node.attributes["sdttag"];
			if (tag && tag.value) {
				levelSdt.SetTag(tag.value);
			}
			let temporary = node.attributes["temporary"];
			if (temporary && temporary.value) {
				levelSdt.SetContentControlTemporary(temporary.value === "t");
			}

			let placeHolder = node.attributes["showingplchdr"];
			if (placeHolder && placeHolder.value === "t") {
				levelSdt.SetPlaceholder(c_oAscDefaultPlaceholderName.Text);
			}

			//TODO поддержать Picture CC
			/*let aspicture = node.attributes["displayaspicture"];
			if (aspicture) {
				blockLevelSdt.SetPicturePr(aspicture.value === "t");
			}*/


			let getCharCode = function (text) {
				let charCode;
				for (let oIterator = text.getUnicodeIterator(); oIterator.check(); oIterator.next()) {
					charCode = oIterator.value();
				}
				return charCode;
			};

			let setListItems = function (addFunc) {
				for (let i = 0, length = node.childNodes.length; i < length; i++) {
					let child = node.childNodes[i];
					if (child && "w:listitem" === child.nodeName.toLowerCase()) {
						if (child.attributes) {
							let listvalue = child.attributes["listvalue"] && child.attributes["listvalue"].value;
							let datavalue = child.attributes["datavalue"] && child.attributes["datavalue"].value;
							if (datavalue) {
								addFunc(datavalue, listvalue ? listvalue : undefined);
							}
						}
					}
				}
			};

			let oPr;
			checkBox = node.attributes["checkbox"];
			if (checkBox && checkBox.value === "t") {
				oPr = new AscWord.CSdtCheckBoxPr();
				let checked = node.attributes["checkboxischecked"];
				if (checked) {
					oPr.Checked = checked.value === "t";
				}
				let checkedFont = node.attributes["checkboxfontchecked"];
				if (checkedFont) {
					oPr.CheckedFont = checkedFont.value;
				}
				let checkedSymbol = node.attributes["checkboxvaluechecked"];
				if (checkedSymbol) {
					oPr.CheckedSymbol = getCharCode(checkedSymbol.value);
				}
				let uncheckedFont = node.attributes["checkboxfontunchecked"];
				if (uncheckedFont) {
					oPr.UncheckedFont = uncheckedFont.value;
				}
				let uncheckedSymbol = node.attributes["checkboxvalueunchecked"];
				if (checkedSymbol) {
					oPr.UncheckedSymbol = getCharCode(uncheckedSymbol.value);
				}

				levelSdt.ApplyCheckBoxPr(oPr);
			}

			let id = node.attributes["id"];
			if (id) {
				levelSdt.Pr.Id = id;
			}

			comboBox = node.attributes["combobox"];
			if (comboBox && comboBox.value === "t") {
				oPr = new AscWord.CSdtComboBoxPr();
			}

			dropdown = node.attributes["dropdown"];
			if (dropdown && dropdown.value === "t") {
				oPr = new AscWord.CSdtComboBoxPr();
			}

			if (comboBox || dropdown) {
				//TODO по грамотному нужно пропарсить всё содержимое и отдать все свойства
				//пока только для listitem делаю
				setListItems(function (sDisplay, sValue) {
					oPr.AddItem(sDisplay, sValue);
				});
				levelSdt.ApplyComboBoxPr(oPr);
			}
		}

		//content
		if (!checkBox && !comboBox && !dropdown) {
			let oPasteProcessor = new PasteProcessor(this.api, false, false, true);
			oPasteProcessor.AddedFootEndNotes = this.AddedFootEndNotes;
			oPasteProcessor.msoComments = this.msoComments;
			oPasteProcessor.oFonts = this.oFonts;
			oPasteProcessor.oImages = this.oImages;
			oPasteProcessor.bIsForFootEndnote = this.bIsForFootEndnote;
			oPasteProcessor.oDocument = isBlockLevelSdt ? levelSdt.Content : this.oDocument;
			oPasteProcessor.bIgnoreNoBlockText = true;
			oPasteProcessor.aMsoHeadStylesStr = this.aMsoHeadStylesStr;
			oPasteProcessor.oMsoHeadStylesListMap = this.oMsoHeadStylesListMap;
			oPasteProcessor.msoListMap = this.msoListMap;
			oPasteProcessor.pasteInExcel = this.pasteInExcel;
			oPasteProcessor._Execute(node, pPr, true, true, false);
			oPasteProcessor._PrepareContent();
			oPasteProcessor._AddNextPrevToContent(levelSdt.Content);

			//добавляем новый параграфы
			let i, j, length, length2;
			if (isBlockLevelSdt) {
				for (i = 0, length = oPasteProcessor.aContent.length; i < length; ++i) {
					if (i === length - 1) {
						levelSdt.Content.Internal_Content_Add(i + 1, oPasteProcessor.aContent[i], true);
					} else {
						levelSdt.Content.Internal_Content_Add(i + 1, oPasteProcessor.aContent[i], false);
					}
				}
				levelSdt.Content.Internal_Content_Remove(0, 1);
			} else {
				for (i = 0, length = oPasteProcessor.aContent.length; i < length; ++i) {
					if (oPasteProcessor.aContent[i] && oPasteProcessor.aContent[i].Content) {
						for (j = 0, length2 = oPasteProcessor.aContent[i].Content.length - 1; j < length2; ++j) {
							levelSdt.AddToContent(j, oPasteProcessor.aContent[i].Content[j]);
						}
					}
				}
			}
		}

		if (isBlockLevelSdt) {
			this.aContent.push(levelSdt);
		} else {
			this._CommitElemToParagraph(levelSdt);
			this.oCur_rPr = new CTextPr();
		}
	},

	_ExecuteBorder: function (computedStyle, node, type, type2, bAddIfNull, setUnifill) {
		var res = null;
		var style = this._getStyle(node, computedStyle, "border-" + type + "-style");
		if (null != style) {
			res = new CDocumentBorder();
			if ("none" === style || "" === style)
				res.Value = border_None;
			else {
				res.Value = border_Single;
				var width = node.style["border" + type2 + "Width"];
				if (!width)
					width = this._getStyle(node, computedStyle, "border-" + type + "-width");
				if (null != width && null != (width = AscCommon.valueToMm(width)))
					res.Size = width;
				var color = this._getStyle(node, computedStyle, "border-" + type + "-color");
				if (null != color && (color = this._ParseColor(color))) {
					if (setUnifill && color) {
						res.Unifill = AscFormat.CreateSolidFillRGB(color.r, color.g, color.b);
					} else {
						res.Color = color;
					}
				}
			}
		}
		if (bAddIfNull && null == res)
			res = new CDocumentBorder();
		return res;
	},
	_ExecuteParagraphBorder: function (border) {
		var res = this.oBorderCache[border];
		if (null != res)
			return res.Copy();
		else {
			//сделано через dom чтобы не писать большую функцию разбора строки
			//todo сделать без dom, анализируя текст.
			res = new CDocumentBorder();
			var oTestDiv = document.createElement("div");
			oTestDiv.setAttribute("style", "border-left:" + border);
			document.body.appendChild(oTestDiv);
			var computedStyle = this._getComputedStyle(oTestDiv);
			if (null != computedStyle) {
				res = this._ExecuteBorder(computedStyle, oTestDiv, "left", "Left", true);
			}
			document.body.removeChild(oTestDiv);
			this.oBorderCache[border] = res;
			return res;
		}
	},

	_ExecuteTable: function (tableNode, node, table, aSumGrid, aColsCountByRow, pPr, bUseScaleKoef, dScaleKoef, arrShapes, arrImages, arrTables) {
		var bPresentation = !PasteElementsId.g_bIsDocumentCopyPaste;

		table.SetTableLayout(tbllayout_AutoFit);
		//Pr
		var Pr = table.Pr;
		//align смотрим у parent tableNode
		var sTableAlign = null;
		if (null != tableNode.align)
			sTableAlign = tableNode.align;
		else if (null != tableNode.parentNode && this.oRootNode !== tableNode.parentNode) {
			var computedStyleParent = this._getComputedStyle(tableNode.parentNode);
			sTableAlign = this._getStyle(tableNode.parentNode, computedStyleParent, "text-align");
		}
		if (null != sTableAlign) {
			if (-1 !== sTableAlign.indexOf('center'))
				table.Set_TableAlign(align_Center);
			else if (-1 !== sTableAlign.indexOf('right'))
				table.Set_TableAlign(align_Right);
		}
		var spacing = null;
		table.Set_TableBorder_InsideH(new CDocumentBorder());
		table.Set_TableBorder_InsideV(new CDocumentBorder());

		var style = tableNode.getAttribute("style");
		if (style) {
			var tblPrMso = {};
			this._parseCss(style, tblPrMso);
			var spacing = tblPrMso["mso-cellspacing"];
			if (null != spacing && null != (spacing = AscCommon.valueToMm(spacing)))
				;
			var padding = tblPrMso["mso-padding-alt"];
			if (null != padding) {
				padding = trimString(padding);
				var aMargins = padding.split(" ");
				if (4 === aMargins.length) {
					var top = aMargins[0];
					if (null != top && null != (top = AscCommon.valueToMm(top)))
						;
					else
						top = Pr.TableCellMar.Top.W;
					var right = aMargins[1];
					if (null != right && null != (right = AscCommon.valueToMm(right)))
						;
					else
						right = Pr.TableCellMar.Right.W;
					var bottom = aMargins[2];
					if (null != bottom && null != (bottom = AscCommon.valueToMm(bottom)))
						;
					else
						bottom = Pr.TableCellMar.Bottom.W;
					var left = aMargins[3];
					if (null != left && null != (left = AscCommon.valueToMm(left)))
						;
					else
						left = Pr.TableCellMar.Left.W;
					table.Set_TableCellMar(left, top, right, bottom);
				}
			}
			var insideh = tblPrMso["mso-border-insideh"];
			if (null != insideh)
				table.Set_TableBorder_InsideH(this._ExecuteParagraphBorder(insideh));
			var insidev = tblPrMso["mso-border-insidev"];
			if (null != insidev)
				table.Set_TableBorder_InsideV(this._ExecuteParagraphBorder(insidev));
		}
		var computedStyle = this._getComputedStyle(tableNode);
		if (align_Left === table.Get_TableAlign()) {
			var margin_left = this._getStyle(tableNode, computedStyle, "margin-left");
			//todo возможно надо еще учесть ширину таблицы
			if (margin_left && null != (margin_left = AscCommon.valueToMm(margin_left)) && margin_left < Page_Width - X_Left_Margin)
				table.Set_TableInd(margin_left);
		}
		var background_color = this._getStyle(tableNode, computedStyle, "background-color");
		if (null != background_color && (background_color = this._ParseColor(background_color)))
			table.Set_TableShd(c_oAscShdClear, background_color.r, background_color.g, background_color.b);
		var oLeftBorder = this._ExecuteBorder(computedStyle, tableNode, "left", "Left", bPresentation);
		if (null != oLeftBorder)
			table.Set_TableBorder_Left(oLeftBorder);
		var oTopBorder = this._ExecuteBorder(computedStyle, tableNode, "top", "Top", bPresentation);
		if (null != oTopBorder)
			table.Set_TableBorder_Top(oTopBorder);
		var oRightBorder = this._ExecuteBorder(computedStyle, tableNode, "right", "Right", bPresentation);
		if (null != oRightBorder)
			table.Set_TableBorder_Right(oRightBorder);
		var oBottomBorder = this._ExecuteBorder(computedStyle, tableNode, "bottom", "Bottom", bPresentation);
		if (null != oBottomBorder)
			table.Set_TableBorder_Bottom(oBottomBorder);

		if (null == spacing) {
			spacing = this._getStyle(tableNode, computedStyle, "padding");
			if (!spacing)
				spacing = tableNode.style.padding;
			if (!spacing)
				spacing = null;
			if (spacing && null != (spacing = AscCommon.valueToMm(spacing)))
				;
		}

		//content
		var oRowSpans = {};
		var bFirstRow = true;
		for (var i = 0, length = node.childNodes.length; i < length; ++i) {
			var tr = node.childNodes[i];
			//TODO временная правка в условии для того, чтобы избежать ошибки при копировании из excel мерженной ячейки
			if ("tr" === tr.nodeName.toLowerCase() && tr.childNodes && tr.childNodes.length) {
				var row = table.private_AddRow(table.Content.length, 0);
				if (bFirstRow && pPr.repeatHeaderRow) {
					row.Pr.TableHeader = true;
				}
				bFirstRow = false;
				this._ExecuteTableRow(tr, row, aSumGrid, spacing, oRowSpans, bUseScaleKoef, dScaleKoef, arrShapes, arrImages, arrTables);
			}
		}
	},

	_ExecuteTableRow: function (node, row, aSumGrid, spacing, oRowSpans, bUseScaleKoef, dScaleKoef, arrShapes, arrImages, arrTables) {
		var oThis = this;
		var table = row.Table;
		var oTableSpacingMinValue = ("undefined" !== typeof tableSpacingMinValue) ? tableSpacingMinValue : 0.02;
		if (null != spacing && spacing >= oTableSpacingMinValue)
			row.Set_CellSpacing(spacing);
		if (node.style.height) {
			var height = node.style.height;
			if (!("auto" === height || "inherit" === height || -1 !== height.indexOf("%")) && null != (height = AscCommon.valueToMm(height)))
				row.Set_Height(height, Asc.linerule_AtLeast);
		}
		var bBefore = false;
		var bAfter = false;
		var style = node.getAttribute("style");
		if (null != style) {
			var tcPr = {};
			this._parseCss(style, tcPr);
			var margin_left = tcPr["mso-row-margin-left"];
			if (margin_left && null != (margin_left = AscCommon.valueToMm(margin_left)))
				bBefore = true;
			var margin_right = tcPr["mso-row-margin-right"];
			if (margin_right && null != (margin_right = AscCommon.valueToMm(margin_right)))
				bAfter = true;
		}

		//content
		var nCellIndex = 0;
		var nCellIndexSpan = 0;
		var fParseSpans = function () {
			var spans = oRowSpans[nCellIndexSpan];
			while (null != spans) {
				var oCurCell = row.Add_Cell(row.Get_CellsCount(), row, null, false);
				if (spans.cell && spans.cell.Pr && spans.cell.Pr.TableCellBorders) {
					//copy props from main cell
					//TODO other options
					let tableCellBorders = spans.cell.Pr.TableCellBorders;
					let border = tableCellBorders.Left;
					if (null != border) {
						oCurCell.Set_Border(border, 3);
					}
					border = tableCellBorders.Top;
					if (null != border) {
						oCurCell.Set_Border(border, 0);
					}
					border = tableCellBorders.Right;
					if (null != border) {
						oCurCell.Set_Border(border, 1);
					}
					border = tableCellBorders.Bottom;
					if (null != border) {
						oCurCell.Set_Border(border, 2);
					}
				}
				oCurCell.SetVMerge(vmerge_Continue);
				if (spans.col > 1)
					oCurCell.Set_GridSpan(spans.col);
				spans.row--;
				if (spans.row <= 0)
					delete oRowSpans[nCellIndexSpan];
				nCellIndexSpan += spans.col;
				spans = oRowSpans[nCellIndexSpan];
			}
		};
		var oBeforeCell = null;
		var oAfterCell = null;
		if (bBefore || bAfter) {
			for (var i = 0, length = node.childNodes.length; i < length; ++i) {
				var tc = node.childNodes[i];
				var tcName = tc.nodeName.toLowerCase();
				if ("td" === tcName || "th" === tcName) {
					if (bBefore && null != oBeforeCell)
						oBeforeCell = tc;
					else if (bAfter)
						oAfterCell = tc;
				}
			}
		}

		var computedStyle = this._getComputedStyle(node);
		var background_color = this._getStyle(node, computedStyle, "background-color");
		var Shd;
		if (null != background_color && (background_color = this._ParseColor(background_color))) {
			Shd = new CDocumentShd();
			Shd.Value = c_oAscShdClear;
			Shd.Color = background_color;
			Shd.Fill = background_color;
		}

		for (var i = 0, length = node.childNodes.length; i < length; ++i) {
			//важно чтобы этот код был после определения td, потому что вертикально замерженые ячейки отсутствуют в dom
			fParseSpans();

			var tc = node.childNodes[i];
			var tcName = tc.nodeName.toLowerCase();
			if ("td" === tcName || "th" === tcName) {
				var nColSpan = tc.getAttribute("colspan");
				if (null != nColSpan)
					nColSpan = nColSpan - 0;
				else
					nColSpan = 1;
				if (tc === oBeforeCell)
					row.Set_Before(nColSpan);
				else if (tc === oAfterCell)
					row.Set_After(nColSpan);
				else {
					var oCurCell = row.Add_Cell(row.Get_CellsCount(), row, null, false);
					if (nColSpan > 1)
						oCurCell.Set_GridSpan(nColSpan);
					if (Shd) {
						oCurCell.Set_Shd(Shd);
					}
					var width = aSumGrid[nCellIndexSpan + nColSpan - 1] - aSumGrid[nCellIndexSpan - 1];
					oCurCell.Set_W(new CTableMeasurement(tblwidth_Mm, width));
					var nRowSpan = tc.getAttribute("rowspan");
					if (null != nRowSpan)
						nRowSpan = nRowSpan - 0;
					else
						nRowSpan = 1;
					if (nRowSpan > 1)
						oRowSpans[nCellIndexSpan] = {row: nRowSpan - 1, col: nColSpan, cell: oCurCell};
					this._ExecuteTableCell(tc, oCurCell, bUseScaleKoef, dScaleKoef, spacing, arrShapes, arrImages, arrTables);
				}
				nCellIndexSpan += nColSpan;
			}
		}
		fParseSpans();
	},

	_ExecuteTableCell: function (node, cell, bUseScaleKoef, dScaleKoef, spacing, arrShapes, arrImages, arrTables) {
		//Pr
		var Pr = cell.Pr;
		var bAddIfNull = false;
		if (null != spacing) {
			bAddIfNull = true;
		}

		var indent; 
		var className = node.className;
		if (className && this.oMsoStylesParser) {
			var msoClass = this.oMsoStylesParser.getMsoClassByName("." + className);
			if (msoClass) {
				indent = msoClass.getAttributeByName("mso-char-indent-count");
			}
		}

		var computedStyle = this._getComputedStyle(node);
		var background_color = this._getStyle(node, computedStyle, "background-color");
		if (null != background_color && (background_color = this._ParseColor(background_color))) {
			var Shd = new CDocumentShd();
			Shd.Value = c_oAscShdClear;
			Shd.Color = background_color;
			Shd.Fill = background_color;
			cell.Set_Shd(Shd);
		}
		var border = this._ExecuteBorder(computedStyle, node, "left", "Left", bAddIfNull);
		if (null != border)
			cell.Set_Border(border, 3);
		border = this._ExecuteBorder(computedStyle, node, "top", "Top", bAddIfNull);
		if (null != border)
			cell.Set_Border(border, 0);
		border = this._ExecuteBorder(computedStyle, node, "right", "Right", bAddIfNull);
		if (null != border)
			cell.Set_Border(border, 1);
		border = this._ExecuteBorder(computedStyle, node, "bottom", "Bottom", bAddIfNull);
		if (null != border)
			cell.Set_Border(border, 2);

		var top = this._getStyle(node, computedStyle, "padding-top");
		if (null != top && null != (top = AscCommon.valueToMm(top)))
			cell.Set_Margins({W: top, Type: tblwidth_Mm}, 0);
		var right = this._getStyle(node, computedStyle, "padding-right");
		if (null != right && null != (right = AscCommon.valueToMm(right)))
			cell.Set_Margins({W: right, Type: tblwidth_Mm}, 1);
		var bottom = this._getStyle(node, computedStyle, "padding-bottom");
		if (null != bottom && null != (bottom = AscCommon.valueToMm(bottom)))
			cell.Set_Margins({W: bottom, Type: tblwidth_Mm}, 2);
		var left = this._getStyle(node, computedStyle, "padding-left");
		if (null != left && null != (left = AscCommon.valueToMm(left)))
			cell.Set_Margins({W: left, Type: tblwidth_Mm}, 3);

		var whiteSpace = this._getStyle(node, computedStyle, "white-space");
		if ("nowrap" === whiteSpace || true === node.noWrap) {
			cell.SetNoWrap(true);
		}

		var vAlign = this._getStyle(node, computedStyle, "vertical-align");
		switch (vAlign) {
			case "middle":
				cell.Set_VAlign(vertalignjc_Center);
				break;
			case "bottom":
				cell.Set_VAlign(vertalignjc_Bottom);
				break;
			case "baseline":
			case "top":
				cell.Set_VAlign(vertalignjc_Top);
				break;
		}

		var i, length;
		var bPresentation = !PasteElementsId.g_bIsDocumentCopyPaste;
		if (bPresentation) {
			var arrShapes2 = [], arrImages2 = [], arrTables2 = [];
			var presentation = editor.WordControl.m_oLogicDocument;
			var shape = new CShape();
			shape.setParent(presentation.Slides[presentation.CurPage]);
			shape.setTxBody(AscFormat.CreateTextBodyFromString("", presentation.DrawingDocument, shape));
			arrShapes2.push(shape);
			this._Execute(node, {}, true, true, false, arrShapes2, arrImages2, arrTables);
			if (arrShapes2.length > 0) {
				var first_shape = arrShapes2[0];
				var content = first_shape.txBody.content;

				//добавляем новый параграфы
				for (i = 0, length = content.Content.length; i < length; ++i) {
					if (i === length - 1) {
						cell.Content.Internal_Content_Add(i + 1, content.Content[i], true);
					} else {
						cell.Content.Internal_Content_Add(i + 1, content.Content[i], false);
					}
				}
				//Удаляем параграф, который создается в таблице по умолчанию
				cell.Content.Internal_Content_Remove(0, 1);
				arrShapes2.splice(0, 1);
			}
			for (i = 0; i < arrShapes2.length; ++i) {
				arrShapes.push(arrShapes2[i]);
			}
			for (i = 0; i < arrImages2.length; ++i) {
				arrImages.push(arrImages2[i]);
			}
			for (i = 0; i < arrTables2.length; ++i) {
				arrTables.push(arrTables2[i]);
			}
		} else {
			//content
			var oPasteProcessor = new PasteProcessor(this.api, false, false, true);
			oPasteProcessor.AddedFootEndNotes = this.AddedFootEndNotes;
			oPasteProcessor.msoComments = this.msoComments;
			oPasteProcessor.oFonts = this.oFonts;
			oPasteProcessor.oImages = this.oImages;
			oPasteProcessor.bIsForFootEndnote = this.bIsForFootEndnote;
			oPasteProcessor.oDocument = cell.Content;
			oPasteProcessor.bIgnoreNoBlockText = true;
			oPasteProcessor.aMsoHeadStylesStr = this.aMsoHeadStylesStr;
			oPasteProcessor.oMsoHeadStylesListMap = this.oMsoHeadStylesListMap;
			oPasteProcessor.msoListMap = this.msoListMap;
			oPasteProcessor.dMaxWidth = this._CalcMaxWidthByCell(cell);
			oPasteProcessor.oMsoStylesParser = this.oMsoStylesParser;
			oPasteProcessor.pasteInExcel = this.pasteInExcel;
			if (true === bUseScaleKoef) {
				oPasteProcessor.bUseScaleKoef = bUseScaleKoef;
				oPasteProcessor.dScaleKoef = dScaleKoef;
			}
			oPasteProcessor._Execute(node, {}, true, true, false);
			oPasteProcessor._PrepareContent(indent);
			oPasteProcessor._AddNextPrevToContent(cell.Content);
			if (0 === oPasteProcessor.aContent.length) {
				var oDocContent = cell.Content;
				var oNewPar = new Paragraph(oDocContent.DrawingDocument, oDocContent);
				//выставляем единичные настройки - важно для копирования из таблиц и других мест где встречаются пустые ячейки
				var oNewSpacing = new CParaSpacing();
				oNewSpacing.Set_FromObject({After: 0, Before: 0, Line: Asc.linerule_Auto});
				oNewPar.Set_Spacing(oNewSpacing);
				oPasteProcessor.aContent.push(oNewPar);
			}
			//добавляем новый параграфы
			for (i = 0, length = oPasteProcessor.aContent.length; i < length; ++i) {
				if (i === length - 1) {
					cell.Content.Internal_Content_Add(i + 1, oPasteProcessor.aContent[i], true);
				} else {
					cell.Content.Internal_Content_Add(i + 1, oPasteProcessor.aContent[i], false);
				}
			}
			//Удаляем параграф, который создается в таблице по умолчанию
			cell.Content.Internal_Content_Remove(0, 1);
		}
	},

	_CheckIsPlainText: function (node, dNotCheckFirstElem) {
		var bIsPlainText = true;

		var checkStyle = function (elem) {
			var res = false;

			//TODO пересмотреть! возможно стоит сделать проверку на computedStyle.
			var _nodeName = elem.nodeName.toLowerCase();
			if ("h1" === _nodeName || "h2" === _nodeName || "h3" === _nodeName || "h4" === _nodeName || "h5" === _nodeName || "h6" === _nodeName) {
				return true;
			}

			var sClass = elem.getAttribute("class");
			var sStyle = elem.getAttribute("style");
			var sHref = elem.getAttribute("href");

			if (sClass || sStyle || sHref) {
				res = true;
			}
			return res;
		};

		//проверяем верхний элемент
		//в случае с плагинами - контент оборачивается в дивку, которая может иметь свои стили
		if ("body" !== node.nodeName.toLowerCase() && !dNotCheckFirstElem) {
			if (Node.ELEMENT_NODE === node.nodeType) {
				if (checkStyle(node)) {
					return false;
				}
			}
		}

		for (var i = 0, length = node.childNodes.length; i < length; i++) {
			var child = node.childNodes[i];
			if (Node.ELEMENT_NODE === child.nodeType) {
				if (checkStyle(child)) {
					bIsPlainText = false;
					break;
				} else if (!this._CheckIsPlainText(child, true)) {
					bIsPlainText = false;
					break;
				}

			}
		}
		return bIsPlainText;
	},

	_Execute: function (node, pPr, bRoot, bAddParagraph, bInBlock, arrShapes, arrImages, arrTables) {

		//bAddParagraph флаг влияющий на функцию _Decide_AddParagraph, добавлять параграф или нет.
		//bAddParagraph выставляется в true, когда встретился блочный элемент и по окончанию блочного элемента

		var oThis = this;
		var bRootHasBlock = false;//Если root есть блочный элемент, то надо все child считать параграфами
		//Для Root node не смотрим стили и не добавляем текст
		//var presentation = editor.WordControl.m_oLogicDocument;

		//для правки бага на релизе обработку добавляю только для вставки из ms excel, потом сделать данный класс как основной для получения данных из стилей ms
		if (AscCommon.g_clipboardBase.pastedFrom === AscCommon.c_oClipboardPastedFrom.Excel) {
			if (!this.oMsoStylesParser) {
				this.oMsoStylesParser = new MsoStylesParser(node);
			}
		}

		var parseTextNode = function () {
			var value = node.nodeValue;
			if (!value) {
				value = "";
			}

			var whiteSpacing = false;
			if (node.parentNode) {
				var computedStyle = oThis._getComputedStyle(node.parentNode);
				var tempWhiteSpacing = oThis._getStyle(node.parentNode, computedStyle, "white-space");
				whiteSpacing = "pre" === tempWhiteSpacing || "pre-wrap" === tempWhiteSpacing;

				//TODO заглушка! разобрать все ситуации(в тч и те, когда браузер добавляет при вставке лишние), когда пробельные символы нужно/не нужно сохранять
				if (!whiteSpacing && node.parentNode.nodeName && "span" === node.parentNode.nodeName.toLowerCase()) {
					if (value === " ") {
						whiteSpacing = true;
					} else if (value === "\n") {
						value = value.replace(/(\r|\t|\n)/g, ' ');
						whiteSpacing = true;
					}
				}
			}

			//в конструкциях вида text/n<b>text<b/> || <b>text<b/>/ntext заменяю символ переноса на пробел
			if ((node.nextSibling && node.nextSibling.nodeType !== Node.TEXT_NODE) ||
				(node.previousSibling && node.previousSibling.nodeType !== Node.TEXT_NODE)) {
				value = value.replace(/(\r|\t|\n)/g, ' ');
			}


			var checkPreviousNotSpaceText = function (_node) {
				if (!_node || oThis._IsBlockElem(_node.nodeName.toLowerCase())) {
					return false;
				}

				let _previousSibling = _node.previousSibling;
				while (_previousSibling) {
					if (oThis._IsBlockElem(_previousSibling.nodeName.toLowerCase())) {
						return false;
					}
					let siblingText = _previousSibling.textContent;
					if (siblingText && siblingText.length) {
						return siblingText.charCodeAt(siblingText.length - 1) !== 32;
					}
					_previousSibling = _previousSibling.previousSibling;
				}
				return _node.parentNode && checkPreviousNotSpaceText(_node.parentNode);
			};

			var _removeSpaces = function (_text, saveBefore) {
				let resText = "";
				for (let i = 0; i < _text.length; i++) {
					let curCode = _text.charCodeAt(i);
					let nextCode = _text.charCodeAt(i + 1);
					if (curCode === 32 && nextCode !== 32) {
						if (saveBefore || (!saveBefore && resText !== "")) {
							resText += _text[i];
						}
					} else if (curCode !== 32) {
						resText += _text[i];
					}
				}
				return resText;
			};

			//потому что(например иногда chrome при вставке разбивает строки с помощью \n)
			if (!whiteSpacing) {
				value = value.replace(/^(\r|\t|\n)+|(\r|\t|\n)+$/g, '');
				value = value.replace(/(\r|\t|\n)/g, ' ');

				//spaces before text - depends on text before. if text end on space -> spaces not save
				//else - convert into 1 space
				//space after text - convert into 1 space
				//all &nbsp; - must save

				//TODO must use in all cases. while in hotfix commit only special case
				//in develop: use this code for all cases
				var checkSpaces = value.replace(/(\s)/g, '');
				if (checkSpaces === "") {
					if (!checkPreviousNotSpaceText(node)) {
						//remove all spaces, exception non-breaking spaces + space after
						value = _removeSpaces(value);
					} else {
						//replace duplicated spaces before + after, exception non-breaking spaces
						value = _removeSpaces(value, true);
					}
				}
			}

			var Item;
			if (value.length > 0) {
				if (bPresentation) {
					oThis.oDocument = shape.txBody.content;
					if (bAddParagraph) {
                        let oParagraph = new Paragraph(oShapeContent.DrawingDocument, oShapeContent, oShapeContent.bPresentation === true);
                        oShapeContent.Internal_Content_Add(oShapeContent.Content.length, oParagraph);
                        oParagraph.CorrectContent();
                        oParagraph.CheckParaEnd();
					}

					if (!oThis.bIsPlainText) {
                        let oParagraph = oShapeContent.GetLastParagraph();
                        if(oParagraph) {
                            let oRun = new AscCommonWord.ParaRun(oShapeContent.GetLastParagraph(), false);
                            var rPr = oThis._read_rPr(node.parentNode);
                            oRun.SetPr(rPr);
                            oParagraph.AddToContentToEnd(oRun)
                        }
					}
				} else {
					var oTargetNode = node.parentNode;
					var bUseOnlyInherit = false;
					if (oThis._IsBlockElem(oTargetNode.nodeName.toLowerCase())) {
						bUseOnlyInherit = true;
					}
					bAddParagraph = oThis._Decide_AddParagraph(oTargetNode, pPr, bAddParagraph);

					//Добавляет элемени стиля если он поменялся
					oThis._commit_rPr(oTargetNode, bUseOnlyInherit);
				}

				//для проблемы с лишними прообелами в начале новой строки при копировании из MS EXCEL ячеек с текстом, разделенным alt+enter
				//мс в данном случае(баг 55851) оборачивает данные в тэг font. чтобы сохранить пробелы, добавляю проверку.
				//можно было бы проверить на наличие переноса строки вначале, но 55851 - перед вторым Bold добавляет символ переноса
				//тег font является устаревшим, но мс его активно испольует
				var ignoreFirstSpaces = false;
				if (AscCommon.g_clipboardBase.pastedFrom === AscCommon.c_oClipboardPastedFrom.Excel && !(node.parentNode && node.parentNode.nodeName.toLowerCase() === "font")) {
					ignoreFirstSpaces = true;
				}

				//bIsPreviousSpace - игнорируем несколько пробелов подряд
				var bIsPreviousSpace = false, clonePr;
				
				for (var oIterator = value.getUnicodeIterator(); oIterator.check(); oIterator.next()) {
					if (oThis.needAddCommentStart) {
						for (var i = 0; i < oThis.needAddCommentStart.length; i++) {
							oThis._CommitElemToParagraph(oThis.needAddCommentStart[i]);
							clonePr = oThis.oCurRun.Pr.Copy();
							oThis.oCurRun = new ParaRun(oThis.oCurPar);
							oThis.oCurRun.Set_Pr(clonePr);
						}
						oThis.needAddCommentStart = null;
					} else if (oThis.needAddCommentEnd) {
						oThis._commitCommentEnd();
					}


					var nUnicode = oIterator.value();
					if (ignoreFirstSpaces) {
						if (nUnicode === 32) {
							continue;
						} else {
							ignoreFirstSpaces = false;
						}
					}

					if (bPresentation) {

                        let oParagraph = oShapeContent.GetLastParagraph();
                        let oRun;
                        if(oParagraph) {
                            oRun = oParagraph.Content[oParagraph.Content.length - 2];
                        }
                        if(oRun) {
                            if (null !== nUnicode) {
                                if (0x20 !== nUnicode && 0xA0 !== nUnicode && 0x2009 !== nUnicode)
                                    Item = new AscWord.CRunText(nUnicode);
                                else
                                    Item = new AscWord.CRunSpace();

                                oRun.AddToContentToEnd(Item, false);
                            }
                        }
					} else if (!oThis.bIsForFootEndnote){
						if (null != nUnicode) {
							if (whiteSpacing && 0xa === nUnicode) {
								Item = null;
								bAddParagraph = oThis._Decide_AddParagraph(oTargetNode, pPr, true);
								oThis._commit_rPr(oTargetNode, bUseOnlyInherit);
							} else if (whiteSpacing && (0x9 === nUnicode || 0x2009 === nUnicode)) {
								Item = new AscWord.CRunTab();
							} else if (0x20 !== nUnicode && 0x2009 !== nUnicode) {
								Item = new AscWord.CRunText(nUnicode);
								bIsPreviousSpace = false;
							} else {
								Item = new AscWord.CRunSpace();
								if (bIsPreviousSpace) {
									continue;
								}
								if (!oThis.bIsPlainText && !whiteSpacing) {
									bIsPreviousSpace = true;
								}
							}
							if (null !== Item) {
								oThis._AddToParagraph(Item);
							}
						}
					}

				}
			}
		};
		/*var checkEndFootnodeText = function (Item, oThis, Text) {
			if (Item.parentNode) {
				if (Item.parentNode.nodeName.toLowerCase() === "a" && (Item.parentNode.name.toLowerCase().includes("ftn") || Item.parentNode.name.toLowerCase().includes("edn"))) {
					var oNewItem = Item;
					while (!(oNewItem.nodeName.toLowerCase() === "a")) {
						if (!oNewItem.parentNode || oNewItem.parentNode.nodeName.toLowerCase() === "body") {
							break;
						}
						oNewItem = oNewItem.parentNode;
					}
					if (oNewItem.nodeName.toLowerCase() === "a") {
						if (oNewItem.name.toLowerCase().includes("ftn") || oNewItem.name.toLowerCase().includes("edn")) {
							if (oNewItem.name.includes("_ftnref") || oNewItem.name.includes("_ednref")) {
								if (oThis.oCur_rPr.Color.b === 238 && oThis.oCur_rPr.Color.g === 0 && oThis.oCur_rPr.Color.r === 0 && oThis.oCur_rPr.Underline === true) {
									if (oNewItem.name.includes("_ftnref")) {
										oThis.AddedFootEndNotes[oNewItem.hash.replace("#_", "")].Ref.Run.SetRStyle(oThis.oLogicDocument.GetStyles().GetDefaultFootnoteReference());
									}
									else {
										oThis.AddedFootEndNotes[oNewItem.hash.replace("#_", "")].Ref.Run.SetRStyle(oThis.oLogicDocument.GetStyles().GetDefaultEndnoteReference());
									}
								}
								else {
									oThis.AddedFootEndNotes[oNewItem.hash.replace("#_", "")].Ref.Run.SetPr(oThis.oCur_rPr);
								}
							}
							else if (oNewItem.name.includes("_ftn") || oNewItem.name.includes("_edn")) {
								if (oThis.oCur_rPr.Color.b === 238 && oThis.oCur_rPr.Color.g === 0 && oThis.oCur_rPr.Color.r === 0 && oThis.oCur_rPr.Underline === true) {
									if (oNewItem.name.includes("_ftnref")) {
										oThis.AddedFootEndNotes[oNewItem.name.replace("_", "")].Content[0].Content[0].SetRStyle(oThis.oLogicDocument.GetStyles().GetDefaultFootnoteReference());
									}
									else {
										oThis.AddedFootEndNotes[oNewItem.name.replace("_", "")].Content[0].Content[0].SetRStyle(oThis.oLogicDocument.GetStyles().GetDefaultEndnoteReference());
									}
								}
								else {
									oThis.AddedFootEndNotes[oNewItem.name.replace("_", "")].Content[0].Content[0].SetPr(oThis.oCur_rPr);
								}
							}
						}
					}
					return true;
				}
				else if (Item.parentNode.nodeName.toLowerCase() === "p") {
					return false;
				}
				else {
					return checkEndFootnodeText(Item.parentNode, oThis, Text);
				}
			}
			return false;
		};*/

		var parseNumbering = function () {
			if ("ul" === sNodeName || "ol" === sNodeName || "li" === sNodeName) {
				if (bPresentation) {
					pPr.bNum = true;
					if (PasteElementsId.g_bIsDocumentCopyPaste) {
						if ("ul" === sNodeName) {
							pPr.numType = Asc.c_oAscNumberingFormat.Bullet;
						} else if ("ol" ===
							sNodeName) {
							pPr.numType = Asc.c_oAscNumberingFormat.Decimal;
						}
					} else {
						if ("ul" === sNodeName) {
							pPr.numType = AscFormat.numbering_presentationnumfrmt_Char;
						} else if ("ol" ===
							sNodeName) {
							pPr.numType = AscFormat.numbering_presentationnumfrmt_ArabicPeriod;
						}
					}
				} else {
					//в данном случае если нет тега li, то списоком не считаем
					if ("li" === sNodeName) {
						pPr.bNum = true;
					}

					if (PasteElementsId.g_bIsDocumentCopyPaste) {
						if ("ul" === sNodeName) {
							pPr.numType = Asc.c_oAscNumberingFormat.Bullet;
						} else if ("ol" ===
							sNodeName) {
							pPr.numType = Asc.c_oAscNumberingFormat.Decimal;
						}
					} else {
						if ("ul" === sNodeName) {
							pPr.numType = AscFormat.numbering_presentationnumfrmt_Char;
						} else if ("ol" ===
							sNodeName) {
							pPr.numType = AscFormat.numbering_presentationnumfrmt_ArabicPeriod;
						}
					}
				}
			} else if (pPr["mso-list"] && !bPresentation && "none" !== pPr['mso-list']) {
				if ("p" === sNodeName) {
					pPr.bNum = true;
				}
			}
		};

		var parseStyle = function () {
			//собираем не html свойства параграфа(те что нельзя получить из getComputedStyle)
			var style = node.getAttribute("style");
			if (style) {
				oThis._parseCss(style, pPr);
			}

			if ("h1" === sNodeName) {
				pPr.hLevel = 0;
			} else if ("h2" === sNodeName) {
				pPr.hLevel = 1;
			} else if ("h3" ===
				sNodeName) {
				pPr.hLevel = 2;
			} else if ("h4" === sNodeName) {
				pPr.hLevel = 3;
			} else if ("h5" ===
				sNodeName) {
				pPr.hLevel = 4;
			} else if ("h6" === sNodeName) {
				pPr.hLevel = 5;
			}
		};

		var parseImage = function () {
			//get width/height
			var sizeObj = oThis._getImgSize(node);
			var nWidth = sizeObj.width;
			var nHeight = sizeObj.height;
			var sSrc = node.getAttribute("src");

			if (bPresentation) {
				if (nWidth && nHeight && sSrc) {
					sSrc = oThis.oImages[sSrc];
					if (sSrc) {
						var image = AscFormat.DrawingObjectsController.prototype.createImage(sSrc, 0, 0, nWidth,
							nHeight);
						arrImages.push(image);
					}
				}
				return bAddParagraph;
			} else {
				if (PasteElementsId.g_bIsDocumentCopyPaste) {
					bAddParagraph = oThis._Decide_AddParagraph(node, pPr, bAddParagraph);

					if (nWidth && nHeight && sSrc) {
						sSrc = oThis.oImages[sSrc];
						if (sSrc) {
							//вписываем в oThis.dMaxWidth
							var bUseScaleKoef = oThis.bUseScaleKoef;
							var dScaleKoef = oThis.dScaleKoef;
							if (nWidth * dScaleKoef > oThis.dMaxWidth) {
								dScaleKoef = dScaleKoef * oThis.dMaxWidth / nWidth;
								bUseScaleKoef = true;
							}
							//закомментировал, потому что при вставке получаем изображения измененного размера
							/*if(bUseScaleKoef)
							 {
							 var dTemp = nWidth;
							 nWidth *= dScaleKoef;
							 nHeight *= dScaleKoef;
							 }*/
							var oTargetDocument = oThis.oDocument;
							var oDrawingDocument = oThis.oDocument.DrawingDocument;
							if (oTargetDocument && oDrawingDocument) {
								//если добавляем изображение в гиперссылку, то кладём его в отдельный ран и делаем не подчёркнутым
								if (oThis.oCurHyperlink) {
									oThis._CommitElemToParagraph(oThis.oCurRun);
									oThis.oCurRun = new ParaRun(oThis.oCurPar);
									oThis.oCurRun.Pr.Underline = false;
								}

								let fitPagePictureSize = oThis.fitPictureSizeToPage(nWidth, nHeight);
								nWidth = fitPagePictureSize.nWidth;
								nHeight = fitPagePictureSize.nHeight;

								var Drawing = CreateImageFromBinary(sSrc, nWidth, nHeight);
								if(!oThis.aNeedRecalcImgSize) {
									oThis.aNeedRecalcImgSize = [];
								}
								oThis.aNeedRecalcImgSize.push({drawing: Drawing, img: node});

								oThis._AddToParagraph(Drawing);

								if (oThis.oCurHyperlink) {
									oThis.oCurRun = new ParaRun(oThis.oCurPar);
								}

								//oDocument.AddInlineImage(nWidth, nHeight, img);
							}
						}
					}

					return bAddParagraph;
				} else {
					return false;
				}
			}
		};

		let checkOnlyBr = function (aElems) {
			if (!aElems) {
				return false;
			}
			let res;
			for (let i = 0; i < aElems.length; i++) {
				if (!aElems[i]) {
					continue;
				}

				var sNodeName = aElems[i].nodeName && aElems[i].nodeName.toLowerCase();
				if (sNodeName === "br") {
					if (res) {
						res = false;
						break;
					}
					res = true;
				} else {
					if (Node.TEXT_NODE === aElems[i].nodeType) {
						if (aElems[i].nodeValue && "" !== aElems[i].nodeValue.replace(/(\r|\t|\n)/g, '')) {
							res = false;
							break;
						}
					} else {
						res = false;
						break;
					}
				}
			}
			return res;
		};

		var parseLineBreak = function () {
			if (bPresentation) {
				//Добавляем linebreak, если он не разделяет блочные элементы и до этого был блочный элемент
				if ("br" === sNodeName || "always" === node.style.pageBreakBefore) {
                    let oParagraph = oShapeContent.GetLastParagraph();
                    let oRun;
                    if(oParagraph) {
                        oRun = oParagraph.Content[oParagraph.Content.length - 2];
                    }
                    if(oRun) {
                        oRun.AddToContentToEnd(new AscWord.CRunBreak(AscWord.break_Line), false);
                    }
				}
			} else {
				//Добавляем linebreak, если он не разделяет блочные элементы и до этого был блочный элемент
				var bPageBreakBefore = "always" === node.style.pageBreakBefore ||
					"left" === node.style.pageBreakBefore || "right" === node.style.pageBreakBefore;
				if ("br" == sNodeName || bPageBreakBefore) {

					//TODO пока комментирую добавление колонок
					/*if (AscCommon.g_clipboardBase.pastedFrom === AscCommon.c_oClipboardPastedFrom.Word && pPr.msoWordSection && "section-break" === pPr["mso-break-type"]) {
						//section break
						oThis._Add_NewParagraph();
						var oSectPr = new CSectionPr(oThis.oLogicDocument);
						var msoWordSection = oThis._findElemFromMsoHeadStyle("@page", pPr.msoWordSection);
						if (msoWordSection && msoWordSection[0]) {
							var sMsoColumns = msoWordSection[0]["mso-columns"];
							if (sMsoColumns) {
								var msoColumns = sMsoColumns.split(" ");
								oSectPr.Set_Columns_Num(msoColumns[0]);
								if (msoColumns[2]) {
									oSectPr.Set_Columns_Space(AscCommon.valueToMmType(msoColumns[2]).val);
								}
							}
							var sMargins = msoWordSection[0]["margin"];
							if (sMargins) {
								var margins = sMargins.split(" ");
								var _l = margins[0] && AscCommon.valueToMmType(margins[0]);
								var _t = margins[1] && AscCommon.valueToMmType(margins[1]);
								var _r = margins[2] && AscCommon.valueToMmType(margins[2]);
								var _b = margins[3] && AscCommon.valueToMmType(margins[3]);
								oSectPr.SetPageMargins(_l ? _l.val : undefined, _t ? _t.val : undefined, _r ? _r.val : undefined, _b ? _b.val : undefined);
							}
							var sPageSize = msoWordSection[0]["size"];
							if (sPageSize) {
								var pageSize = sPageSize.split(" ");
								var _w = pageSize[0] && AscCommon.valueToMmType(pageSize[0]);
								var _h = pageSize[1] && AscCommon.valueToMmType(pageSize[1]);
								//oSectPr.SetPageSize(_w ? _w.val : undefined, _h ? _h.val : undefined);
							}
						}
						oSectPr.Set_Type(c_oAscSectionBreakType.Continuous);
						oThis.oCurPar.Set_SectionPr(oSectPr, true);
						pPr.msoWordSection = null;
					} else */

					//by section
					/*if (oThis.oMsoSections && oThis.oMsoSections[pPr.msoWordSection]) {
						var oSectPr = oThis.getMsoSectionPr("@page", pPr.msoWordSection);

						var columnProps = new CDocumentColumnsProps();
						columnProps.From_SectPr(oSectPr);
						oThis.oMsoSections[pPr.msoWordSection].columnProps = columnProps;
						oThis.oMsoSections[pPr.msoWordSection].end = oThis.aContent.length;
					}
					oSectPr.Set_Type(c_oAscSectionBreakType.Continuous);
					oThis.oCurPar.Set_SectionPr(oSectPr, true);

					pPr.msoWordSection = null;*/

					if (bPageBreakBefore) {
						bAddParagraph = oThis._Decide_AddParagraph(node.parentNode, pPr, bAddParagraph);
						bAddParagraph = true;
						oThis._Commit_Br(0, node, pPr);
						oThis._AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));
					} else {
						bAddParagraph = oThis._Decide_AddParagraph(node.parentNode, pPr, bAddParagraph);

						//exception - ignore 1 br tag
						if (checkOnlyBr(node.parentNode.childNodes)) {
							return bAddParagraph;
						}

						oThis._Commit_Br(0, node, pPr);

						let breakLine = new AscWord.CRunBreak(AscWord.break_Line);
						if (breakLine.Flags && oThis.pasteInExcel) {
							//save mso flag for unite cell text with line breaks
							//use only in excel
							breakLine.Flags.sameCell = pPr["mso-data-placement"] === "same-cell";
						}
						oThis._AddToParagraph(breakLine);
					} /*else {
						bAddParagraph = oThis._Decide_AddParagraph(node.parentNode, pPr, bAddParagraph, false);
						oThis.nBrCount++;
						if ("line-break" === pPr["mso-special-character"] ||
							"always" === pPr["mso-column-break-before"]) {
							oThis._Commit_Br(0, node, pPr);
						}
					}*/
					return bAddParagraph;
				}
			}
			return null;
		};

		var parseTab = function () {
			var nTabCount;
			if (bPresentation) {
				nTabCount = parseInt(pPr["mso-tab-count"] || 0);
				if (nTabCount > 0) {
                    let oParagraph = oShapeContent.GetLastParagraph();
                    if(oParagraph) {
                        if (!oThis.bIsPlainText) {
                            var rPr = oThis._read_rPr(node);
                            let oRun = new AscCommonWord.ParaRun(oParagraph, false);
                            oRun.SetPr(rPr);
                            oParagraph.AddToContentToEnd(oRun)
                        }

                        let oRun;
                        oRun = oParagraph.Content[oParagraph.Content.length - 2];
                        if(oRun) {
                            for (var i = 0; i < nTabCount; i++) {
                                oRun.AddToContentToEnd(new AscWord.CRunTab(), false);
                            }
                        }
                    }
					return;
				}
			} else {
				nTabCount = parseInt(pPr["mso-tab-count"] || 0);
				if (nTabCount > 0) {
					bAddParagraph = oThis._Decide_AddParagraph(node, pPr, bAddParagraph);
					oThis._commit_rPr(node);
					for (var i = 0; i < nTabCount; i++) {
						oThis._AddToParagraph(new AscWord.CRunTab());
					}
					return bAddParagraph;
				}
			}

			return null;
		};

		let pushMathContent = function (_child) {
			if (oThis.isSupportPasteMathContent(child.nodeValue, true) && !oThis.pasteInExcel && oThis.apiEditor["asc_isSupportFeature"]("ooxml")) {
				let oPar = new Paragraph(oThis.oLogicDocument.DrawingDocument, bPresentation ? oShapeContent : null, bPresentation);

				History.TurnOff();
				let bAddNewParagraph = oThis._parseMathContent(_child, oPar);
				History.TurnOn();

				if (bAddNewParagraph !== false || !oThis.oCurPar || bPresentation) {
					let addedPar = oPar.Copy();
					if (bPresentation) {
						oShapeContent.Internal_Content_Add(oShapeContent.Content.length, addedPar);
						addedPar.CorrectContent();
						addedPar.CheckParaEnd();
					} else {
						oThis.aContent.push(addedPar);
					}
				} else {
					for (let i = 0; i < oPar.Content.length; i++) {
						if (oPar.Content[i] && oPar.Content[i].Get_Type && oPar.Content[i].Get_Type() === para_Math) {
							let oAddedParaMath = oPar.Content[i].Copy();
							oAddedParaMath.SetParagraph && oAddedParaMath.SetParagraph(oThis.oCurPar);
							oThis._CommitElemToParagraph(oAddedParaMath);
						}
					}
				}
			}
		};

		var parseChildNodes = function () {
			var sChildNodeName, bIsBlockChild, value, href, title;
			if (bPresentation) {
				sChildNodeName = child.nodeName.toLowerCase();

				if (!(Node.ELEMENT_NODE === nodeType || Node.TEXT_NODE === nodeType) || sChildNodeName === "style" ||
					sChildNodeName === "#comment" || sChildNodeName === "script" /*|| sChildNodeName === "o:p"*/) {
					if (sChildNodeName === "#comment") {
						pushMathContent(child);
					}
					return;
				}
				//попускам элеметы состоящие только из \t,\n,\r
				if (Node.TEXT_NODE === child.nodeType) {
					value = child.nodeValue;
					if (!value) {
						return;
					}
					value = value.replace(/(\r|\t|\n)/g, '');
					if ("" == value) {
						return;
					}
				}
				sChildNodeName = child.nodeName.toLowerCase();
				bIsBlockChild = oThis._IsBlockElem(sChildNodeName);
				if (bRoot) {
					oThis.bInBlock = false;
				}
				if (bIsBlockChild) {
					bAddParagraph = true;
					oThis.bInBlock = true;
				}

				var bHyperlink = false;
				var isPasteHyperlink = null;
				if ("a" === sChildNodeName) {
					href = child.href;
					if (null != href) {
						/*var sDecoded;
						//decodeURI может выдавать malformed exception, потому что наш сайт в utf8, а некоторые сайты могут кодировать url в своей кодировке(например windows-1251)
						try {
							sDecoded = decodeURI(href);
						} catch (e) {
							sDecoded = href;
						}
						href = sDecoded;*/
						bHyperlink = true;
						title = child.getAttribute("title");

						oThis.oDocument = shape.txBody.content;

						var Pos = (true === oThis.oDocument.Selection.Use ? oThis.oDocument.Selection.StartPos :
							oThis.oDocument.CurPos.ContentPos);
						isPasteHyperlink = node.getElementsByTagName('img');

						var text = null;
						if (isPasteHyperlink && isPasteHyperlink.length) {
							isPasteHyperlink = null;
						} else {
							text = child.innerText ? child.innerText : child.textContent;
						}

						if (href && href.length > Asc.c_nMaxHyperlinkLength) {
							isPasteHyperlink = false;
						}
						if (isPasteHyperlink) {
							var HyperProps = new Asc.CHyperlinkProperty({Text: text, Value: href, ToolTip: title});
							oThis.oDocument.Content[Pos].AddHyperlink(HyperProps);
						}
					}
				}

				//TODO временная правка. пересмотреть обработку тега math
				if (!child.style && Node.TEXT_NODE !== child.nodeType) {
					child.style = {};
				}

				if (!isPasteHyperlink) {
					bAddParagraph =
						oThis._Execute(child, Common_CopyObj(pPr), false, bAddParagraph, bIsBlockChild || bInBlock,
							arrShapes, arrImages, arrTables);
				}

				if (bIsBlockChild) {
					bAddParagraph = true;
				}
			} else {
				sChildNodeName = child.nodeName.toLowerCase();

				//исключаю чтение тега "o:p", потому что ms в пустые ячейки таблицы добавляется внутрь данного тега неразрывные пробелы
				//не нашёл такой ситуации, когда пропадают пробелы между словами
				//todo протестровать "o:p"!
				if (!(Node.ELEMENT_NODE === nodeType || Node.TEXT_NODE === nodeType) || sChildNodeName === "style" ||
					sChildNodeName === "#comment" || sChildNodeName === "script" /*|| sChildNodeName === "o:p"*/) {
					if (sChildNodeName === "#comment") {
						if (child.nodeValue === "[if !supportAnnotations]") {
							oThis.startMsoAnnotation = true;
						} else if (oThis.startMsoAnnotation && child.nodeValue === "[endif]") {
							oThis.startMsoAnnotation = false;
						} else {
							pushMathContent(child);
						}
					}
					return;
				}

				//добавляю пока флаг startMsoAnnotation для игнорирования комментариев при вставке из ms
				//TODO в дальнейшем необходимо поддержать вставку комментариев из ms
				//так же рассмотреть где ещё используется тег [if !supportAnnotations]
				if (oThis.startMsoAnnotation) {
					if (child.id) {
						//ориентируюсь по id для закрытия комментариев
						var idAnchor = child.id.split("_anchor_");
						if (idAnchor && idAnchor[1] && oThis.msoComments[idAnchor[1]] && oThis.msoComments[idAnchor[1]].start) {
							if (null != oThis.oCurRun) {
								if (!oThis.needAddCommentEnd) {
									oThis.needAddCommentEnd = [];
								}
								oThis.needAddCommentEnd.push(new AscCommon.ParaComment(false, oThis.msoComments[idAnchor[1]].start));
								delete oThis.msoComments[idAnchor[1]];
							}
						}
					}
					return;
				}

				//comments start
				var msoCommentReference = pPr["mso-comment-reference"];
				if (msoCommentReference) {
					var commentId = msoCommentReference.split("_");
					if (commentId && undefined !== commentId[1]) {
						var startComment = oThis.msoComments[commentId[1]];
						if (startComment && !startComment.start) {
							//добавляем комментарий AscCommon.CComment и получаем его id
							var newCCommentId = oThis._addComment({
								Date: pPr["mso-comment-date"],
								Text: startComment.text
							});
							//удаляем из map
							oThis.msoComments[commentId[1]].start = newCCommentId;
							//добавляем paraComment
							if (!oThis.needAddCommentStart) {
								oThis.needAddCommentStart = [];
							}
							oThis.needAddCommentStart.push(new AscCommon.ParaComment(true, newCCommentId));
						}
					}
				}

				//пропускаем одиночный неразрывный пробел перед комментарием
				if ("comment" === pPr["mso-special-character"]) {
					return;
				}

				//попускам элеметы состоящие только из \t,\n,\r
				if (Node.TEXT_NODE === child.nodeType) {
					value = child.nodeValue;
					if (!value) {
						return;
					}
					//TODO заглушка! разобрать все ситуации(в тч и те, когда браузер добавляет при вставке лишние), когда пробельные символы нужно/не нужно сохранять
					if (child.parentNode && child.parentNode.nodeName && "span" === child.parentNode.nodeName.toLowerCase()) {
						value = value.replace(/(\r|\t|\n)/g, ' ');
					} else {
						value = value.replace(/(\r|\t|\n)/g, '');
					}
					if ("" == value) {
						return;
					}
				}
				bIsBlockChild = oThis._IsBlockElem(sChildNodeName);
				if (bRoot) {
					oThis.bInBlock = false;
				}
				if (bIsBlockChild) {
					bAddParagraph = true;
					oThis.bInBlock = true;
				}
				var oOldHyperlink = null;
				var oOldHyperlinkContentPos = null;
				var oHyperlink = null;
			
				if ("a" === sChildNodeName) {
					href = child.href;
					if (null != href) {
						
						/*var sDecoded;
						//decodeURI может выдавать malformed exception, потому что наш сайт в utf8, а некоторые сайты могут кодировать url в своей кодировке(например windows-1251)
						try {
							sDecoded = decodeURI(href);
						} catch (e) {
							sDecoded = href;
						}
						href = sDecoded;*/
						if (href && href.length > 0) {
							// проверяем, является ли ссылка ссылкой на контент сноски
							// если да, то создаём сноску и добавляем в контент
							// если нет, значит создаём гиперссылку
							var sStr = href.split("#");
							var sText, oAddedRun;
							if (sStr[1] && -1 !== sStr[1].indexOf("_ftnref")) {
							}
							// обычная сноска
							else if (sStr[1] && -1 !== sStr[1].indexOf("_ftn")) {
								sText = child.innerText;
								// проверяем, является ли название сноски кастомной или нет
								if (sText[0] === "[" && sText[sText.length - 1] === "]") {
									sText = undefined;
								}
								bAddParagraph = oThis._Decide_AddParagraph(child, pPr, bAddParagraph);
								var oFootnote = oThis.oLogicDocument.Footnotes.CreateFootnote();
								oFootnote.AddDefaultFootnoteContent(sText);
								oAddedRun = new ParaRun(oThis.oCurPar, false);
								oAddedRun.SetRStyle(oThis.oLogicDocument.GetStyles().GetDefaultFootnoteReference());
								if (sText) {
									oAddedRun.AddToContent(0, new AscWord.CRunFootnoteReference(oFootnote));
									oAddedRun.AddText(sText, 1);
								}
								else {
									oAddedRun.AddToContent(0, new AscWord.CRunFootnoteReference(oFootnote));
								}
								oThis._CommitElemToParagraph(oAddedRun);
								if (oThis.AddedFootEndNotes) {
									oThis.AddedFootEndNotes[sStr[1].replace("_", "")] = oFootnote;
								}
							}
							else if (sStr[1] && -1 !== sStr[1].indexOf("_ednref")) {
							}
							// концевая сноска
							else if (sStr[1] && -1 !== sStr[1].indexOf("_edn")) {
								sText = child.innerText;
								// проверяем, является ли название сноски кастомной или нет
								if (sText[0] === "[" && sText[sText.length - 1] === "]") {
									sText = undefined;
								}
								bAddParagraph = oThis._Decide_AddParagraph(child, pPr, bAddParagraph);
								var oEndnote = oThis.oLogicDocument.Endnotes.CreateEndnote();
								oEndnote.AddDefaultEndnoteContent(sText);
								oAddedRun = new ParaRun(oThis.oCurPar, false);
								oAddedRun.SetRStyle(oThis.oLogicDocument.GetStyles().GetDefaultEndnoteReference());
								if (sText) {
									oAddedRun.AddToContent(0, new AscWord.CRunEndnoteReference(oEndnote));
									oAddedRun.AddText(sText, 1);
								}
								else {
									oAddedRun.AddToContent(0, new AscWord.CRunEndnoteReference(oEndnote));
								}
								oThis._CommitElemToParagraph(oAddedRun);
								if (oThis.AddedFootEndNotes) {
									oThis.AddedFootEndNotes[sStr[1].replace("_", "")] = oEndnote;
								}
							}
							else {
								title = child.getAttribute("title");

								bAddParagraph = oThis._Decide_AddParagraph(child, pPr, bAddParagraph);
								oHyperlink = new ParaHyperlink();
								oHyperlink.SetParagraph(oThis.oCurPar);
								oHyperlink.Set_Value(href);
								if (null != title) {
									oHyperlink.SetToolTip(title);
								}
								oOldHyperlink = oThis.oCurHyperlink;
								oOldHyperlinkContentPos = oThis.oCurHyperlinkContentPos;
								oThis.oCurHyperlink = oHyperlink;
								oThis.oCurHyperlinkContentPos = 0;
							}
						}
					}
				}

				//TODO временная правка. пересмотреть обработку тега math
				if (!child.style && Node.TEXT_NODE !== child.nodeType) {
					child.style = {};
				}

				bAddParagraph = oThis._Execute(child, Common_CopyObj(pPr), false, bAddParagraph, bIsBlockChild || bInBlock);

				//by section
				/*if (oThis.previousMsoWordSecion && oThis.previousMsoWordSecion === pPr.msoWordSection) {
					if (oThis.oMsoSections && oThis.oMsoSections[pPr.msoWordSection]) {
						oThis.oMsoSections[pPr.msoWordSection].end = oThis.aContent.length;
					}
					pPr.msoWordSection = null;
				}*/

				if (bIsBlockChild) {
					bAddParagraph = true;
				}
				if ("a" === sChildNodeName && null != oHyperlink) {
					oThis.oCurHyperlink = oOldHyperlink;
					oThis.oCurHyperlinkContentPos = oOldHyperlinkContentPos;
					if (oHyperlink.Content.length > 0) {
						if (oThis.pasteInPresentationShape) {
							var TextPr = new CTextPr();
							TextPr.Unifill = AscFormat.CreateUniFillSchemeColorWidthTint(11, 0);
							TextPr.Underline = true;
							oHyperlink.Apply_TextPr(TextPr, undefined, true);
						}

						//проставляем rStyle
						if (oHyperlink.Content && oHyperlink.Content.length && oHyperlink.Paragraph.bFromDocument) {
							if (oThis.oLogicDocument && oThis.oLogicDocument.Styles &&
								oThis.oLogicDocument.Styles.Default && oThis.oLogicDocument.Styles.Default.Hyperlink &&
								oThis.oLogicDocument.Styles.Style) {
								var hyperLinkStyle = oThis.oLogicDocument.Styles.Default.Hyperlink;

								for (var k = 0; k < oHyperlink.Content.length; k++) {
									if (oHyperlink.Content[k].Type === para_Run) {
										oHyperlink.Content[k].Set_RStyle(hyperLinkStyle);
									}
								}
							}
						}

						oThis._AddToParagraph(oHyperlink);
					}
				}
			}
		};
		var startExecuteNotes = function () {
			// Проверяем является ли тег контентом сноски
			if (node.nodeName.toLowerCase() === "div" && (-1 !== node.id.indexOf("ftn") || -1 !== node.id.indexOf("edn"))) {
				oThis.aContentForNotes = oThis.aContent;
				oThis.aContent = [];
			}
		};
		var endExecuteNotes = function () {
			// Проверяем является ли тег контентом сноски
			if (node.nodeName.toLowerCase() === "div" && (-1 !== node.id.indexOf("ftn") || -1 !== node.id.indexOf("edn"))) {
				var tmp = oThis.aContent;
				// Меняем контент обратно, для работы вне контента сносок
				oThis.aContent = oThis.aContentForNotes;
	
				// Заполняем контент сносок
				if (oThis.AddedFootEndNotes[node.id]) {
					for (var i = 0; i < tmp.length; i++) {
						if (i === 0) {
							oThis.AddedFootEndNotes[node.id].Content[0].Concat(tmp[i], false)
						}
						else {
							oThis.AddedFootEndNotes[node.id].Content.push(tmp[i]);
						}
					}
					// Удаляем сноску из PasteProcessor, после того, как добавили в неё контент
					delete oThis.AddedFootEndNotes[node.id];
				}
			}
		};
		var checkStylesForNotes = function() {
			if (node && node.name && node.nodeName && node.hash &&
				node.nodeName.toLowerCase() === "a" && (-1 !== node.name.toLowerCase().indexOf("ftn") || -1 !== node.name.toLowerCase().indexOf("edn"))) {
				if (-1 !== node.name.toLowerCase().indexOf("ftn") || -1 !== node.name.toLowerCase().indexOf("edn")) {

					if (oThis.AddedFootEndNotes) {
						
						if ((-1 !== node.name.indexOf("_ftnref") || -1 !== node.name.indexOf("_ednref"))
						&& oThis.AddedFootEndNotes[node.hash.replace("#_", "")] && oThis.AddedFootEndNotes[node.hash.replace("#_", "")].Ref && oThis.AddedFootEndNotes[node.hash.replace("#_", "")].Ref.Run) {

							if (oThis.oCur_rPr && oThis.oCur_rPr.Color && oThis.oCur_rPr.Color.b != null && oThis.oCur_rPr.Color.g != null && oThis.oCur_rPr.Color.r != null && oThis.oCur_rPr.Underline != null
								&& oThis.oCur_rPr.Color.b === 238 && oThis.oCur_rPr.Color.g === 0 && oThis.oCur_rPr.Color.r === 0 && oThis.oCur_rPr.Underline === true) {
								
								if (-1 !== node.name.indexOf("_ftnref")) {
									oThis.AddedFootEndNotes[node.hash.replace("#_", "")].Ref.Run.SetRStyle(oThis.oLogicDocument.GetStyles().GetDefaultFootnoteReference());
								}
								else {
									oThis.AddedFootEndNotes[node.hash.replace("#_", "")].Ref.Run.SetRStyle(oThis.oLogicDocument.GetStyles().GetDefaultEndnoteReference());
								}
							}
							else {
								oThis.AddedFootEndNotes[node.hash.replace("#_", "")].Ref.Run.SetPr(oThis.oCur_rPr);
							}
						}
						else if ((-1 !== node.name.indexOf("_ftn") || -1 !== node.name.indexOf("_edn"))
							&& oThis.AddedFootEndNotes[node.name.replace("_", "")] && oThis.AddedFootEndNotes[node.name.replace("_", "")].Content[0] && oThis.AddedFootEndNotes[node.name.replace("_", "")].Content[0].Content[0]) {
							
							if (oThis.oCur_rPr && oThis.oCur_rPr.Color && oThis.oCur_rPr.Color.b != null && oThis.oCur_rPr.Color.g != null && oThis.oCur_rPr.Color.r != null && oThis.oCur_rPr.Underline != null
								&& oThis.oCur_rPr.Color.b === 238 && oThis.oCur_rPr.Color.g === 0 && oThis.oCur_rPr.Color.r === 0 && oThis.oCur_rPr.Underline === true) {
								
								if (-1 !== node.name.indexOf("_ftnref")) {
									oThis.AddedFootEndNotes[node.name.replace("_", "")].Content[0].Content[0].SetRStyle(oThis.oLogicDocument.GetStyles().GetDefaultFootnoteReference());
								}
								else {
									oThis.AddedFootEndNotes[node.name.replace("_", "")].Content[0].Content[0].SetRStyle(oThis.oLogicDocument.GetStyles().GetDefaultEndnoteReference());
								}
							}
							else {
								oThis.AddedFootEndNotes[node.name.replace("_", "")].Content[0].Content[0].SetPr(oThis.oCur_rPr);
							}
						}
					}
				}
			}
		};

		if (node && node.nodeName && node.name &&
			node.nodeName.toLowerCase() === "a" && (-1 !== node.name.indexOf("ftn") || -1 !== node.name.indexOf("edn"))) {
				oThis.bIsForFootEndnote = true;
			}
		// Меняем контент, если начинается контент сноски
		startExecuteNotes();

		var bPresentation = !PasteElementsId.g_bIsDocumentCopyPaste;
        var oShapeContent;
        var shape;
		if (bPresentation) {
			shape = arrShapes[arrShapes.length - 1];
            oShapeContent = shape.txBody.content;
			this.aContent = oShapeContent.Content;
		}


		if (true === bRoot) {
			//Если блочных элементов нет, то отменяем флаг
			var bExist = false;
			for (var i = 0, length = node.childNodes.length; i < length; i++) {
				var child = node.childNodes[i];
				var bIsBlockChild = this._IsBlockElem(child.nodeName.toLowerCase());
				if (true === bIsBlockChild) {
					bRootHasBlock = true;
					bExist = true;
					break;
				}
			}
			if (false === bExist && true === this.bIgnoreNoBlockText) {
				this.bIgnoreNoBlockText = false;
			}
		} else {
			//TEXT NODE
			if (Node.TEXT_NODE === node.nodeType) {
				//TODO пересмотреть условия
				if (false === this.bIgnoreNoBlockText || true === bInBlock || (node.parentElement && "a" === node.parentElement.nodeName.toLowerCase())) {
					parseTextNode();
				}
				return bPresentation ? false : bAddParagraph;
			}

			//TABLE
			var sNodeName = node.nodeName.toLowerCase();
			if (bPresentation) {
				if ("table" === sNodeName) {
					this._StartExecuteTable(node, pPr, arrShapes, arrImages, arrTables);
					return;
				}
			} else {
				if ("table" === sNodeName && this.pasteInPresentationShape !== true) {
					if (PasteElementsId.g_bIsDocumentCopyPaste) {
						this._Commit_Br(1, node, pPr);

						this._StartExecuteTable(node, pPr);
						return bAddParagraph;
					} else {
						return false;
					}
				}
			}

			if ("w:sdt" === sNodeName && this.pasteInExcel !== true && this.pasteInPresentationShape !== true) {
				//CBlockLevelSdt
				bAddParagraph = this._Decide_AddParagraph(node, pPr, bAddParagraph);
				this._ExecuteBlockLevelStd(node, pPr);
				return bAddParagraph;
			}

			//STYLE
			parseStyle();

			//NUMBERING
			parseNumbering();

			//IMAGE
			if ("img" === sNodeName && (bPresentation || (!bPresentation && this.pasteInPresentationShape !== true))) {
				return parseImage();
			}

			//LINEBREAK
			var lineBreak = parseLineBreak();
			if (null !== lineBreak) {
				return lineBreak;
			}

			//TAB
			if ("span" === sNodeName) {
				var tab = parseTab();
				if (null !== tab) {
					return tab;
				}
			}
		}


		//рекурсивно вызываем для childNodes
		for (var i = 0, length = node.childNodes.length; i < length; i++) {
			var child = node.childNodes[i];
			var nodeType = child.nodeType;
			//При копировании из word может встретиться комментарий со списком
			//Комментарии пропускаем, списки делаем своими
			if (Node.COMMENT_NODE === nodeType) {
				var value = child.nodeValue;
				var bSkip = false;
				if (value) {
					if (-1 !== value.indexOf("supportLists")) {
						//todo распознать тип списка
						pPr.bNum = true;
						bSkip = true;
					}
					if (-1 !== value.indexOf("supportLineBreakNewLine")) {
						bSkip = true;
					}
					if (this.isSupportPasteMathContent(value) && !this.pasteInExcel && this.apiEditor["asc_isSupportFeature"]("ooxml")) {
						//TODO пока только в документы разрешаю вставку математики математику
						bSkip = true;
					}

					// TODO пересмотреть информацию в <![if !supportFootnotes]>
					/*if (-1 !== value.indexOf("supportFootnotes")) {
						bSkip = true;
					}*/
				}
				if (true === bSkip) {
					//пропускаем все до закрывающегося комментария
					var j = i + 1;
					for (; j < length; j++) {
						var tempNode = node.childNodes[j];
						var tempNodeType = tempNode.nodeType;
						if (Node.COMMENT_NODE === tempNodeType) {
							var tempvalue = tempNode.nodeValue;
							if (tempvalue && -1 !== tempvalue.indexOf("endif")) {
								break;
							}
						}
					}
					i = j;
					continue;
				}
			}

			parseChildNodes();

			//TODO пока не используется, поскольку есть проблемы при вставке колонок
			//by section
			/*if (AscCommon.g_clipboardBase.pastedFrom === AscCommon.c_oClipboardPastedFrom.Word) {
				if (child.className && -1 !== child.className.indexOf("WordSection")) {
					if (!oThis.oMsoSections) {
						oThis.oMsoSections = {};
					}
					pPr.msoWordSection = child.className;
					oThis.oMsoSections[pPr.msoWordSection] = {start: oThis.aContent.length};
				}
			}*/

			if (i === length - 1) {
				oThis._commitCommentEnd();
			}
		}

		if (bRoot && bPresentation) {
			this._Commit_Br(2, node, pPr);//word игнорируем 2 последних br
		}
		checkStylesForNotes();
		// Если тег с контентом для сноски, то записываем контент в сноску и заменяем его на обычный
		endExecuteNotes();

		if (node && node.nodeName && node.name &&
			node.nodeName.toLowerCase() === "a" && (-1 !== node.name.indexOf("ftn") || -1 !== node.name.indexOf("edn"))) {
				oThis.bIsForFootEndnote = false;
			}
		return bAddParagraph;
	},

	isSupportPasteMathContent: function (val, checkVersion) {
		let res = false;
		if (AscCommon.g_clipboardBase.pastedFrom === AscCommon.c_oClipboardPastedFrom.Word) {
			if (checkVersion && -1 !== val.indexOf("[if gte msEquation 12]")) {
				res = true;
			} else if (!checkVersion && -1 !== val.indexOf("[if !msEquation]")) {
				res = true;
			}
		}
		return res;
	},

	fitPictureSizeToPage: function (nWidth, nHeight) {
		if (this.apiEditor && this.apiEditor.isDocumentEditor) {
			if (this.oLogicDocument && this.oLogicDocument.GetColumnSize) {
				var oColumnSize = this.oLogicDocument.GetColumnSize();
				if (oColumnSize) {
					if (nWidth > oColumnSize.W || nHeight > oColumnSize.H) {
						if (oColumnSize.W > 0 && oColumnSize.H > 0) {
							var dScaleW = oColumnSize.W / nWidth;
							var dScaleH = oColumnSize.H / nHeight;
							var dScale = Math.min(dScaleW, dScaleH);
							nWidth *= dScale;
							nHeight *= dScale;
						}
					}
				}
			}
		}
		return {nWidth: nWidth, nHeight: nHeight};
	},

	_parseMathContent: function (node, oPar) {
		//получаем строку, которая содержит необходимые данные для получения уравнения
		let bAddNewParagraph = true;
		let str = node.nodeValue;
		str = str.replace('[if gte msEquation 12]>', '');
		str = str.replace('<![endif]', '');
		str = str.replace(/lang=\w*-\w*/g, '');
		if (0 === str.indexOf("<m:oMath>")) {
			bAddNewParagraph = false;
		}

		let xmlParserContext = typeof AscCommon.XmlParserContext !== "undefined" ? new AscCommon.XmlParserContext() : null;

		if (xmlParserContext) {
			xmlParserContext.oReadResult.bCopyPaste = true;
			var reader = typeof StaxParser !== "undefined" ? new StaxParser(str, /*documentPart*/null, xmlParserContext) : null;

			if (!reader || !reader.ReadNextNode()) {
				return;
			}

			let elem;
			if (bAddNewParagraph) {
				elem = typeof AscCommon.CT_OMathPara !== "undefined" ? new AscCommon.CT_OMathPara() : null;
			} else {
				elem = new ParaMath();
				oPar.AddToContentToEnd(elem);
			}

			if (elem && elem.fromXml) {
				elem.fromXml(reader, oPar);
				return bAddNewParagraph;
			}
		}

		return bAddNewParagraph;
	},

	_commitCommentEnd: function () {
		if (this.needAddCommentEnd && this.oCurRun && this.oCurPar) {
			for (var n = 0; n < this.needAddCommentEnd.length; n++) {
				this._CommitElemToParagraph(this.needAddCommentEnd[n]);
				var clonePr = this.oCurRun.Pr.Copy();
				this.oCurRun = new ParaRun(this.oCurPar);
				this.oCurRun.Set_Pr(clonePr);
			}
			this.needAddCommentEnd = null;
		}
	},

	_getImgSize: function (node) {
		if (!node) {
			return;
		}
		var oThis = this;
		var nWidth = parseInt(node.getAttribute("width"));
		var nHeight = parseInt(node.getAttribute("height"));
		if (!nWidth || !nHeight) {
			var computedStyle = oThis._getComputedStyle(node);
			nWidth = parseInt(oThis._getStyle(node, computedStyle, "width"));
			nHeight = parseInt(oThis._getStyle(node, computedStyle, "height"));
		}

		//TODO пересмотреть! node.getAttribute("width") в FF возврашает "auto" -> изображения в FF не всталяются
		if ((!nWidth || !nHeight)) {
			if (AscBrowser.isMozilla || AscBrowser.isIE) {
				nWidth = parseInt(node.width);
				nHeight = parseInt(node.height);
			} else if (AscBrowser.isChrome) {
				if (nWidth && !nHeight) {
					nHeight = nWidth;
				} else if (!nWidth && nHeight) {
					nWidth = nHeight;
				} else {
					nWidth = parseInt(node.width);
					nHeight = parseInt(node.height);
				}
			}
		}

		var sSrc = node.getAttribute("src");

		if (isNaN(nWidth)) {
			nWidth = 0;
		}
		if (isNaN(nHeight)) {
			nHeight = 0;
		}

		if (sSrc && (nWidth === 0 || nHeight === 0)) {
			var img_prop = new Asc.asc_CImgProperty();
			img_prop.asc_putImageUrl(sSrc);
			var or_sz = img_prop.asc_getOriginSize(window['Asc']['editor'] || window['editor']);
			nWidth = or_sz.Width;
			nHeight = or_sz.Height;
		} else {
			nWidth *= AscCommon.g_dKoef_pix_to_mm;
			nHeight *= AscCommon.g_dKoef_pix_to_mm;
		}

		if (!nWidth) {
			nWidth = oThis.defaultImgWidth;
		}
		if (!nHeight) {
			nHeight = oThis.defaultImgHeight;
		}

		return {width: nWidth, height: nHeight};
	},

	_applyStylesToTable: function (cTable, cStyle) {
		if (!cTable || !cStyle || (cTable && !cTable.Content))
			return;


		var row, tableCell;
		for (var i = 0; i < cTable.Content.length; i++) {
			row = cTable.Content[i];

			for (var j = 0; j < row.Content.length; j++) {
				tableCell = row.Content[j];
				//пока не заливаю функцию Internal_Compile_Pr(находится в table.js + правки)
				var test = this.Internal_Compile_Pr(cTable, cStyle, tableCell);
				tableCell.Set_Pr(test.CellPr);

				//проверка цвета
				/*cStyle.TableFirstRow.TableCellPr.Shd.Unifill.check(cTable.Get_Theme(), cTable.Get_ColorMap());
				var RGBA = cStyle.TableFirstRow.TableCellPr.Shd.Unifill.getRGBAColor();
				var theme = cTable.Get_Theme();*/
			}
		}
	},

	_addComment: function (oOldComment) {

		var convertMsDate = function (msDate) {
			var res = "";

			if (msDate) {
				var dateTimeSplit = msDate.split("T");
				if (dateTimeSplit[0] && dateTimeSplit[1]) {
					var year = dateTimeSplit[0].substring(0, 4);
					var month = dateTimeSplit[0].substring(4, 6);
					var day = dateTimeSplit[0].substring(6, 8);
					var hour = dateTimeSplit[1].substring(0, 2);
					var min = dateTimeSplit[1].substring(2, 4);
					var date = new Date(year, month - 1, day, hour, min);
					res = (date.getTime() - (new Date()).getTimezoneOffset() * 60000).toString();
				}
			}

			return res;
		};
		var fInitCommentData = function (comment) {
			var oCommentObj = new AscCommon.CCommentData();
			oCommentObj.m_nDurableId = AscCommon.CreateDurableId();
			if (null != comment.UserName) {
				oCommentObj.m_sUserName = comment.UserName;
			}
			if (null != comment.UserId) {
				oCommentObj.m_sUserId = comment.UserId;
			}
			if (null != comment.Date) {
				oCommentObj.m_sTime = convertMsDate(comment.Date);
			}
			if (null != comment.m_sQuoteText) {
				oCommentObj.m_sQuoteText = comment.m_sQuoteText;
			}
			if (null != comment.Text) {
				oCommentObj.m_sText = comment.Text;
			}
			if (null != comment.Solved) {
				oCommentObj.m_bSolved = comment.Solved;
			}
			if (null != comment.Replies) {
				for (var i = 0, length = comment.Replies.length; i < length; ++i) {
					oCommentObj.Add_Reply(fInitCommentData(comment.Replies[i]));
				}
			}
			return oCommentObj;
		};

		var oCommentsNewId = {};
		//меняем CDocumentContent на Document для возможности вставки комментариев в колонтитул и таблицу
		var isIntoShape = this.oDocument && this.oDocument.Parent && this.oDocument.Parent instanceof AscFormat.CShape ? true : false;
		var isIntoDocumentContent = this.oDocument instanceof CDocumentContent ? true : false;
		var document = this.oDocument && isIntoDocumentContent && !isIntoShape ? this.oDocument.LogicDocument : this.oDocument;

		var oNewComment = new AscCommon.CComment(document.Comments, fInitCommentData(oOldComment));
		document.Comments.Add(oNewComment);

		//посылаем событие о добавлении комментариев
		this.api.sync_AddComment && this.api.sync_AddComment(oNewComment.Id, oNewComment.Data);

		return oNewComment.Id;
	},

	//by section
	_applyMsoSections: function (aContent, sections) {
		if (sections) {
			for (var i in sections) {
				var columnsProps = sections[i].columnProps;
				var nStartPos = sections[i].start;
				var nEndPos = sections[i].end;
				if (nEndPos < nStartPos) {
					nStartPos = sections[i].end;
					nEndPos = sections[i].start;
				}

				/*var nStartPosSelection = this.oLogicDocument.Selection.StartPos;
				var nEndPosSelection = this.oLogicDocument.Selection.EndPos;
				if (nEndPosSelection < nStartPosSelection) {
					nStartPosSelection = this.oLogicDocument.Selection.EndPos;
					nEndPosSelection = this.oLogicDocument.Selection.StartPos;
				}
				var oStartSectPr = this.oLogicDocument.SectionsInfo.Get_SectPr(nStartPosSelection).SectPr;
				var oEndSectPr = this.oLogicDocument.SectionsInfo.Get_SectPr(nEndPosSelection).SectPr;
				if (!oStartSectPr || !oEndSectPr ||
					(oStartSectPr === oEndSectPr && oStartSectPr.IsEqualColumnProps(columnsProps))) {
					return;
				}
				var oEndParagraph = null;
				if (type_Paragraph !== aContent[nEndPos].GetType()) {
					oEndParagraph = new Paragraph(this.DrawingDocument, this.oLogicDocument);
					aContent.splice(nEndPos, 0, oEndParagraph);
				} else {
					oEndParagraph = aContent[nEndPos];
				}
				var oSectPr;
				if (nStartPos > 0 && (type_Paragraph !== aContent[nStartPos - 1].GetType() || !aContent[nStartPos - 1].Get_SectionPr())) {
					oSectPr = new CSectionPr(this.oLogicDocument);
					oSectPr.Copy(oStartSectPr, false);
					var oStartParagraph = new Paragraph(this.oLogicDocument.DrawingDocument, this.oLogicDocument);
					aContent.splice(nStartPos, 0, oStartParagraph);
					oStartParagraph.Set_SectionPr(oSectPr, true);
					nStartPos++;
					nEndPos++;
				}
				if (nEndPos !== aContent.length - 1) {
					oEndSectPr.Set_Type(c_oAscSectionBreakType.Continuous);
					oSectPr = new CSectionPr(this.oLogicDocument);
					oSectPr.Copy(oEndSectPr, false);
					oEndParagraph.Set_SectionPr(oSectPr, true);
					oSectPr.SetColumnProps(columnsProps);
				} else {
					oEndSectPr.Set_Type(c_oAscSectionBreakType.Continuous);
					oEndSectPr.SetColumnProps(columnsProps);
				}*/

				for (var nIndex = nStartPos; nIndex < nEndPos; ++nIndex) {
					var oElement = aContent[nIndex];
					if (type_Paragraph === oElement.GetType()) {
						var oCurSectPr = oElement.Get_SectionPr();
						if (oCurSectPr) {
							oCurSectPr.SetColumnProps(columnsProps);
						}
					}
				}
			}
		}
	},

	getMsoSectionPr: function (prefix, sectionName) {
		var oSectPr = new CSectionPr(this.oLogicDocument);
		var msoWordSection = this._findElemFromMsoHeadStyle(prefix, sectionName);
		if (msoWordSection && msoWordSection[0]) {
			var sMsoColumns = msoWordSection[0]["mso-columns"];
			//columns count + space
			if (sMsoColumns) {
				var msoColumns = sMsoColumns.split(" ");
				oSectPr.Set_Columns_Num(msoColumns[0]);
				if (msoColumns[2]) {
					oSectPr.Set_Columns_Space(AscCommon.valueToMmType(msoColumns[2]).val);
				}
			}
			//columns margin
			var sMargins = msoWordSection[0]["margin"];
			if (sMargins) {
				var margins = sMargins.split(" ");
				var _l = margins[0] && AscCommon.valueToMmType(margins[0]);
				var _t = margins[1] && AscCommon.valueToMmType(margins[1]);
				var _r = margins[2] && AscCommon.valueToMmType(margins[2]);
				var _b = margins[3] && AscCommon.valueToMmType(margins[3]);
				oSectPr.SetPageMargins(_l ? _l.val : undefined, _t ? _t.val : undefined, _r ? _r.val : undefined, _b ? _b.val : undefined);
			}
			/*var sPageSize = msoWordSection[0]["size"];
			if (sPageSize) {
				var pageSize = sPageSize.split(" ");
				var _w = pageSize[0] && AscCommon.valueToMmType(pageSize[0]);
				var _h = pageSize[1] && AscCommon.valueToMmType(pageSize[1]);
				oSectPr.SetPageSize(_w ? _w.val : undefined, _h ? _h.val : undefined);
			}*/
		}
		//oSectPr.Set_Type(c_oAscSectionBreakType.Continuous);

		return oSectPr;
	},

	getLastSectionPr: function () {
		var lastSectionIndex = 0;

		if (this.oMsoSections) {
			for (var i in this.oMsoSections) {
				var curSection = i.split("WordSection");
				if (curSection && curSection[1]) {
					var curIndex = parseInt(curSection[1]);
					if (curIndex && !isNaN(curIndex) && curIndex > lastSectionIndex) {
						lastSectionIndex = curIndex;
					}
				}
			}
		}

		var oSectPr = null;
		if (lastSectionIndex) {
			oSectPr = this.getMsoSectionPr("@page", "WordSection" + (lastSectionIndex + 1));
		}

		return oSectPr;
	}
};

function CheckDefaultFontFamily(val, api)
{
	return "onlyofficeDefaultFont" === val && api && api.getDefaultFontFamily ? api.getDefaultFontFamily() : val;
}

function CheckDefaultFontSize(val, api)
{
	return "0px" === val && api && api.getDefaultFontSize ? api.getDefaultFontSize() + "pt" : val;
}

function CreateImageFromBinary(bin, nW, nH)
{
    var w, h;

    if (nW === undefined || nH === undefined)
    {
        var _image = editor.ImageLoader.map_image_index[bin];
        if (_image != undefined && _image.Image != null && _image.Status == AscFonts.ImageLoadStatus.Complete)
        {
            var _w = Math.max(1, Page_Width - (X_Left_Margin + X_Right_Margin));
            var _h = Math.max(1, Page_Height - (Y_Top_Margin + Y_Bottom_Margin));

            var bIsCorrect = false;
            if (_image.Image != null)
            {
                var __w = Math.max(parseInt(_image.Image.width * g_dKoef_pix_to_mm), 1);
                var __h = Math.max(parseInt(_image.Image.height * g_dKoef_pix_to_mm), 1);

                var dKoef = Math.max(__w / _w, __h / _h);
                if (dKoef > 1)
                {
                    _w = Math.max(5, __w / dKoef);
                    _h = Math.max(5, __h / dKoef);

                    bIsCorrect = true;
                }
                else
                {
                    _w = __w;
                    _h = __h;
                }
            }

            w = __w;
            h = __h;
        }
        else
        {
            w = 50;
            h = 50;
        }
    }
    else
    {
        w = nW;
        h = nH;
    }
    var para_drawing = new ParaDrawing(w, h, null, editor.WordControl.m_oLogicDocument.DrawingDocument, editor.WordControl.m_oLogicDocument, null);
    var word_image = AscFormat.DrawingObjectsController.prototype.createImage(bin, 0, 0, w, h);
    para_drawing.Set_GraphicObject(word_image);
    word_image.setParent(para_drawing);
    para_drawing.Set_GraphicObject(word_image);
    return para_drawing;
}

function Check_LoadingDataBeforePrepaste(_api, _fonts, _images, _callback)
{
    let aPrepeareFonts = [];
    for (var font_family in _fonts)
    {
        aPrepeareFonts.push(new CFont(font_family));
    }
    AscFonts.FontPickerByCharacter.extendFonts(aPrepeareFonts);

    let aImagesToDownload = [];
	let aAllImages = [];
    for (let image in _images)
    {
        var src = _images[image];
        if (undefined !== window["Native"] && undefined !== window["Native"]["GetImageUrl"])
        {
            _images[image] = window["Native"]["GetImageUrl"](_images[image]);
        }
        else if (!g_oDocumentUrls.getImageUrl(src) && !g_oDocumentUrls.getImageLocal(src) && !g_oDocumentUrls.isThemeUrl(src))
        {
			aImagesToDownload.push(src);
        }
	    aAllImages.push({Url: _images[image]});
    }
    if (aImagesToDownload.length > 0)
    {
        AscCommon.sendImgUrls(_api, aImagesToDownload, function (data) {
            var image_map = {};
            for (var i = 0, length = Math.min(data.length, aImagesToDownload.length); i < length; ++i)
            {
                var elem = data[i];
                var sFrom = aImagesToDownload[i];
                if (null != elem.url)
                {
                    var name = g_oDocumentUrls.imagePath2Local(elem.path);
                    _images[sFrom] = name;
                    image_map[i] = name;
                }
                else
                {
                    image_map[i] = sFrom;
                }
            }
	        addThemeImagesToMap(image_map, aImagesToDownload, aAllImages);
            _api.pre_Paste(aPrepeareFonts, image_map, _callback);
        }, true);
    }
    else
        _api.pre_Paste(aPrepeareFonts, _images, _callback);
}

function addTextIntoRun(oCurRun, value, bIsAddTabBefore, dNotAddLastSpace, bIsAddTabAfter) {
	if (bIsAddTabBefore) {
		oCurRun.AddToContent(-1, new AscWord.CRunTab(), false);
	}

	for (var oIterator = value.getUnicodeIterator(); oIterator.check(); oIterator.next()) {
		var nUnicode = oIterator.value();

		var bIsSpace = true;
		var Item;
		if (0x2009 === nUnicode || 9 === nUnicode) {
			Item = new AscWord.CRunTab();
		} else if (0x20 !== nUnicode && 0xA0 !== nUnicode) {
			Item = new AscWord.CRunText(nUnicode);
			bIsSpace = false;
		} else {
			Item = new AscWord.CRunSpace();
		}

		//add text
		if (!(dNotAddLastSpace && oIterator.position() === value.length - 1 && bIsSpace)) {
			oCurRun.AddToContent(-1, Item, false);
		}
	}

	if (bIsAddTabAfter) {
		oCurRun.AddToContent(-1, new AscWord.CRunTab(), false);
	}
}

function searchBinaryClass(node)
{
	var res = null;
	if(node.children && node.children[0])
	{
		var child = node.children[0];
		var childClass = child ? child.getAttribute("class") : null;
		if(childClass != null && (childClass.indexOf("xslData;") > -1 || childClass.indexOf("docData;") > -1 || childClass.indexOf("pptData;") > -1))
		{
			return childClass;
		}
		else
		{
			return searchBinaryClass(node.children[0]);
		}
	}
	return res;
}

function SpecialPasteShowOptions()
{
	this.options = null;
	this.cellCoord = null;

	this.range = null;
	this.shapeId = null;
	this.fixPosition = null;
	this.position = null;

	//для электронных таблиц
	//показывать или нет дополнительный пункт специальной вставки
	this.showPasteSpecial = null;
	this.containTables = null;
}

SpecialPasteShowOptions.prototype = {
	constructor: SpecialPasteShowOptions,

	isClean: function() {
		var res = false;
		if(null === this.options && null === this.cellCoord && null === this.range && null === this.shapeId && null === this.fixPosition) {
			res = true;
		}
		return res;
	},

	clean: function() {
		this.options = null;
		this.cellCoord = null;

		this.range = null;
		this.shapeId = null;
		this.fixPosition = null;
		this.position = null;

		this.showPasteSpecial = null;
		this.containTables = null;
	},

	setRange: function(val) {
		this.range = val;
	},
	setShapeId: function(val) {
		this.shapeId = val;
	},
	setFixPosition: function(val) {
		this.fixPosition = val;
	},
	setPosition: function(val) {
		this.position = val;
	},

	asc_setCellCoord: function (val) {
		this.cellCoord = val;
	},
	asc_setOptions: function (val) {
		if(val === null) {
			this.options = [];
		} else {
			this.options = val;
		}
	},
	asc_getCellCoord: function () {
		return this.cellCoord;
	},
	asc_getOptions: function () {
		return this.options;
	},
	asc_getShowPasteSpecial: function () {
		return this.showPasteSpecial;
	},
	asc_setShowPasteSpecial: function (val) {
		this.showPasteSpecial = val;
	},
	asc_getContainTables: function () {
		return this.containTables;
	},
	asc_setContainTables: function (val) {
		this.containTables = val;
	}
};

function MsoStylesParser(node) {
	this.node = node;
	this.styleParsers = null;

	this._isInit = null;
}
MsoStylesParser.prototype.init = function () {
	//TODO копия функции _findMsoHeadStyle, пока не трогаю новый функционал, потом порефакторить и избавиться от старых функций
	var msoStyleParser;
	var styleTags = this.node && this.node.parentElement && this.node.parentElement.getElementsByTagName("style");
	if (styleTags && styleTags.length) {
		for (var i = 0; i < styleTags.length; ++i) {
			var headText = styleTags[i].innerText;
			if (!msoStyleParser) {
				msoStyleParser = new MsoStyleParser(headText);
			} else {
				msoStyleParser.clean();
				msoStyleParser.setStr(headText);
			}

			msoStyleParser.parse();
			if (!this.styleParsers) {
				this.styleParsers = [];
			}
			this.styleParsers.push(msoStyleParser);
		}
	}

	this._isInit = true;
};
MsoStylesParser.prototype.getMsoClassByName = function (name) {
	if (!this._isInit) {
		this.init();
	}
	if (this.styleParsers) {
		for (var i = 0; i < this.styleParsers.length; i++) {
			if (this.styleParsers[i]) {
				var msoStyleClass = this.styleParsers[i].getMsoClassByName(name);
				if (msoStyleClass) {
					return msoStyleClass;
				}
			}
		}
	}

	return null;
};

function MsoStyleParser(str) {
	this.str = str;

	this.startAttrs = null;
	this.startComment = null;
	this.attrName = null;
	this.attrVal = null;

	this.msoClass= null;

	this.classes = null;
}
MsoStyleParser.prototype.clean = function () {
	this.startAttrs = null;
	this.startComment = null;
	this.attrName = null;
	this.attrVal = null;

	this.msoClass= null;

	this.classes = null;
};
MsoStyleParser.prototype.setStr = function (str) {
	this.str = str;
};
MsoStyleParser.prototype.parse = function () {
	var _style = this.str;

	if (!_style) {
		return null;
	}

	for (var j = 0; j < _style.length; j++) {

		if (_style[j] === "\n" || _style[j] === "\t") {
			continue;
		}

		if (_style[j] === "<" && _style[j + 1] === "!" && _style[j + 2] === "-" && _style[j + 3] === "-") {
			j += 3;
			continue;
		}

		//skip comment
		if (_style[j] === "/" && _style[j + 1] === "*") {
			this.startComment = true;
			j++;
			continue;
		} else if (this.startComment) {
			if (_style[j] === "*" && _style[j + 1] === "/") {
				j++;
				this.startComment = false;
			}
			continue;
		}

		if (_style[j] === "{") {
			this.startReadAttrs();
		} else if (_style[j] === "}") {
			this.endReadClass();
		} else if (this.isStartReadAttrs()) {
			if (_style[j] === ":") {
				this.startReadAttrVal();
			} else if (_style[j] === ";") {
				this.endReadAttrVal();
			} else if (this.isStartReadAttrVal()) {
				this.addByAttrVal(_style[j])
			} else {
				this.addByAttrName(_style[j]);
			}
		} else {
			this.addByClassName(_style[j]);
		}
	}
};

MsoStyleParser.prototype.startReadAttrs = function () {
	this.startAttrs = true;
	this.startName = null;
};

MsoStyleParser.prototype.endReadClass = function () {
	if (!this.msoClass) {
		this.msoClass = new MsoStyleClass();
	}
	if (!this.msoClass.attributes) {
		this.msoClass.attributes = [];
	}
	if (this.attrName !== null) {
		this.msoClass.attributes[this.attrName] = this.attrVal;
	}
	if (!this.classes) {
		this.classes = [];
	}
	this.classes.push(this.msoClass);

	this.msoClass = null;
	this.startAttrs = false;
	this.attrName = null;
	this.attrVal = null;
};
MsoStyleParser.prototype.startReadAttrVal = function () {
	this.attrVal = "";
};
MsoStyleParser.prototype.isStartReadAttrs = function () {
	return this.startAttrs;
};
MsoStyleParser.prototype.endReadAttrVal = function () {
	if (!this.msoClass) {
		this.msoClass = new MsoStyleClass();
	}

	if (!this.msoClass.attributes) {
		this.msoClass.attributes = [];
	}
	this.msoClass.attributes[this.attrName] = this.attrVal;

	this.attrName = null;
	this.attrVal = null;
};
MsoStyleParser.prototype.isStartReadAttrVal = function () {
	return this.attrVal !== null;
};
MsoStyleParser.prototype.addByAttrVal = function (sym) {
	this.attrVal += sym;
};
MsoStyleParser.prototype.addByAttrName = function (sym) {
	if (!this.attrName) {
		this.attrName = "";
	}
	this.attrName += sym;
};
MsoStyleParser.prototype.addByClassName = function (sym) {
	//TODO - классов может быть несколько с одинаковыми аттрибутами
	//TODO имя класса разделить по пробелам - @page WordSection1(префикс + имя)
	/*p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
	{mso-style-priority:34;
		mso-style-unhide:no;
		mso-style-qformat:yes;
		margin-top:0in;
		margin-right:0in;
		margin-bottom:8.0pt;
		margin-left:.5in;
		mso-add-space:auto;*/

	if (!this.msoClass) {
		this.msoClass = new MsoStyleClass();
	}
	if (!this.msoClass.name) {
		this.msoClass.name = "";
	}
	if (this.msoClass.name === "" && sym === " ") {
		return;
	}
	this.msoClass.name += sym;
};

MsoStyleParser.prototype.getMsoClassByName = function (name) {
	if (this.classes) {
		for (var i = 0; i < this.classes.length; i++) {
			if (this.classes[i].name === name) {
				return this.classes[i];
			}
		}
	}

	return null;
};



function MsoStyleClass() {
	this.name = null;
	this.attributes = null;
}

MsoStyleClass.prototype.getAttributeByName = function (name) {
	if (this.attributes) {
		return this.attributes[name];
	}

	return null;
};

function ParseHtmlMathContent(paragraph, opt_textPr) {
	this.paraRuns = null;
	this.paragraph = paragraph;

	this.cTextPr = opt_textPr ? opt_textPr : new CTextPr();
}

ParseHtmlMathContent.prototype.fromXml = function (reader) {
	var state = reader.getState();
	CMathBase.prototype.fromHtmlCtrlPr.call(this, reader, this.cTextPr);
	this.getXmlRunsRecursive(reader);
	reader.setState(state);
};

ParseHtmlMathContent.prototype.getXmlRunsRecursive = function (reader, textPr, doNotCopyPr) {
	let depth = reader.GetDepth();
	if (!textPr) {
		textPr = this.cTextPr;
	}
	var copyTextPr;
	while (reader.ReadNextSiblingNode(depth)) {
		let name = reader.GetNameNoNS();
		switch (name) {
			case "u": {
				textPr.Underline = true;
				this.getXmlRunsRecursive(reader, textPr);
				break;
			}
			case "s": {
				textPr.Strikeout = true;
				this.getXmlRunsRecursive(reader, textPr);
				break;
			}
			case "r": {
				var paraRun = new ParaRun(this.paragraph, true);
				var state = reader.getState();
				paraRun.fromXml(reader);
				reader.setState(state);

				copyTextPr = textPr.Copy();
				this.getXmlRunsRecursive(reader, copyTextPr, true);
				reader.setState(state);

				var txt = reader.GetTextDecodeXml();
				for (var oIterator = txt.getUnicodeIterator(); oIterator.check(); oIterator.next()) {
					var nUnicode = oIterator.value();
					var cMath = new CMathText();
					cMath.add(nUnicode);
					paraRun.Add_ToContent(paraRun.GetElementsCount(), cMath, false);
				}

				if (copyTextPr != null) {
					if (paraRun.Pr) {
						paraRun.Pr.Merge(copyTextPr);
					} else {
						paraRun.Set_Pr(copyTextPr);
					}
				}

				if (!this.paraRuns) {
					this.paraRuns = [];
				}
				this.paraRuns.push(paraRun);
				break;
			}
			case "span": {
				copyTextPr = doNotCopyPr ? textPr : textPr.Copy();
				CMathBase.prototype.fromHtmlCtrlPr.call(this, reader, copyTextPr);
				this.getXmlRunsRecursive(reader, copyTextPr);
				break;
			}
			default:
				this.getXmlRunsRecursive(reader, textPr);
				break;
		}
	}
};

function ParseHtmlStyle(styles) {
	this.styles = styles;
	this.map = null;
}

ParseHtmlStyle.prototype.parseStyles = function () {
	if (!this.map) {
		this.map = new Map();
	}

	var map = this.map;
	var buff = this.styles.split(';');
	var tagname;
	var val;
	for (var i = 0; i < buff.length; i++) {
		tagname = buff[i].substring(0, buff[i].indexOf(':')).replace(/\s/g, '');
		val = buff[i].substring(buff[i].indexOf(':') + 1).replace(/\s/g, '');

		map.set(tagname, val);
	}

	return map;
};

ParseHtmlStyle.prototype.applyStyles = function (textPr) {
	if (!textPr || !this.map) {
		return;
	}

	var map = this.map;

	//color
	var color = map.get('color');
	if (color) {
		var cDocColor = new CDocumentColor(0, 0, 0);
		cDocColor.SetFromHexColor(color);
		textPr.Color = cDocColor;
	}
	/*var color = this._getStyle(node, computedStyle, "color");
	if (color && (color = this._ParseColor(color))) {
		if (PasteElementsId.g_bIsDocumentCopyPaste) {
			rPr.Color = color;
		} else {
			if (color) {
				rPr.Unifill = AscFormat.CreateUnfilFromRGB(color.r, color.g, color.b);
			}
		}
	}*/


	var font_size = map.get('font-size');
	//font_size = CheckDefaultFontSize(font_size, this.apiEditor);
	if (font_size) {
		var obj = AscCommon.valueToMmType(font_size);
		if (obj && "%" !== obj.type && "none" !== obj.type) {
			font_size = obj.val;
			//Если браузер не поддерживает нецелые пикселы отсекаем половинные шрифты, они появляются при вставке 8, 11, 14, 20, 26pt
			if ("px" === obj.type && false === this.bIsDoublePx)
				font_size = Math.round(font_size * g_dKoef_mm_to_pt);
			else
				font_size = Math.round(2 * font_size * g_dKoef_mm_to_pt) / 2;//половинные значения допустимы.

			//TODO use constant
			if (font_size > 300)
				font_size = 300;
			else if (font_size === 0)
				font_size = 1;

			textPr.FontSize = font_size;
			textPr.FontSizeCS = font_size;
		}
	}


	var font_weight = map.get("font-weight");
	if (font_weight) {
		if ("bold" === font_weight || "bolder" === font_weight || 400 < font_weight) {
			textPr.Bold = true;
		}
	}

	var font_style = map.get("font-style");
	if ("italic" === font_style) {
		textPr.Italic = true;
	}

	var spacing = map.get("letter-spacing");
	if (spacing && null != (spacing = AscCommon.valueToMm(spacing))) {
		textPr.Spacing = spacing;
	}


	var text_decoration = map.get("text-decoration");
	if (text_decoration) {
		if (-1 !== text_decoration.indexOf("underline")) {
			textPr.Underline = true;
		}
		if (-1 !== text_decoration.indexOf("line-through")) {
			textPr.Strikeout = true;
		}
	}

	var background_color = map.get("background");
	if (background_color) {
		var highLight = AscCommon.PasteProcessor.prototype._ParseColor(background_color);
		if (highLight != null) {
			textPr.HighLight = highLight;
		}
	}


	var vertical_align = map.get("vertical-align");
	switch (vertical_align) {
		case "sub":
			textPr.VertAlign = AscCommon.vertalign_SubScript;
			break;
		case "super":
			textPr.VertAlign = AscCommon.vertalign_SuperScript;
			break;
	}
};

function checkOnlyOneImage(node)
{
	var res = false;

	if (node && node.childNodes) {
		for (var i = 0; i < node.childNodes.length; i++) {
			var sChildNodeName = node.childNodes[i].nodeName.toLowerCase();
			if (sChildNodeName === "style" || sChildNodeName === "#comment" || sChildNodeName === "script") {
				continue;
			}
			if (sChildNodeName === "img") {
				if (res) {
					res = false;
					break;
				} else {
					res = true;
				}
			} else {
				res = false;
				break;
			}
		}
	}

	return res;
}

function addThemeImagesToMap(oImageMap, aDwnldUrls, aImages) {
	if(!Array.isArray(aDwnldUrls)) {
		return;
	}
	if(!Array.isArray(aImages)) {
		return;
	}
	if(aDwnldUrls.length === aImages.length) {
		return;
	}
	let oAddMap = {};
	let nAddIdx = 0;
	for(let nImg = 0; nImg < aImages.length; ++nImg) {
		let oImgObject = aImages[nImg];
		if(!AscCommon.isRealObject(oImgObject)) {
			continue;
		}
		let sUrl = oImgObject.Url;
		let nDnldImg;
		for(nDnldImg = 0; nDnldImg < aDwnldUrls.length; ++nDnldImg) {
			let sDnldUrl = aDwnldUrls[nDnldImg];
			if(sUrl === sDnldUrl) {
				break;
			}
		}
		if(nDnldImg === aDwnldUrls.length) {
			oAddMap[nAddIdx++] = sUrl;
		}
	}
	for(let sKey in oImageMap) {
		if(oImageMap.hasOwnProperty(sKey)) {
			oAddMap[nAddIdx++] = oImageMap[sKey];
			delete oImageMap[sKey];
		}
	}
	for(let sKey in oAddMap) {
		if(oAddMap.hasOwnProperty(sKey)) {
			oImageMap[sKey] = oAddMap[sKey];
		}
	}
}

  //---------------------------------------------------------export---------------------------------------------------
  window['AscCommon'] = window['AscCommon'] || {};
  window["AscCommon"].Check_LoadingDataBeforePrepaste = Check_LoadingDataBeforePrepaste;
  window["AscCommon"].CDocumentReaderMode = CDocumentReaderMode;
  window["AscCommon"].GetObjectsForImageDownload = GetObjectsForImageDownload;
  window["AscCommon"].ResetNewUrls = ResetNewUrls;
  window["AscCommon"].CopyProcessor = CopyProcessor;
  window["AscCommon"].CopyPasteCorrectString = CopyPasteCorrectString;
  window["AscCommon"].Editor_Paste_Exec = Editor_Paste_Exec;
  window["AscCommon"].sendImgUrls = sendImgUrls;
  window["AscCommon"].PasteProcessor = PasteProcessor;
  window["AscCommon"].GetContentFromHtml = GetContentFromHtml;


  window["AscCommon"].addTextIntoRun = addTextIntoRun;
  window["AscCommon"].searchBinaryClass = searchBinaryClass;

  window["AscCommon"].PasteElementsId = PasteElementsId;
  window["AscCommon"].CheckDefaultFontFamily = CheckDefaultFontFamily;


  window["Asc"]["SpecialPasteShowOptions"] = window["Asc"].SpecialPasteShowOptions = SpecialPasteShowOptions;
  prot									 = SpecialPasteShowOptions.prototype;
  prot["asc_getCellCoord"]					= prot.asc_getCellCoord;
  prot["asc_getOptions"]					= prot.asc_getOptions;
  prot["asc_getShowPasteSpecial"]			= prot.asc_getShowPasteSpecial;
  prot["asc_getContainTables"]			    = prot.asc_getContainTables;

  window["AscCommon"].checkOnlyOneImage = checkOnlyOneImage;

  window["AscCommon"].koef_mm_to_indent = koef_mm_to_indent;

  window["AscCommon"].ParseHtmlMathContent = ParseHtmlMathContent;
  window["AscCommon"].ParseHtmlStyle = ParseHtmlStyle;


})(window);
