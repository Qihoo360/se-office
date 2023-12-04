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

(
	/**
	 * @param {Window} window
	 * @param {undefined} undefined
	 */
	function ( window, undefined) {
		/** @constructor */
		function asc_CDownloadOptions(fileType, isDownloadEvent) {
			this.fileType = fileType;
			this.isDownloadEvent = !!isDownloadEvent;
			this.advancedOptions = null;
			this.compatible = true;

			this.isNaturalDownload = false;
			this.errorDirect = null;
			this.oDocumentMailMerge = null;
			this.oMailMergeSendData = null;
			this.callback = null;
			
			this.isGetTextFromUrl = null;
			this.isPdfPrint = false;
			this.pdfChanges = null;

			this.textParams = null;
		}

		asc_CDownloadOptions.prototype.asc_setFileType = function (fileType) {this.fileType = fileType;};
		asc_CDownloadOptions.prototype.asc_setIsDownloadEvent = function (isDownloadEvent) {this.isDownloadEvent = isDownloadEvent;};
		asc_CDownloadOptions.prototype.asc_setAdvancedOptions = function (advancedOptions) {this.advancedOptions = advancedOptions;};
		asc_CDownloadOptions.prototype.asc_setCompatible = function (compatible) {this.compatible = compatible;};
		asc_CDownloadOptions.prototype.asc_setTextParams = function (textParams) {this.textParams = textParams;};

		/** @constructor */
		function asc_CAdvancedOptions(opt) {
			this.codePages = function () {
				var arr = [], c, encodings = opt["encodings"];
				for (var i = 0; i < encodings.length; i++) {
					c = new asc_CCodePage();
					c.init(encodings[i]);
					arr.push(c);
				}
				return arr;
			}();
			this.recommendedSettings = new asc_CTextOptions(opt["codepage"], opt["delimiter"]);
			this.data = opt["data"];
		}
		asc_CAdvancedOptions.prototype.asc_getCodePages = function () {return this.codePages;};
		asc_CAdvancedOptions.prototype.asc_getRecommendedSettings = function () {return this.recommendedSettings;};
		asc_CAdvancedOptions.prototype.asc_getData = function () {return this.data;};

		/** @constructor */
		function asc_CTextOptions(codepage, delimiter, delimiterChar) {
			this.codePage = codepage;
			this.delimiter = delimiter;
			this.delimiterChar = delimiterChar;

			this.textQualifier = null;

			this.numberDecimalSeparator = null;
			this.numberGroupSeparator = null;

			this.delimiterRows = null;
			this.matchMode = null;
		}
		asc_CTextOptions.prototype.asc_getDelimiter = function(){return this.delimiter;};
		asc_CTextOptions.prototype.asc_setDelimiter = function(v){this.delimiter = v;};
		asc_CTextOptions.prototype.asc_getDelimiterChar = function(){return this.delimiterChar;};
		asc_CTextOptions.prototype.asc_setDelimiterChar = function(v){this.delimiterChar = v;};
		asc_CTextOptions.prototype.asc_getCodePage = function(){return this.codePage;};
		asc_CTextOptions.prototype.asc_setCodePage = function(v){this.codePage = v;};
		asc_CTextOptions.prototype.asc_getNumberDecimalSeparator = function () {
			return this.numberDecimalSeparator !== null ? this.numberDecimalSeparator : AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator;
		};
		asc_CTextOptions.prototype.asc_getNumberGroupSeparator = function () {
			return this.numberGroupSeparator !== null ? this.numberGroupSeparator : AscCommon.g_oDefaultCultureInfo.NumberGroupSeparator;
		};
		asc_CTextOptions.prototype.asc_setNumberDecimalSeparator = function(v){this.numberDecimalSeparator = v;};
		asc_CTextOptions.prototype.asc_setNumberGroupSeparator = function(v){this.numberGroupSeparator = v;};
		asc_CTextOptions.prototype.asc_getNumberDecimalSeparatorsArr = function () {
			var arr = [",", "."];
			if (arr.indexOf(AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator)) {
				arr.push(AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator);
			}

			return arr;
		};
		asc_CTextOptions.prototype.asc_getNumberGroupSeparatorsArr = function () {
			var arr = [",", ".", "'"];
			if (arr.indexOf(AscCommon.g_oDefaultCultureInfo.NumberGroupSeparator)) {
				arr.push(AscCommon.g_oDefaultCultureInfo.NumberGroupSeparator);
			}

			return arr;
		};
		asc_CTextOptions.prototype.asc_setTextQualifier = function(v){this.textQualifier = v;};
		asc_CTextOptions.prototype.asc_getTextQualifier = function(){return this.textQualifier;};
		asc_CTextOptions.prototype.asc_getTextQualifierArr = function () {
			return ["\"", "'", null];
		};
		asc_CTextOptions.prototype.getDelimiterRows = function () {
			return this.delimiterRows;
		};


		/** @constructor */
		function asc_CDRMAdvancedOptions(password){
			this.password = password;
		}
		asc_CDRMAdvancedOptions.prototype.asc_getPassword = function(){return this.password;};
		asc_CDRMAdvancedOptions.prototype.asc_setPassword = function(v){this.password = v;};

		/** @constructor */
		function asc_CCodePage(){
			this.codePageName = null;
			this.codePage = null;
			this.text = null;
			this.lcid = null;
		}
		asc_CCodePage.prototype.init = function (encoding) {
			this.codePageName = encoding["name"];
			this.codePage = encoding["codepage"];
			this.text = encoding["text"];
			this.lcid = encoding["lcid"];
		};
		asc_CCodePage.prototype.asc_getCodePageName = function(){return this.codePageName;};
		asc_CCodePage.prototype.asc_setCodePageName = function(v){this.codePageName = v;};
		asc_CCodePage.prototype.asc_getCodePage = function(){return this.codePage;};
		asc_CCodePage.prototype.asc_setCodePage = function(v){this.codePage = v;};
		asc_CCodePage.prototype.asc_getText = function(){return this.text;};
		asc_CCodePage.prototype.asc_setText = function(v){this.text = v;};
		asc_CCodePage.prototype.asc_getLcid = function(){return this.lcid;};
		asc_CCodePage.prototype.asc_setLcid = function(v){this.lcid = v;};

		/** @constructor */
		function asc_CDelimiter(delimiter){
			this.delimiterName = delimiter;
		}
		asc_CDelimiter.prototype.asc_getDelimiterName = function(){return this.delimiterName;};
		asc_CDelimiter.prototype.asc_setDelimiterName = function(v){ this.delimiterName = v;};

		/** @constructor */
		function asc_CFormulaGroup(name){
			this.groupName = name;
			this.formulasArray = [];
		}
		asc_CFormulaGroup.prototype.asc_getGroupName = function() { return this.groupName; };
		asc_CFormulaGroup.prototype.asc_getFormulasArray = function() { return this.formulasArray; };
		asc_CFormulaGroup.prototype.asc_addFormulaElement = function(o) { return this.formulasArray.push(o); };

		/** @constructor */
		function asc_CFormula(o){
			this.name = o.name;
		}
		asc_CFormula.prototype.asc_getName = function () {
			return this.name;
		};
		asc_CFormula.prototype.asc_getLocaleName = function () {
			return AscCommonExcel.cFormulaFunctionToLocale ? AscCommonExcel.cFormulaFunctionToLocale[this.name] : this.name;
		};

		/** @constructor */
		function asc_CTextParams(association){
			this.association = association;
		}
		asc_CTextParams.prototype.asc_getAssociation = function () {
			return this.association;
		};

		//----------------------------------------------------------export----------------------------------------------------
		var prot;
		window['Asc'] = window['Asc'] || {};
		window['AscCommon'] = window['AscCommon'] || {};

		window["Asc"].asc_CDownloadOptions = window["Asc"]["asc_CDownloadOptions"] = asc_CDownloadOptions;
		prot = asc_CDownloadOptions.prototype;
		prot["asc_setFileType"] = prot.asc_setFileType;
		prot["asc_setIsDownloadEvent"] = prot.asc_setIsDownloadEvent;
		prot["asc_setAdvancedOptions"] = prot.asc_setAdvancedOptions;
		prot["asc_setCompatible"] = prot.asc_setCompatible;
		prot["asc_setTextParams"] = prot.asc_setTextParams;

		window["AscCommon"].asc_CAdvancedOptions = asc_CAdvancedOptions;
		prot = asc_CAdvancedOptions.prototype;
		prot["asc_getCodePages"] = prot.asc_getCodePages;
		prot["asc_getRecommendedSettings"] = prot.asc_getRecommendedSettings;
		prot["asc_getData"]	= prot.asc_getData;

		window["Asc"].asc_CTextOptions = window["Asc"]["asc_CTextOptions"] = asc_CTextOptions;
		prot = asc_CTextOptions.prototype;
		prot["asc_getDelimiter"] = prot.asc_getDelimiter;
		prot["asc_setDelimiter"] = prot.asc_setDelimiter;
		prot["asc_getDelimiterChar"] = prot.asc_getDelimiterChar;
		prot["asc_setDelimiterChar"] = prot.asc_setDelimiterChar;
		prot["asc_getCodePage"] = prot.asc_getCodePage;
		prot["asc_setCodePage"] = prot.asc_setCodePage;
		prot["asc_setNumberDecimalSeparator"] = prot.asc_setNumberDecimalSeparator;
		prot["asc_setNumberGroupSeparator"] = prot.asc_setNumberGroupSeparator;
		prot["asc_setTextQualifier"] = prot.asc_setTextQualifier;
		prot["asc_getTextQualifier"] = prot.asc_getTextQualifier;
		
		window["Asc"].asc_CDRMAdvancedOptions = window["Asc"]["asc_CDRMAdvancedOptions"] = asc_CDRMAdvancedOptions;
		prot = asc_CDRMAdvancedOptions.prototype;
		prot["asc_getPassword"] = prot.asc_getPassword;
		prot["asc_setPassword"] = prot.asc_setPassword;

		prot = asc_CCodePage.prototype;
		prot["asc_getCodePageName"]		= prot.asc_getCodePageName;
		prot["asc_setCodePageName"]		= prot.asc_setCodePageName;
		prot["asc_getCodePage"]			= prot.asc_getCodePage;
		prot["asc_setCodePage"]			= prot.asc_setCodePage;
		prot["asc_getText"]				= prot.asc_getText;
		prot["asc_setText"]				= prot.asc_setText;
		prot["asc_getLcid"]				= prot.asc_getLcid;
		prot["asc_setLcid"]				= prot.asc_setLcid;

		prot = asc_CDelimiter.prototype;
		prot["asc_getDelimiterName"]			= prot.asc_getDelimiterName;
		prot["asc_setDelimiterName"]			= prot.asc_setDelimiterName;

		window["AscCommon"].asc_CFormulaGroup = asc_CFormulaGroup;
		prot = asc_CFormulaGroup.prototype;
		prot["asc_getGroupName"]				= prot.asc_getGroupName;
		prot["asc_getFormulasArray"]			= prot.asc_getFormulasArray;
		prot["asc_addFormulaElement"]			= prot.asc_addFormulaElement;

		window["AscCommon"].asc_CFormula = asc_CFormula;
		prot = asc_CFormula.prototype;
		prot["asc_getName"]				= prot.asc_getName;
		prot["asc_getLocaleName"]	= prot.asc_getLocaleName;

		window["AscCommon"].asc_CTextParams = window["AscCommon"]["asc_CTextParams"] = asc_CTextParams;
		prot = asc_CTextParams.prototype;
		prot["asc_getAssociation"]				= prot.asc_getAssociation;

	}
)(window);
