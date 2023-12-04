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

(function(window)
{

	/**
	 * Класс для поиска внутри параграфа
	 * @constructor
	 */
	function CParagraphSearch(Paragraph, SearchEngine, Type)
	{
		this.Paragraph    = Paragraph;
		this.SearchEngine = SearchEngine;
		this.Type         = Type;

		this.ContentPos   = new AscWord.CParagraphContentPos();

		this.StartPos     = null; // Запоминаем здесь стартовую позицию поиска
		this.SearchIndex  = 0;    // Номер символа, с которым мы проверяем совпадение
		this.StartPosBuf  = [];
	}

	CParagraphSearch.prototype.Reset = function()
	{
		this.StartPos    = null;
		this.SearchIndex = 0;
		this.StartPosBuf = [];
	};
	CParagraphSearch.prototype.Check = function(nIndex, oItem)
	{
		return this.SearchEngine.Pattern.Check(nIndex, oItem, this.SearchEngine);
	};
	CParagraphSearch.prototype.GetPrefix = function(nIndex)
	{
		return this.SearchEngine.GetPrefix(nIndex);
	};
	CParagraphSearch.prototype.CheckSearchEnd = function()
	{
		return (this.SearchIndex === this.SearchEngine.Pattern.GetLength());
	};

	/**
	 * Найденные элементы в параграфе. Записаны в массиве Paragraph.SearchResults
	 * @constructor
	 */
	function CParagraphSearchElement(StartPos, EndPos, Type, Id)
	{
		this.StartPos  = StartPos;
		this.EndPos    = EndPos;
		this.Type      = Type;
		this.ResultStr = "";
		this.Id        = Id;

		this.ClassesS = [];
		this.ClassesE = [];
	}
	CParagraphSearchElement.prototype.RegisterClass = function(isStart, oClass)
	{
		if (isStart)
			this.ClassesS.push(oClass);
		else
			this.ClassesE.push(oClass);
	};


	/**
	 * Метки начала и конца найденных элементов внутри параграфа
	 * @constructor
	 */
	function CParagraphSearchMark(SearchResult, Start, Depth)
	{
		this.SearchResult = SearchResult;
		this.Start        = Start;
		this.Depth        = Depth;
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CParagraphSearch        = CParagraphSearch;
	window['AscCommonWord'].CParagraphSearchElement = CParagraphSearchElement;
	window['AscCommonWord'].CParagraphSearchMark    = CParagraphSearchMark;

})(window);
