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

(function (window)
{
	const DEFAULT_EXCEPTIONS = {};

	DEFAULT_EXCEPTIONS[lcid_enUS] = [
		"a", "abbr", "abs", "acct", "addn", "adj", "advt", "al", "alt", "amt", "anon", "approx", "appt", "apr", "apt", "assn", "assoc", "asst", "attn", "attrib", "aug", "aux", "ave", "avg",
		"b", "bal", "bldg", "blvd", "bot", "bro", "bros",
		"c", "ca", "calc", "cc", "cert", "certif", "cf", "cit", "cm", "co", "comp", "conf", "confed", "const", "cont", "contrib", "coop", "corp", "ct",
		"d", "dbl", "dec", "decl", "def", "defn", "dept", "deriv", "diag", "diff", "div", "dm", "dr", "dup", "dupl",
		"e", "encl", "eq", "eqn", "equip", "equiv", "esp", "esq", "est", "etc", "excl", "ext",
		"f", "feb", "ff", "fig", "freq", "fri", "ft", "fwd",
		"g", "gal", "gen", "gov", "govt",
		"h", "hdqrs", "hgt", "hist", "hosp", "hq", "hr", "hrs", "ht", "hwy",
		"i", "ib", "ibid", "illus", "in", "inc", "incl", "incr", "int", "intl", "irreg", "ital",
		"j", "jan", "jct", "jr", "jul", "jun",
		"k", "kg", "km", "kmh",
		"l", "lang", "lb", "lbs", "lg", "lit", "ln", "lt",
		"m", "mar", "masc", "max", "mfg", "mg", "mgmt", "mgr", "mgt", "mhz", "mi", "min", "misc", "mkt", "mktg", "ml", "mm", "mngr", "mon", "mph", "mr", "mrs", "msec", "msg", "mt", "mtg", "mtn", "mun",
		"n", "na", "name", "nat", "natl", "ne", "neg", "ng", "no", "norm", "nos", "nov", "num", "nw",
		"o", "obj", "occas", "oct", "op", "opt", "ord", "org", "orig", "oz",
		"p", "pa", "pg", "pkg", "pl", "pls", "pos", "pp", "ppt", "pred", "pref", "prepd", "prev", "priv", "prof", "proj", "pseud", "psi", "pt", "publ",
		"q", "qlty", "qt", "qty",
		"r", "rd", "re", "rec", "ref", "reg", "rel", "rep", "req", "reqd", "resp", "rev",
		"s", "sat", "sci", "se", "sec", "sect", "sep", "sept", "seq", "sig", "soln", "soph", "spec", "specif", "sq", "sr", "st", "sta", "stat", "std", "subj", "subst", "sun", "supvr", "sw",
		"t", "tbs", "tbsp", "tech", "tel", "temp", "thur", "thurs", "tkt", "tot", "transf", "transl", "tsp", "tues",
		"u", "univ", "util",
		"v", "var", "veg", "vert", "viz", "vol", "vs",
		"w", "wed", "wk", "wkly", "wt",
		"x",
		"y", "yd", "yr",
		"z"
	];

	DEFAULT_EXCEPTIONS[lcid_ruRU] = [
		"а",
		"б",
		"вв",
		"гг", "гл",
		"д", "др",
		"е", "ед",
		"ё",
		"ж",
		"з",
		"и",
		"й",
		"к", "кв", "кл", "коп", "куб",
		"лл",
		"м", "мл", "млн", "млрд",
		"н", "наб", "нач",
		"о", "обл", "обр", "ок",
		"п", "пер", "пл", "пос", "пр",
		"руб",
		"сб", "св", "см", "соч", "ср", "ст", "стр",
		"тт", "тыс",
		"у",
		"ф",
		"х",
		"ц",
		"ш", "шт",
		"щ",
		"ъ",
		"ы",
		"ь",
		"э", "экз",
		"ю"
	];


	/**
	 * Класс для работы с исключениями автозамены первого символа в предложении
	 * @constructor
	 */
	function CFirstLetterExceptions()
	{
		this.Exceptions = {};
		this.MaxLen     = 0;
	}
	CFirstLetterExceptions.GetDefaultExceptions = function(lang)
	{
		return DEFAULT_EXCEPTIONS[lang] ? DEFAULT_EXCEPTIONS[lang] : [];
	};
	CFirstLetterExceptions.prototype.Check = function(word, lang)
	{
		if (!word)
			return false;

		let exceptions = this.GetExceptionsByLang(lang);
		let _word      = word.toLowerCase();

		let firstCodePoint = _word.codePointAt(0);
		if (!exceptions[firstCodePoint])
			return false;

		return (-1 !== exceptions[firstCodePoint].indexOf(_word));
	};
	CFirstLetterExceptions.prototype.GetExceptions = function(lang)
	{
		let exceptions = this.GetExceptionsByLang(lang);
		let result = [];
		for (let codePoint in exceptions)
		{
			result = result.concat(exceptions[codePoint]);
		}
		return result;
	};
	CFirstLetterExceptions.prototype.SetExceptions = function(exceptions, lang)
	{
		this.Exceptions[lang] = this.ToExceptionArray(exceptions);
	};
	CFirstLetterExceptions.prototype.AddException = function(word, lang)
	{
		let exceptions = this.GetExceptionsByLang(lang);
		let _word      = word.toLowerCase();

		let firstCodePoint = _word.codePointAt(0);
		if (!exceptions[firstCodePoint])
			return false;

		if (-1 === exceptions[firstCodePoint].indexOf(_word))
			exceptions[firstCodePoint].push(_word);
	};
	CFirstLetterExceptions.prototype.RemoveException = function(word, lang)
	{
		let exceptions = this.GetExceptionsByLang(lang);
		let _word      = word.toLowerCase();

		let firstCodePoint = _word.codePointAt(0);
		if (!exceptions[firstCodePoint])
			return false;

		let index = exceptions[firstCodePoint].indexOf(_word);
		if (-1 !== index)
			exceptions[firstCodePoint].splice(index, 1);
	};
	CFirstLetterExceptions.prototype.GetMaxLen = function()
	{
		return this.MaxLen;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CFirstLetterExceptions.prototype.GetExceptionsByLang = function(lang)
	{
		if (!this.Exceptions[lang])
			this.Exceptions[lang] = this.ToExceptionArray(DEFAULT_EXCEPTIONS[lang]);

		return this.Exceptions[lang];
	};
	CFirstLetterExceptions.prototype.ToExceptionArray = function(words)
	{
		if (!words)
			return {};

		let result = {};
		for (let index = 0, count = words.length; index < count; ++index)
		{
			let word = words[index];
			if (!word)
				continue;

			if (word.length > this.MaxLen)
				this.MaxLen = word.length;

			let codePoint = word.codePointAt(0);
			if (!result[codePoint])
				result[codePoint] = [];

			result[codePoint].push(word);
		}

		return result;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'].CFirstLetterExceptions = CFirstLetterExceptions;

	CFirstLetterExceptions.prototype["get_Exceptions"]        = CFirstLetterExceptions.prototype.get_Exceptions = CFirstLetterExceptions.prototype.GetExceptions;
	CFirstLetterExceptions.prototype["put_Exceptions"]        = CFirstLetterExceptions.prototype.put_Exceptions = CFirstLetterExceptions.prototype.SetExceptions;
	CFirstLetterExceptions.prototype["get_DefaultExceptions"] = CFirstLetterExceptions.prototype.get_DefaultExceptions = CFirstLetterExceptions.GetDefaultExceptions;
	CFirstLetterExceptions.prototype["add_Exception"]         = CFirstLetterExceptions.prototype.add_Exception = CFirstLetterExceptions.prototype.AddException;
	CFirstLetterExceptions.prototype["remove_Exception"]      = CFirstLetterExceptions.prototype.remove_Exception = CFirstLetterExceptions.prototype.RemoveException;

})(window);
