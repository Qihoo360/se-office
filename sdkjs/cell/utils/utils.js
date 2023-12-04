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
	function (window, undefined) {
		// Import
		var gc_nMaxRow0 = AscCommon.gc_nMaxRow0;
		var gc_nMaxCol0 = AscCommon.gc_nMaxCol0;
		var g_oCellAddressUtils = AscCommon.g_oCellAddressUtils;
		var AscBrowser = AscCommon.AscBrowser;

		var c_oAscSelectionType = Asc.c_oAscSelectionType;


		/** @const */
		var kLeftLim1 = .999999999999999;
		var MAX_EXCEL_INT = 1e308;
		var MIN_EXCEL_INT = -MAX_EXCEL_INT;
		var c_sPerDay = 86400;
		var c_msPerDay = c_sPerDay * 1000;

		/** @const */
		var kUndefinedL = "undefined";
		/** @const */
		var kNullL = "null";
		/** @const */
		var kObjectL = "object";
		/** @const */
		var kFunctionL = "function";
		/** @const */
		var kNumberL = "number";
		/** @const */
		var kArrayL = "array";

		var recalcType = {
			recalc	: 0, // без пересчета
			full		: 1, // пересчитываем все
			newLines: 2  // пересчитываем новые строки

		};

		var sizePxinPt = 72 / 96;

		function applyFunction(callback) {
			if (kFunctionL === typeof callback) {
				callback.apply(null, Array.prototype.slice.call(arguments, 1));
			}
		}

		function typeOf(obj) {
			if (obj === undefined) {return kUndefinedL;}
			if (obj === null) {return kNullL;}
			return Object.prototype.toString.call(obj).slice(8, -1).toLowerCase();
		}

		function lastIndexOf(s, regExp, fromIndex) {
			var end = fromIndex >= 0 && fromIndex <= s.length ? fromIndex : s.length;
			for (var i = end - 1; i >= 0; --i) {
				var j = s.slice(i, end).search(regExp);
				if (j >= 0) {return i + j;}
			}
			return -1;
		}

		function search(arr, fn) {
			for (var i = 0; i < arr.length; ++i) {
				if ( fn(arr[i]) ) {return i;}
			}
			return -1;
		}

		function getUniqueRangeColor (arrRanges, curElem, tmpColors) {
			var colorIndex, j, range = arrRanges[curElem];
			for (j = 0; j < curElem; ++j) {
				if (range.isEqual(arrRanges[j])) {
					colorIndex = tmpColors[j];
					break;
				}
			}
			return colorIndex;
		}

		function getMinValueOrNull (val1, val2) {
			return null === val2 ? val1 : (null === val1 ? val2 : Math.min(val1, val2));
		}

		function round(x) {
			var y = x + (x >= 0 ? .5 : -.5);
			return y | y;
			//return Math.round(x);
		}

		function floor(x) {
			var y = x | x;
			y -= x < 0 && y > x ? 1 : 0;
			return y + (x - y > kLeftLim1 ? 1 : 0); // to fix float number precision caused by binary presentation
			//return Math.floor(x);
		}

		function ceil(x) {
			var y = x | x;
			y += x > 0 && y < x ? 1 : 0;
			return y - (y - x > kLeftLim1 ? 1 : 0); // to fix float number precision caused by binary presentation
			//return Math.ceil(x);
		}

		function incDecFonSize (bIncrease, oValue) {
			// Закон изменения размеров :
			// Результатом должно быть ближайшее из отрезка [8,72] по следующим числам 8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72
			// Если значение меньше или равно 8 и мы уменьшаем, то ничего не меняется
			// Если значение больше или равно 72 и мы увеличиваем, то ничего не меняется

			var aSizes = [8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72];
			var nLength = aSizes.length;
			var i;
			if (true === bIncrease) {
				if (oValue >= aSizes[nLength - 1])
					return null;
				for (i = 0; i < nLength; ++i)
					if (aSizes[i] > oValue)
						break;
			} else {
				if (oValue <= aSizes[0])
					return null;
				for (i = nLength - 1; i >= 0; --i)
					if (aSizes[i] < oValue)
						break;
			}

			return aSizes[i];
		}

		function calcDecades(num) {
			return Math.abs(num) < 10 ? 1 : 1 + calcDecades(floor(num * 0.1));
		}

		function convertPtToPx(value) {
			value = value / sizePxinPt;
			value = (value * AscBrowser.retinaPixelRatio) >> 0;			
			return value;
		}
		function convertPxToPt(value) {
			value = value * sizePxinPt;
			//пункты округляем до сотых
			value = Asc.ceil(value / AscBrowser.retinaPixelRatio * 100) / 100;
			return value;
		}

		// Определяет времени работы функции
		function profileTime(fn/*[, arguments]*/) {
			var start, end, arg = [], i;
			if (arguments.length) {
				if (arguments.length > 1) {
					for (i = 1; i < arguments.length; ++i)
						arg.push(arguments[i]);
					start = new Date();
					fn.apply(window, arg);
					end = new Date();
				} else {
					start = new Date();
					fn();
					end = new Date();
				}
				return end.getTime() - start.getTime();
			}
			return undefined;
		}

		function getMatchingBorder(border1, border2) {
			// ECMA-376 Part 1 17.4.67 tcBorders (Table Cell Borders)
			if (!border1) {
				return border2;
			}
			if (!border2) {
				return border1;
			}

			if (border1.w > border2.w) {
				return border1;
			} else if (border1.w < border2.w) {
				return border2;
			}

			var bc1 = border1.getColorOrDefault();
			var bc2 = border2.getColorOrDefault();
			var r1 = bc1.getR(), g1 = bc1.getG(), b1 = bc1.getB();
			var r2 = bc2.getR(), g2 = bc2.getG(), b2 = bc2.getB();
			var Brightness_1_1 = r1 + b1 + 2 * g1;
			var Brightness_1_2 = r2 + b2 + 2 * g2;
			if (Brightness_1_1 < Brightness_1_2) {
				return border1;
			} else if (Brightness_1_1 > Brightness_1_2) {
				return border2;
			}

			var Brightness_2_1 = Brightness_1_1 - r1;
			var Brightness_2_2 = Brightness_1_2 - r2;
			if (Brightness_2_1 < Brightness_2_2) {
				return border1;
			} else if (Brightness_2_1 > Brightness_2_2) {
				return border2;
			}

			var Brightness_3_1 = g1;
			var Brightness_3_2 = g2;
			if (Brightness_3_1 < Brightness_3_2) {
				return border1;
			} else if (Brightness_3_1 > Brightness_3_2) {
				return border2;
			}

			// borders equal
			return border1;
		}

		function WordSplitting(str) {
			var trueLetter = false;
			var index = 0;
			var wordsArray = [];
			var wordsIndexArray = [];
			for (var i = 0; i < str.length; i++) {
				var nCharCode = str.charCodeAt(i);
				if (AscCommon.g_aPunctuation[nCharCode] !== undefined || nCharCode === 32 || nCharCode === 10) {
					if (trueLetter) {
						trueLetter = false;
						index++;
					}
				} else {
					if(trueLetter === false) {
					  wordsIndexArray.push(i);
					}
					trueLetter = true;
					wordsArray[index] = wordsArray[index] || "";
					wordsArray[index] = wordsArray[index] + str[i];
				}
			}
			return {
				wordsArray: wordsArray,
				wordsIndex: wordsIndexArray
			};
		}

		function replaceSpellCheckWords(cellValue, options) {
			// ToDo replace one command
			if (1 === options.indexInArray && options.replaceWith) {
				cellValue = options.replaceWith;
			} else {
				for (var i = 0; i < options.replaceWords.length; ++i) {
					cellValue = cellValue.replace(options.replaceWords[i][0], function () {
						return options.replaceWords[i][1];
					});
				}
			}
			return cellValue;
		}

		function getFindRegExp(value, options, checkEmptyVal) {
			var findFlags = "g"; // Заменяем все вхождения
			// Не чувствителен к регистру
			if (true !== options.isMatchCase) {
				findFlags += "i";
			}
			if (value === "" && checkEmptyVal) {
				value = /^$/;
			} else {
				value = value
					.replace(/(\\)/g, "\\\\").replace(/(\^)/g, "\\^")
					.replace(/(\()/g, "\\(").replace(/(\))/g, "\\)")
					.replace(/(\+)/g, "\\+").replace(/(\[)/g, "\\[")
					.replace(/(\])/g, "\\]").replace(/(\{)/g, "\\{")
					.replace(/(\})/g, "\\}").replace(/(\$)/g, "\\$")
					.replace(/(\.)/g, "\\.")
					.replace(/(~)?\*/g, function ($0, $1) {
						return $1 ? $0 : '(.*)';
					})
					.replace(/(~)?\?/g, function ($0, $1) {
						return $1 ? $0 : '.';
					})
					.replace(/(~\*)/g, "\\*").replace(/(~\?)/g, "\\?");
			}

			if (options.isWholeWord)
				value = '\\b' + value + '\\b';

			return new RegExp(value, findFlags);
		}

		function convertFillToUnifill(fill) {
			var oUniFill = null;
			if(!fill) {
				return AscFormat.CreateNoFillUniFill();
			}
			var oSF = fill.getSolidFill();
			if(oSF) {
				oUniFill = new AscFormat.CreateSolidFillRGBA(oSF.getR(), oSF.getG(), oSF.getB(), Math.min(255, 255*oSF.getA() + 0.5 >> 0));
			} else if (fill.patternFill) {
				oUniFill = new AscFormat.CUniFill();
				oUniFill.fill = new AscFormat.CPattFill();
				oUniFill.fill.ftype = fill.patternFill.getHatchOffset();
				oUniFill.fill.fgClr = AscFormat.CreateUniColorRGB2(fill.patternFill.fgColor || AscCommonExcel.createRgbColor(0, 0, 0));
				oUniFill.fill.bgClr = AscFormat.CreateUniColorRGB2(fill.patternFill.bgColor || AscCommonExcel.createRgbColor(255, 255, 255));
			} else if (fill.gradientFill) {
				oUniFill = new AscFormat.CUniFill();
				oUniFill.fill = new AscFormat.CGradFill();
				if (fill.gradientFill.type === Asc.c_oAscFillGradType.GRAD_LINEAR) {
					oUniFill.fill.lin = new AscFormat.GradLin();
					oUniFill.fill.lin.angle = fill.gradientFill.degree * 60000;
				} else {
					oUniFill.fill.path = new AscFormat.GradPath();
				}
				for (var i = 0; i < fill.gradientFill.stop.length; ++i) {
					var oGradStop = new AscFormat.CGs();
					oGradStop.pos = fill.gradientFill.stop[i].position * 100000;
					oGradStop.color = AscFormat.CreateUniColorRGB2(fill.gradientFill.stop[i].color || AscCommonExcel.createRgbColor(255, 255, 255));
					oUniFill.fill.addColor(oGradStop);
				}
			}
			return oUniFill;
		}

		function getFullHyperlinkLength(str) {
			var res = 0;
			if (!str) {
				return res;
			}

			var validStr = "ABCDEFabcdef0123456789";
			//new RegExp('/^[xX]?[0-9a-fA-F]{6}$/', 'g')
			var checkHex = function (_val) {
				if (_val !== undefined && validStr.indexOf(_val) !== -1) {
					return true;
				}
				return false;
			};


			for (var i = 0; i < str.length; i++) {
				if (str[i] === "%") {
					if (checkHex(str[i + 1]) && checkHex(str[i + 2])) {
						res++;
					} else {
						res += 3;
					}
				} else {
					res++;
				}
			}

			return res;
		}

		var referenceType = {
			A: 0,			// Absolute
			ARRC: 1,	// Absolute row; relative column
			RRAC: 2,	// Relative row; absolute column
			R: 3			// Relative
		};

		/**
		 * Rectangle region of cells
		 * @constructor
		 * @memberOf Asc
		 * @param c1 {Number} Left side of range.
		 * @param r1 {Number} Top side of range.
		 * @param c2 {Number} Right side of range (inclusively).
		 * @param r2 {Number} Bottom side of range (inclusively).
		 * @param normalize {Boolean=} Optional. If true, range will be converted to form (left,top) - (right,bottom).
		 * @return {Range}
		 */
		function Range(c1, r1, c2, r2, normalize) {
			if (!(this instanceof Range)) {
				return new Range(c1, r1, c2, r2, normalize);
			}

			/** @type Number */
			this.c1 = c1;
			/** @type Number */
			this.r1 = r1;
			/** @type Number */
			this.c2 = c2;
			/** @type Number */
			this.r2 = r2;
			this.refType1 = referenceType.R;
			this.refType2 = referenceType.R;

			return normalize ? this.normalize() : this;
		}

		Range.prototype.compareCell = function (c1, r1, c2, r2) {
			var dif = r1 - r2;
			return 0 !== dif ? dif : c1 - c2;
		};
		Range.prototype.compareByLeftTop = function (a, b) {
			return Range.prototype.compareCell(a.c1, a.r1, b.c1, b.r1);
		};
		Range.prototype.compareByRightBottom = function (a, b) {
			return Range.prototype.compareCell(a.c2, a.r2, b.c2, b.r2);
		};
		Range.prototype.assign = function (c1, r1, c2, r2, normalize) {
			this.c1 = c1;
			this.r1 = r1;
			this.c2 = c2;
			this.r2 = r2;
			return normalize ? this.normalize() : this;
		};
		Range.prototype.assign2 = function (range) {
			this.refType1 = range.refType1;
			this.refType2 = range.refType2;
			return this.assign(range.c1, range.r1, range.c2, range.r2);
		};

		Range.prototype.clone = function (normalize) {
			var oRes = new Range(this.c1, this.r1, this.c2, this.r2, normalize);
			oRes.refType1 = this.refType1;
			oRes.refType2 = this.refType2;
			return oRes;
		};

		Range.prototype.normalize = function () {
			var tmp;
			if (this.c1 > this.c2) {
				tmp = this.c1;
				this.c1 = this.c2;
				this.c2 = tmp;
			}
			if (this.r1 > this.r2) {
				tmp = this.r1;
				this.r1 = this.r2;
				this.r2 = tmp;
			}
			return this;
		};

		Range.prototype.isEqual = function (range) {
			return range && this.c1 === range.c1 && this.r1 === range.r1 && this.c2 === range.c2 && this.r2 === range.r2;
		};

		Range.prototype.isEqualCols = function (range) {
			return range && this.c1 === range.c1 && this.c2 === range.c2;
		};

		Range.prototype.isEqualRows = function (range) {
			return range && this.r1 === range.r1 && this.r2 === range.r2;
		};
		Range.prototype.isNeighbor = function (range) {
			if(this.isEqualCols(range)) {
				if(this.r2 === range.r1 - 1 || range.r2 === this.r1 - 1) {
					return true;
				}
			}
			else if(this.isEqualRows(range)) {
				if(this.c2 === range.c1 - 1 || range.c2 === this.c1 - 1) {
					return true;
				}
			}
			return false;
		};

		Range.prototype.isEqualAll = function (range) {
			return this.isEqual(range) && this.refType1 === range.refType1 && this.refType2 === range.refType2;
		};

		Range.prototype.isEqualWithOffsetRow = function (range, offsetRow) {
			return this.c1 === range.c1 && this.c2 === range.c2 &&
				this.isAbsC1() === range.isAbsC1() && this.isAbsC2() === range.isAbsC2() &&
				this.isAbsR1() === range.isAbsR1() && this.isAbsR2() === range.isAbsR2() &&
				(((this.isAbsR1() ? this.r1 === range.r1 : this.r1 + offsetRow === range.r1) &&
				(this.isAbsR2() ? this.r2 === range.r2 : this.r2 + offsetRow === range.r2)) ||
				(this.r1 === 0 && this.r2 === gc_nMaxRow0 && this.r1 === range.r1 && this.r2 === range.r2));
		};

		Range.prototype.contains = function (c, r) {
			return this.c1 <= c && c <= this.c2 && this.r1 <= r && r <= this.r2;
		};
		Range.prototype.contains2 = function (cell) {
			return this.contains(cell.col, cell.row);
		};
		Range.prototype.containsCol = function (c) {
			return this.c1 <= c && c <= this.c2;
		};
		Range.prototype.containsRow = function (r) {
			return this.r1 <= r && r <= this.r2;
		};

		Range.prototype.containsRange = function (range) {
			return this.contains(range.c1, range.r1) && this.contains(range.c2, range.r2);
		};

		Range.prototype.containsRanges = function (ranges) {
			if (ranges && ranges.length) {
				for (var i = 0; i < ranges.length; i++) {
					if (!this.containsRange(ranges[i])) {
						return false;
					}
				}
				return true;
			}
			return false;
		};

		Range.prototype.containsFirstLineRange = function (range) {
			return this.contains(range.c1, range.r1) && this.contains(range.c2, range.r1);
		};

		Range.prototype.intersection = function (range) {
			var s1 = this.clone(true), s2 = range instanceof Range ? range.clone(true) :
				new Range(range.c1, range.r1, range.c2, range.r2, true);

			if (s2.c1 > s1.c2 || s2.c2 < s1.c1 || s2.r1 > s1.r2 || s2.r2 < s1.r1) {
				return null;
			}

			return new Range(s2.c1 >= s1.c1 && s2.c1 <= s1.c2 ? s2.c1 : s1.c1, s2.r1 >= s1.r1 && s2.r1 <= s1.r2 ? s2.r1 :
				s1.r1, Math.min(s1.c2, s2.c2), Math.min(s1.r2, s2.r2));
		};

		Range.prototype.intersectionSimple = function (range) {
			var oRes = null;
			var r1 = Math.max(this.r1, range.r1);
			var c1 = Math.max(this.c1, range.c1);
			var r2 = Math.min(this.r2, range.r2);
			var c2 = Math.min(this.c2, range.c2);
			if (r1 <= r2 && c1 <= c2) {
				oRes = new Range(c1, r1, c2, r2);
			}
			return oRes;
		};

		Range.prototype.isIntersect = function (range) {
			var bRes = true;
			if (range.r2 < this.r1 || this.r2 < range.r1) {
				bRes = false;
			} else if (range.c2 < this.c1 || this.c2 < range.c1) {
				bRes = false;
			}
			return bRes;
		};

		Range.prototype.isIntersectForInsertColRow = function (range, isInsertCol) {
			var bRes = true;
			if (range.r2 < this.r1 || this.r2 < range.r1)
				bRes = false;
			else if (range.c2 < this.c1 || this.c2 < range.c1) 
				bRes = false;
			else if (isInsertCol && (this.c1 >= range.c1))
				bRes = false;
			else if (!isInsertCol && (this.r1 >= range.r1))
				bRes = false;
			return bRes;
		};

		Range.prototype.isIntersectWithRanges = function (ranges, exceptionIndex) {
			if (ranges) {
				for (var i = 0; i < ranges.length; i++) {
					if ((null == exceptionIndex || (null != exceptionIndex && i !== exceptionIndex)) && this.isIntersect(ranges[i])) {
						return true;
					}
				}
			}
			return false;
		};

		Range.prototype.isIntersectForShift = function(range, offset) {
			var isHor = offset && offset.col;
			var isDelete = offset && (offset.col < 0 || offset.row < 0);
			if (isHor) {
				if (this.r1 <= range.r1 && range.r2 <= this.r2 && this.c1 <= range.c2) {
					return true;
				} else if (isDelete && this.c1 <= range.c1 && range.c2 <= this.c2) {
					var topIn = this.r1 <= range.r1 && range.r1 <= this.r2;
					var bottomIn = this.r1 <= range.r2 && range.r2 <= this.r2;
					return topIn || bottomIn;
				}
			} else {
				if (this.c1 <= range.c1 && range.c2 <= this.c2 && this.r1 <= range.r2) {
					return true;
				} else if (isDelete && this.r1 <= range.r1 && range.r2 <= this.r2) {
					var leftIn = this.c1 <= range.c1 && range.c1 <= this.c2;
					var rightIn = this.c1 <= range.c2 && range.c2 <= this.c2;
					return leftIn || rightIn;
				}
			}
			return false;
		};

		Range.prototype.difference = function(range) {
			var res = [];
			var intersect;
			if (this.r1 > 0) {
				intersect = new Range(0, 0, gc_nMaxCol0, this.r1 - 1).intersectionSimple(range);
				if (intersect) {
					res.push(intersect);
				}
			}
			if (this.c1 > 0) {
				intersect = new Range(0, this.r1, this.c1 - 1, this.r2).intersectionSimple(range);
				if (intersect) {
					res.push(intersect);
				}
			}
			if (this.c2 < gc_nMaxCol0) {
				intersect = new Range(this.c2 + 1, this.r1, gc_nMaxCol0, this.r2).intersectionSimple(range);
				if (intersect) {
					res.push(intersect);
				}
			}
			if (this.r2 < gc_nMaxRow0) {
				intersect = new Range(0, this.r2 + 1, gc_nMaxCol0, gc_nMaxRow0).intersectionSimple(range);
				if (intersect) {
					res.push(intersect);
				}
			}
			return res;
		};
		Range.prototype.isIntersectForShiftCell = function(col, row, offset) {
			var isHor = offset && 0 != offset.col;
			if (isHor) {
				return this.r1 <= row && row <= this.r2 && this.c1 <= col;
			} else {
				return this.c1 <= col && col <= this.c2 && this.r1 <= row;
			}
		};

		Range.prototype.forShift = function(bbox, offset, bUndo) {
			var isNoDelete = true;
			var isHor = 0 != offset.col;
			var toDelete = offset.col < 0 || offset.row < 0;
			var isLastRow = this.r2 === gc_nMaxRow0;
			var isLastCol = this.c2 === gc_nMaxCol0;
			if (isHor) {
				if (toDelete) {
					if (this.c1 < bbox.c1) {
						if (this.c2 <= bbox.c2) {
							this.setOffsetLast(new AscCommon.CellBase(0, -(this.c2 - bbox.c1 + 1)));
						} else {
							this.setOffsetLast(offset);
						}
					} else if (this.c1 <= bbox.c2) {
						if (this.c2 <= bbox.c2) {
							if(!bUndo){
								var topIn = bbox.r1 <= this.r1 && this.r1 <= bbox.r2;
								var bottomIn = bbox.r1 <= this.r2 && this.r2 <= bbox.r2;
								if (topIn && bottomIn) {
									isNoDelete = false;
								} else if (topIn) {
									this.setOffsetFirst(new AscCommon.CellBase(bbox.r2 - this.r1 + 1, 0));
								} else if (bottomIn) {
									this.setOffsetLast(new AscCommon.CellBase(bbox.r1 - this.r2 - 1, 0));
								}
							}
						} else {
							this.setOffsetFirst(new AscCommon.CellBase(0, bbox.c1 - this.c1));
							this.setOffsetLast(offset);
						}
					} else {
						this.setOffset(offset);
					}
				} else {
					if (this.c1 < bbox.c1) {
						this.setOffsetLast(offset);
					} else {
						if (this.c1 + offset.col <= gc_nMaxCol0) {
							this.setOffset(offset);
						} else {
							isNoDelete = false;
						}
					}
				}
			} else {
				if (toDelete) {
					if (this.r1 < bbox.r1) {
						if (this.r2 <= bbox.r2) {
							this.setOffsetLast(new AscCommon.CellBase(-(this.r2 - bbox.r1 + 1), 0));
						} else {
							this.setOffsetLast(offset);
						}
					} else if (this.r1 <= bbox.r2) {
						if (this.r2 <= bbox.r2) {
							if(!bUndo) {
								var leftIn = bbox.c1 <= this.c1 && this.c1 <= bbox.c2;
								var rightIn = bbox.c1 <= this.c2 && this.c2 <= bbox.c2;
								if (leftIn && rightIn) {
									isNoDelete = false;
								} else if (leftIn) {
									this.setOffsetFirst(new AscCommon.CellBase(0, bbox.c2 - this.c1 + 1));
								} else if (rightIn) {
									this.setOffsetLast(new AscCommon.CellBase(0, bbox.c1 - this.c2 - 1));
								}
							}
						} else {
							this.setOffsetFirst(new AscCommon.CellBase(bbox.r1 - this.r1, 0));
							this.setOffsetLast(offset);
						}
					} else {
						this.setOffset(offset);
					}
				} else {
					if (this.r1 < bbox.r1) {
						this.setOffsetLast(offset);
					} else {
						if (this.r1 + offset.row <= gc_nMaxRow0) {
							this.setOffset(offset);
						} else {
							isNoDelete = false;
						}
					}
				}
			}
			//range sticks to the gc_nMaxRow0/gc_nMaxCol0(but not to 0) and cannot be shifted
			if(isLastRow) {
				this.r2 = gc_nMaxRow0;
			}
			if(isLastCol) {
				this.c2 = gc_nMaxCol0;
			}
			return isNoDelete;
		};

		Range.prototype.isOneCell = function () {
			return this.r1 === this.r2 && this.c1 === this.c2;
		};

		Range.prototype.isOneCol = function () {
			return this.c1 === this.c2;
		};

		Range.prototype.isOneRow = function () {
			return this.r1 === this.r2;
		};

		Range.prototype.isOnTheEdge = function (c, r) {
			return this.r1 === r || this.r2 === r || this.c1 === c || this.c2 === c;
		};

		Range.prototype.union = function (range) {
			var s1 = this.clone(true), s2 = range instanceof Range ? range.clone(true) :
				new Range(range.c1, range.r1, range.c2, range.r2, true);

			return new Range(Math.min(s1.c1, s2.c1), Math.min(s1.r1, s2.r1), Math.max(s1.c2, s2.c2), Math.max(s1.r2, s2.r2));
		};

		Range.prototype.union2 = function (range) {
			this.c1 = Math.min(this.c1, range.c1);
			this.c2 = Math.max(this.c2, range.c2);
			this.r1 = Math.min(this.r1, range.r1);
			this.r2 = Math.max(this.r2, range.r2);
		};

		Range.prototype.union3 = function (c, r) {
			this.c1 = Math.min(this.c1, c);
			this.c2 = Math.max(this.c2, c);
			this.r1 = Math.min(this.r1, r);
			this.r2 = Math.max(this.r2, r);
		};
		Range.prototype.setOffsetWithAbs = function(offset, opt_canResize, opt_circle) {
			var temp;
			var row = offset.row;
			var col = offset.col;
			//todo offset A1048576:A1 (row = 1 -> A1:A2; row = -1 -> A1048575:A1048576)
			if (0 === this.r1 && gc_nMaxRow0 === this.r2) {
				//full sheet is 1:1048576 but row is valid for it
				row = 0;
			} else if (0 === this.c1 && gc_nMaxCol0 === this.c2) {
				col = 0;
			}
			var isAbsRow1 = this.isAbsRow(this.refType1);
			var isAbsCol1 = this.isAbsCol(this.refType1);
			var isAbsRow2 = this.isAbsRow(this.refType2);
			var isAbsCol2 = this.isAbsCol(this.refType2);
			if (!isAbsRow1) {
				this.r1 += row;
				if (this.r1 < 0) {
					if (opt_circle) {
						this.r1 += gc_nMaxRow0 + 1;
					} else {
						this.r1 = 0;
						if (!opt_canResize) {
							return false;
						}
					}
				}
				if (this.r1 > gc_nMaxRow0) {
					if (opt_circle) {
						this.r1 -= gc_nMaxRow0 + 1;
					} else {
						this.r1 = gc_nMaxRow0;
						return false;
					}
				}
			}
			if (!isAbsCol1) {
				this.c1 += col;
				if (this.c1 < 0) {
					if (opt_circle) {
						this.c1 += gc_nMaxCol0 + 1;
					} else {
						this.c1 = 0;
						if (!opt_canResize) {
							return false;
						}
					}
				}
				if (this.c1 > gc_nMaxCol0) {
					if (opt_circle) {
						this.c1 -= gc_nMaxCol0 + 1;
					} else {
						this.c1 = gc_nMaxCol0;
						return false;
					}
				}
			}
			if (!isAbsRow2) {
				this.r2 += row;
				if (this.r2 < 0) {
					if (opt_circle) {
						this.r2 += gc_nMaxRow0 + 1;
					} else {
						this.r2 = 0;
						return false;
					}
				}
				if (this.r2 > gc_nMaxRow0) {
					if (opt_circle) {
						this.r2 -= gc_nMaxRow0 + 1;
					} else {
						this.r2 = gc_nMaxRow0;
						if (!opt_canResize) {
							return false;
						}
					}
				}
			}
			if (!isAbsCol2) {
				this.c2 += col;
				if (this.c2 < 0) {
					if (opt_circle) {
						this.c2 += gc_nMaxCol0 + 1;
					} else {
						this.c2 = 0;
						return false;
					}
				}
				if (this.c2 > gc_nMaxCol0) {
					if (opt_circle) {
						this.c2 -= gc_nMaxCol0 + 1;
					} else {
						this.c2 = gc_nMaxCol0;
						if (!opt_canResize) {
							return false;
						}
					}
				}
			}
			//switch abs flag
			if (this.r1 > this.r2) {
				temp = this.r1;
				this.r1 = this.r2;
				this.r2 = temp;
				if (!isAbsRow1 && isAbsRow2) {
					isAbsRow1 = !isAbsRow1;
					isAbsRow2 = !isAbsRow2;
					this.setAbs(isAbsRow1, isAbsCol1, isAbsRow2, isAbsCol2);
				}
			}
			if (this.c1 > this.c2) {
				temp = this.c1;
				this.c1 = this.c2;
				this.c2 = temp;
				if (!isAbsCol1 && isAbsCol2) {
					isAbsCol1 = !isAbsCol1;
					isAbsCol2 = !isAbsCol2;
					this.setAbs(isAbsRow1, isAbsCol1, isAbsRow2, isAbsCol2);
				}
			}
			return true;
		};
		Range.prototype.setOffset = function (offset) {
			if (this.r1 == 0 && this.r2 == gc_nMaxRow0 && offset.row != 0 ||
				this.c1 == 0 && this.c2 == gc_nMaxCol0 && offset.col != 0) {
				return;
			}
			this.setOffsetFirst(offset);
			this.setOffsetLast(offset);
		};

		Range.prototype.setOffsetFirst = function (offset) {
			this.c1 += offset.col;
			if (this.c1 < 0) {
				this.c1 = 0;
			}
			if (this.c1 > gc_nMaxCol0) {
				this.c1 = gc_nMaxCol0;
			}
			this.r1 += offset.row;
			if (this.r1 < 0) {
				this.r1 = 0;
			}
			if (this.r1 > gc_nMaxRow0) {
				this.r1 = gc_nMaxRow0;
			}
		};

		Range.prototype.setOffsetLast = function (offset) {
			this.c2 += offset.col;
			if (this.c2 < 0) {
				this.c2 = 0;
			}
			if (this.c2 > gc_nMaxCol0) {
				this.c2 = gc_nMaxCol0;
			}
			this.r2 += offset.row;
			if (this.r2 < 0) {
				this.r2 = 0;
			}
			if (this.r2 > gc_nMaxRow0) {
				this.r2 = gc_nMaxRow0;
			}
		};

		Range.prototype._getName = function (val, isCol, abs) {
			var isR1C1Mode = AscCommonExcel.g_R1C1Mode;
			val += 1;
			if (isCol && !isR1C1Mode) {
				val = g_oCellAddressUtils.colnumToColstr(val);
			}
			return (isR1C1Mode ? (isCol ? 'C' : 'R') : '') + (abs ? (isR1C1Mode ? val : '$' + val) :
				(isR1C1Mode ? ((0 !== (val = (val - (isCol ? AscCommonExcel.g_ActiveCell.c1 :
					AscCommonExcel.g_ActiveCell.r1) - 1))) ? '[' + val + ']' : '') : val));
		};
		Range.prototype.getName = function (refType) {
			var isR1C1Mode = AscCommonExcel.g_R1C1Mode;
			var c, r, type = this.getType();
			var sRes = "";
			var c1Abs, c2Abs, r1Abs, r2Abs;
			if (referenceType.A === refType) {
				c1Abs = c2Abs = r1Abs = r2Abs = true;
			} else if (referenceType.R === refType) {
				c1Abs = c2Abs = r1Abs = r2Abs = false;
			} else {
				c1Abs = this.isAbsCol(this.refType1);
				c2Abs = this.isAbsCol(this.refType2);
				r1Abs = this.isAbsRow(this.refType1);
				r2Abs = this.isAbsRow(this.refType2);
			}

			if ((c_oAscSelectionType.RangeMax === type || c_oAscSelectionType.RangeRow === type) && c1Abs === c2Abs) {
				sRes = this._getName(this.r1, false, r1Abs);
				if (this.r1 !== this.r2 || r1Abs !== r2Abs || !isR1C1Mode) {
					sRes += ':' + this._getName(this.r2, false, r2Abs);
				}
			} else if ((c_oAscSelectionType.RangeMax === type || c_oAscSelectionType.RangeCol === type) && r1Abs === r2Abs) {
				sRes = this._getName(this.c1, true, c1Abs);
				if (this.c1 !== this.c2 || c1Abs !== c2Abs || !isR1C1Mode) {
					sRes += ':' + this._getName(this.c2, true, c2Abs);
				}
			} else {
				r = this._getName(this.r1, false, r1Abs);
				c = this._getName(this.c1, true, c1Abs);
				sRes = isR1C1Mode ? r + c : c + r;

				if (!this.isOneCell() || r1Abs !== r2Abs || c1Abs !== c2Abs) {
					r = this._getName(this.r2, false, r2Abs);
					c = this._getName(this.c2, true, c2Abs);
					sRes += ':' + (isR1C1Mode ? r + c : c + r);
				}
			}
			return sRes;
		};

		Range.prototype.getAbsName = function () {
			return this.getName(referenceType.A);
		};

		Range.prototype.getType = function () {
			var bRow = 0 === this.c1 && gc_nMaxCol0 === this.c2;
			var bCol = 0 === this.r1 && gc_nMaxRow0 === this.r2;
			var res;
			if (bCol && bRow) {
				res = c_oAscSelectionType.RangeMax;
			} else if (bCol) {
				res = c_oAscSelectionType.RangeCol;
			} else if (bRow) {
				res = c_oAscSelectionType.RangeRow;
			} else {
				res = c_oAscSelectionType.RangeCells;
			}
			return res;
		};

		Range.prototype.getSharedRange = function (sharedRef, c, r) {
			var isAbsR1 = this.isAbsR1();
			var isAbsC1 = this.isAbsC1();
			var isAbsR2 = this.isAbsR2();
			var isAbsC2 = this.isAbsC2();
			if(this.r1 === 0 && this.r2 === gc_nMaxRow0){
				isAbsR1 = isAbsR2 = true;
			}
			if(this.c1 === 0 && this.c2 === gc_nMaxCol0){
				isAbsC1 = isAbsC2 = true;
			}
			var r1 = isAbsR2 ? sharedRef.r1 : Math.max(sharedRef.r2 + (r - this.r2), sharedRef.r1);
			var c1 = isAbsC2 ? sharedRef.c1 : Math.max(sharedRef.c2 + (c - this.c2), sharedRef.c1);
			var r2 = isAbsR1 ? sharedRef.r2 : Math.min(sharedRef.r1 + (r - this.r1), sharedRef.r2);
			var c2 = isAbsC1 ? sharedRef.c2 : Math.min(sharedRef.c1 + (c - this.c1), sharedRef.c2);
			return new Range(c1, r1, c2, r2);
		};
		Range.prototype.getSharedRangeBbox = function(ref, base) {
			var res = this.clone();
			var shiftBase;
			var offset = new AscCommon.CellBase(ref.r1 - base.nRow, ref.c1 - base.nCol);
			if (!offset.isEmpty()) {
				shiftBase = this.clone();
				shiftBase.setOffsetWithAbs(offset, false, false);
			}
			offset.row = ref.r2 - base.nRow;
			offset.col = ref.c2 - base.nCol;
			res.setOffsetWithAbs(offset, false, false);
			res.union2(shiftBase ? shiftBase : this);
			return res;
		};
		Range.prototype.getSharedIntersect = function(sharedRef, bbox) {
			var leftTop = this.getSharedRange(sharedRef, bbox.c1, bbox.r1);
			var rightBottom = this.getSharedRange(sharedRef, bbox.c2, bbox.r2);
			return leftTop.union(rightBottom);
		};
		Range.prototype.setAbs = function (absRow1, absCol1, absRow2, absCol2) {
			this.refType1 = (absRow1 ? 0 : 2) + (absCol1 ? 0 : 1);
			this.refType2 = (absRow2 ? 0 : 2) + (absCol2 ? 0 : 1);
		};
		Range.prototype.isAbsCol = function (refType) {
			return (refType === referenceType.A || refType === referenceType.RRAC);
		};
		Range.prototype.isAbsRow = function (refType) {
			return (refType === referenceType.A || refType === referenceType.ARRC);
		};
		Range.prototype.isAbsR1 = function () {
			return this.isAbsRow(this.refType1);
		};
		Range.prototype.isAbsC1 = function () {
			return this.isAbsCol(this.refType1);
		};
		Range.prototype.isAbsR2 = function () {
			return this.isAbsRow(this.refType2);
		};
		Range.prototype.isAbsC2 = function () {
			return this.isAbsCol(this.refType2);
		};
		Range.prototype.isAbsAll = function () {
			return this.isAbsR1() && this.isAbsC1() && this.isAbsR2() && this.isAbsC2();
		};
		Range.prototype.switchReference = function () {
			this.refType1 = (this.refType1 + 1) % 4;
			this.refType2 = (this.refType2 + 1) % 4;
		};
		Range.prototype.getWidth = function() {
			return this.c2 - this.c1 + 1;
		};
		Range.prototype.getHeight = function() {
			return this.r2 - this.r1 + 1;
		};
		Range.prototype.transpose = function(startCol, startRow) {
			if (startCol === undefined) {
				startCol = this.c1;
			}
			if (startRow === undefined) {
				startRow = this.r1;
			}
			var row0 = this.c1 - startCol + startRow;
			var col0 = this.r1 - startRow + startCol;

			return new Range(col0, row0,  col0 + (this.r2 - this.r1), row0 + (this.c2 - this.c1));
		};
		Range.prototype.sliceAfter = function(col, row) {
			let res = null;
			if (col !== this.c2) {
				if (!res) {
					res = [];
				}
				res.push(new Range(col + 1, row, this.c2, row));
			}
			if (row !== this.r2) {
				if (!res) {
					res = [];
				}
				res.push(new Range(this.c1, row + 1, this.c2, this.r2));
			}
			return res;
		};


		/**
		 *
     * @constructor
		 * @extends {Range}
     */
		function Range3D() {
			this.sheet = '';
			this.sheet2 = '';

			if (3 == arguments.length) {
				var range = arguments[0];
				Range.call(this, range.c1, range.r1, range.c2, range.r2);
				// ToDo стоит пересмотреть конструкторы.
				this.refType1 = range.refType1;
				this.refType2 = range.refType2;

				this.sheet = arguments[1];
				this.sheet2 = arguments[2];
			} else if (arguments.length > 1) {
				Range.apply(this, arguments);
			} else {
				Range.call(this, 0, 0, 0, 0);
      }
		}
		Range3D.prototype = Object.create(Range.prototype);
		Range3D.prototype.constructor = Range3D;
		Range3D.prototype.isIntersect = function () {
			var oRes = true;

			if (2 == arguments.length) {
				oRes = this.sheet === arguments[1];
			}
			return oRes && Range.prototype.isIntersect.apply(this, arguments);
		};
		Range3D.prototype.clone = function () {
			return new Range3D(Range.prototype.clone.apply(this, arguments), this.sheet, this.sheet2);
		};
		Range3D.prototype.setSheet = function (sheet, sheet2) {
			this.sheet = sheet;
			this.sheet2 = sheet2 ? sheet2 : sheet;
		};
		Range3D.prototype.getName = function () {
			return AscCommon.parserHelp.get3DRef(this.sheet, Range.prototype.getName.apply(this));
		};

		/**
		 * @constructor
		 */
		function SelectionRange(ws) {
			this.ranges = [new Range(0, 0, 0, 0)];
			this.activeCell = new AscCommon.CellBase(0, 0); // Active cell
			this.activeCellId = 0;

			this.worksheet = ws;
		}

		SelectionRange.prototype.clean = function () {
			this.ranges = [new Range(0, 0, 0, 0)];
			this.activeCellId = 0;
			this.activeCell.clean();
		};
		SelectionRange.prototype.contains = function (c, r) {
			return this.ranges.some(function (item) {
				return item.contains(c, r);
			});
		};
		SelectionRange.prototype.contains2 = function (cell) {
			return this.contains(cell.col, cell.row);
		};
		SelectionRange.prototype.containsCol = function (c) {
			return this.ranges.some(function (item) {
				return item.containsCol(c);
			});
		};
		SelectionRange.prototype.containsRow = function (r) {
			return this.ranges.some(function (item) {
				return item.containsRow(r);
			});
		};
		SelectionRange.prototype.inContains = function (ranges) {
			return this.ranges.every(function (item1) {
				return ranges.some(function (item2) {
					return item2.containsRange(item1);
				});
			});
		};
		SelectionRange.prototype.containsRange = function (range) {
			return this.ranges.some(function (item) {
				return item.containsRange(range);
			});
		};
		SelectionRange.prototype.clone = function (worksheet) {
			var res = new this.constructor();
			res.ranges = this.ranges.map(function (range) {
				return range.clone();
			});
			res.activeCell = this.activeCell.clone();
			res.activeCellId = this.activeCellId;
			res.worksheet = worksheet || this.worksheet;
			return res;
		};
		SelectionRange.prototype.isEqual = function (range) {
			if (this.activeCellId !== range.activeCellId || !this.activeCell.isEqual(range.activeCell) ||
				this.ranges.length !== range.ranges.length) {
				return false;
			}
			for (var i = 0; i < this.ranges.length; ++i) {
				if (!this.ranges[i].isEqual(range.ranges[i])) {
					return false;
				}
			}
			return true;
		};
		SelectionRange.prototype.addRange = function () {
			this.activeCellId = this.ranges.push(new Range(0, 0, 0, 0)) - 1;
			this.activeCell.clean();
		};
		SelectionRange.prototype.assign2 = function (range) {
			this.clean();
			this.getLast().assign2(range);
			this.update();
		};
		SelectionRange.prototype.union = function (range) {
			var res = this.ranges.some(function (item) {
				var success = false;
				if (item.c1 === range.c1 && item.c2 === range.c2) {
					if (range.r1 === item.r2 + 1) {
						item.r2 = range.r2;
						success = true;
					} else if (range.r2 === item.r1 - 1) {
						item.r1 = range.r1;
						success = true;
					}
				} else if (item.r1 === range.r1 && item.r2 === range.r2) {
					if (range.c1 === item.c2 + 1) {
						item.c2 = range.c2;
						success = true;
					} else if (range.c2 === item.c1 - 1) {
						item.c1 = range.c1;
						success = true;
					}
				}
				return success;
			});
			if (!res) {
				this.addRange();
				this.getLast().assign2(range);
			}
		};
		SelectionRange.prototype.getUnion = function () {
			var result = new this.constructor(this.worksheet);
			var unionRanges = function (ranges, res) {
				for (var i = 0; i < ranges.length; ++i) {
					if (0 === i) {
						res.assign2(ranges[i]);
					} else {
						res.union(ranges[i]);
					}
				}
			};
			unionRanges(this.ranges, result);

			var isUnion = true, resultTmp;
			while (isUnion && !result.isSingleRange()) {
				resultTmp = new this.constructor(this.worksheet);
				unionRanges(result.ranges, resultTmp);
				isUnion = result.ranges.length !== resultTmp.ranges.length;
				result = resultTmp;
			}
			return result;
		};
		SelectionRange.prototype.offsetCell = function (dr, dc, changeRange, fCheckSize) {
			var done, curRange, mc, incompleate;
			// Check one cell
			if (1 === this.ranges.length) {
				curRange = this.ranges[this.activeCellId];
				if (curRange.isOneCell()) {
					return 0;
				} else {
					mc = this.worksheet.getMergedByCell(this.activeCell.row, this.activeCell.col);
					if (mc && curRange.isEqual(mc)) {
						return 0;
					}
				}
			}

			var lastRow = this.activeCell.row;
			var lastCol = this.activeCell.col;
			this.activeCell.row += dr;
			this.activeCell.col += dc;

			while (!done) {
				done = true;

				curRange = this.ranges[this.activeCellId];
				if (!curRange.contains2(this.activeCell)) {
					if (dr) {
						if (0 < dr) {
							this.activeCell.row = curRange.r1;
							this.activeCell.col += 1;
						} else {
							this.activeCell.row = curRange.r2;
							this.activeCell.col -= 1;
						}
					} else {
						if (0 < dc) {
							this.activeCell.row += 1;
							this.activeCell.col = curRange.c1;
						} else {
							this.activeCell.row -= 1;
							this.activeCell.col = curRange.c2;
						}
					}

					if (!curRange.contains2(this.activeCell)) {
						if (!changeRange) {
							this.activeCell.row = lastRow;
							this.activeCell.col = lastCol;
							return -1;
						}
						if (0 < dc || 0 < dr) {
							this.activeCellId += 1;
							this.activeCellId = (this.ranges.length > this.activeCellId) ? this.activeCellId : 0;
							curRange = this.ranges[this.activeCellId];

							this.activeCell.row = curRange.r1;
							this.activeCell.col = curRange.c1;
						} else {
							this.activeCellId -= 1;
							this.activeCellId = (0 <= this.activeCellId) ? this.activeCellId : this.ranges.length - 1;
							curRange = this.ranges[this.activeCellId];

							this.activeCell.row = curRange.r2;
							this.activeCell.col = curRange.c2;
						}
					}
				}

				mc = this.worksheet.getMergedByCell(this.activeCell.row, this.activeCell.col);
				if (mc) {
					incompleate = !curRange.containsRange(mc);
					if (dc > 0 && (incompleate || this.activeCell.col > mc.c1 || this.activeCell.row !== mc.r1)) {
						// Движение слева направо
						this.activeCell.col = mc.c2 + 1;
						done = false;
					} else if (dc < 0 && (incompleate || this.activeCell.col < mc.c2 || this.activeCell.row !== mc.r1)) {
						// Движение справа налево
						this.activeCell.col = mc.c1 - 1;
						done = false;
					}
					if (dr > 0 && (incompleate || this.activeCell.row > mc.r1 || this.activeCell.col !== mc.c1)) {
						// Движение сверху вниз
						this.activeCell.row = mc.r2 + 1;
						done = false;
					} else if (dr < 0 && (incompleate || this.activeCell.row < mc.r2 || this.activeCell.col !== mc.c1)) {
						// Движение снизу вверх
						this.activeCell.row = mc.r1 - 1;
						done = false;
					}
				}
				if (!done) {
					continue;
				}

				while (this.activeCell.col >= curRange.c1 && this.activeCell.col <= curRange.c2 && fCheckSize(-1, this.activeCell.col)) {
					this.activeCell.col += dc || (dr > 0 ? +1 : -1);
					done = false;
				}
				if (!done) {
					continue;
				}

				while (this.activeCell.row >= curRange.r1 && this.activeCell.row <= curRange.r2 && fCheckSize(this.activeCell.row, -1)) {
					this.activeCell.row += dr || (dc > 0 ? +1 : -1);
					done = false;
				}

				if (!done) {
					continue;
				}

				break;
			}
			return (lastRow !== this.activeCell.row || lastCol !== this.activeCell.col) ? 1 : -1;
		};
		SelectionRange.prototype.setActiveCell = function (r, c) {
			this.activeCell.row = r;
			this.activeCell.col = c;
			this.update();
		};
		SelectionRange.prototype.validActiveCell = function () {
			var res = true;
			// Check active cell in merge cell for selection row or column (bug 36708)
			var mc = this.worksheet.getMergedByCell(this.activeCell.row, this.activeCell.col);
			if (mc) {
				var curRange = this.ranges[this.activeCellId];
				if (!curRange.containsRange(mc)) {
					if (-1 === this.offsetCell(1, 0, false, function () {return false;})) {
						res = false;
						this.activeCell.row = mc.r1;
						this.activeCell.col = mc.c1;
					}
				}
			}
			return res;
		};
		SelectionRange.prototype.getLast = function () {
			return this.ranges[this.ranges.length - 1];
		};
		SelectionRange.prototype.isSingleRange = function () {
			return 1 === this.ranges.length;
		};
		SelectionRange.prototype.update = function () {
			//меняем выделеную ячейку, если она не входит в диапазон
			//возможно, в будующем придется пределать логику, пока нет примеров, когда это работает плохо
			var range = this.ranges[this.activeCellId];
			if (!range || !range.contains(this.activeCell.col, this.activeCell.row)) {
				range = this.getLast();
				this.activeCell.col = range.c1;
				this.activeCell.row = range.r1;
				this.activeCellId = this.ranges.length - 1;
			}
		};
		SelectionRange.prototype.WriteToBinary = function(w) {
			w.WriteLong(this.ranges.length);
			for (var i = 0; i < this.ranges.length; ++i) {
				var range = this.ranges[i];
				w.WriteLong(range.c1);
				w.WriteLong(range.r1);
				w.WriteLong(range.c2);
				w.WriteLong(range.r2);
			}
			w.WriteLong(this.activeCell.row);
			w.WriteLong(this.activeCell.col);
			w.WriteLong(this.activeCellId);
		};
		SelectionRange.prototype.ReadFromBinary = function(r) {
			this.clean();
			var count = r.GetLong();
			var rangesNew = [];
			for (var i = 0; i < count; ++i) {
				var range = new Asc.Range(r.GetLong(), r.GetLong(), r.GetLong(), r.GetLong());
				rangesNew.push(range);
			}
			if (rangesNew.length > 0) {
				this.ranges = rangesNew;
			}
			this.activeCell.row = r.GetLong();
			this.activeCell.col = r.GetLong();
			this.activeCellId = r.GetLong();
			this.update();
		};
		SelectionRange.prototype.Select = function () {
			this.worksheet.selectionRange = this.clone();
			this.worksheet.workbook.handlers.trigger('updateSelection');
		};
		SelectionRange.prototype.isContainsOnlyFullRowOrCol = function (byCol) {
			var res = true;
			for (var i = 0; i < this.ranges.length; ++i) {
				var range = this.ranges[i];
				var type = range.getType();
				if (byCol && c_oAscSelectionType.RangeCol !== type) {
					res = false;
					break;
				}
				if (!byCol && c_oAscSelectionType.RangeRow !== type) {
					res = false;
					break;
				}
			}
			return res;
		};
		SelectionRange.prototype.getSize = function () {
			let res = 0;
			this.ranges.forEach(function (item) {
				res += (item.r2 - item.r1 + 1) * (item.c2 - item.c1 + 1);
			});
			return res;
		};

    /**
     *
     * @constructor
     * @extends {Range}
     */
		function ActiveRange(){
			if(1 == arguments.length)
			{
				var range = arguments[0];
				Range.call(this, range.c1, range.r1, range.c2, range.r2);
				// ToDo стоит пересмотреть конструкторы.
				this.refType1 = range.refType1;
				this.refType2 = range.refType2;
			}
			else if(arguments.length > 1)
				Range.apply(this, arguments);
			else
				Range.call(this, 0, 0, 0, 0);
			this.startCol = 0; // Активная ячейка в выделении
			this.startRow = 0; // Активная ячейка в выделении
			this._updateAdditionalData();
		}

		ActiveRange.prototype = Object.create(Range.prototype);
		ActiveRange.prototype.constructor = ActiveRange;
		ActiveRange.prototype.assign = function () {
			Range.prototype.assign.apply(this, arguments);
			this._updateAdditionalData();
			return this;
		};
		ActiveRange.prototype.assign2 = function () {
			Range.prototype.assign2.apply(this, arguments);
			this._updateAdditionalData();
			return this;
		};
		ActiveRange.prototype.clone = function(){
			var oRes = new ActiveRange(Range.prototype.clone.apply(this, arguments));
			oRes.startCol = this.startCol;
			oRes.startRow = this.startRow;
			return oRes;
		};
		ActiveRange.prototype.normalize = function () {
			Range.prototype.normalize.apply(this, arguments);
			this._updateAdditionalData();
			return this;
		};
		ActiveRange.prototype.isEqualAll = function () {
			var bRes = Range.prototype.isEqual.apply(this, arguments);
			if(bRes && arguments.length > 0)
			{
				var range = arguments[0];
				bRes = this.startCol == range.startCol && this.startRow == range.startRow;
			}
			return bRes;
		};
		ActiveRange.prototype.contains = function () {
			return Range.prototype.contains.apply(this, arguments);
		};
		ActiveRange.prototype.containsRange = function () {
			return Range.prototype.containsRange.apply(this, arguments);
		};
		ActiveRange.prototype.containsFirstLineRange = function () {
			return Range.prototype.containsFirstLineRange.apply(this, arguments);
		};
		ActiveRange.prototype.intersection = function () {
			var oRes = Range.prototype.intersection.apply(this, arguments);
			if(null != oRes)
			{
				oRes = new ActiveRange(oRes);
				oRes._updateAdditionalData();
			}
			return oRes;
		};
		ActiveRange.prototype.intersectionSimple = function () {
			var oRes = Range.prototype.intersectionSimple.apply(this, arguments);
			if(null != oRes)
			{
				oRes = new ActiveRange(oRes);
				oRes._updateAdditionalData();
			}
			return oRes;
		};
		ActiveRange.prototype.union = function () {
			var oRes = new ActiveRange(Range.prototype.union.apply(this, arguments));
			oRes._updateAdditionalData();
			return oRes;
		};
		ActiveRange.prototype.union2 = function () {
			Range.prototype.union2.apply(this, arguments);
			this._updateAdditionalData();
			return this;
		};
		ActiveRange.prototype.union3 = function () {
			Range.prototype.union3.apply(this, arguments);
			this._updateAdditionalData();
			return this;
		};
		ActiveRange.prototype.setOffset = function(offset){
			this.setOffsetFirst(offset);
			this.setOffsetLast(offset);
		};
		ActiveRange.prototype.setOffsetFirst = function(offset){
			Range.prototype.setOffsetFirst.apply(this, arguments);
			this._updateAdditionalData();
			return this;
		};
		ActiveRange.prototype.setOffsetLast = function(offset){
			Range.prototype.setOffsetLast.apply(this, arguments);
			this._updateAdditionalData();
			return this;
		};
		ActiveRange.prototype._updateAdditionalData = function(){
			//меняем выделеную ячейку, если она не входит в диапазон
			//возможно, в будующем придется пределать логику, пока нет примеров, когда это работает плохо
			if(!this.contains(this.startCol, this.startRow))
			{
				this.startCol = this.c1;
				this.startRow = this.r1;
			}
		};

    /**
     *
     * @constructor
     * @extends {Range}
     */
		function FormulaRange(){
			if(1 == arguments.length)
			{
				var range = arguments[0];
				Range.call(this, range.c1, range.r1, range.c2, range.r2);
			}
			else if(arguments.length > 1)
				Range.apply(this, arguments);
			else
				Range.call(this, 0, 0, 0, 0);

			this.refType1 = referenceType.R;
			this.refType2 = referenceType.R;
		}

		FormulaRange.prototype = Object.create(Range.prototype);
		FormulaRange.prototype.constructor = FormulaRange;
		FormulaRange.prototype.clone = function () {
			var oRes = new FormulaRange(Range.prototype.clone.apply(this, arguments));
			oRes.refType1 = this.refType1;
			oRes.refType2 = this.refType2;
			return oRes;
		};
		FormulaRange.prototype.intersection = function () {
			var oRes = Range.prototype.intersection.apply(this, arguments);
			if(null != oRes)
				oRes = new FormulaRange(oRes);
			return oRes;
		};
		FormulaRange.prototype.intersectionSimple = function () {
			var oRes = Range.prototype.intersectionSimple.apply(this, arguments);
			if(null != oRes)
				oRes = new FormulaRange(oRes);
			return oRes;
		};
		FormulaRange.prototype.union = function () {
			return new FormulaRange(Range.prototype.union.apply(this, arguments));
		};
		FormulaRange.prototype.getName = function () {
			var sRes = "";
			var c1Abs = this.isAbsCol(this.refType1), c2Abs = this.isAbsCol(this.refType2);
			var r1Abs = this.isAbsRow(this.refType1), r2Abs = this.isAbsRow(this.refType2);

			if(0 == this.c1 && gc_nMaxCol0 == this.c2)
			{
				if(r1Abs)
					sRes += "$";
				sRes += (this.r1 + 1) + ":";
				if(r2Abs)
					sRes += "$";
				sRes += (this.r2 + 1);
			}
			else if(0 == this.r1 && gc_nMaxRow0 == this.r2)
			{
				if(c1Abs)
					sRes += "$";
				sRes += g_oCellAddressUtils.colnumToColstr(this.c1 + 1) + ":";
				if(c2Abs)
					sRes += "$";
				sRes += g_oCellAddressUtils.colnumToColstr(this.c2 + 1);
			}
			else
			{
				if(c1Abs)
					sRes += "$";
				sRes += g_oCellAddressUtils.colnumToColstr(this.c1 + 1);
				if(r1Abs)
					sRes += "$";
				sRes += (this.r1 + 1);
				if(!this.isOneCell())
				{
					sRes += ":";
					if(c2Abs)
						sRes += "$";
					sRes += g_oCellAddressUtils.colnumToColstr(this.c2 + 1);
					if(r2Abs)
						sRes += "$";
					sRes += (this.r2 + 1);
				}
			}
			return sRes;
		};

		function MultiplyRange(ranges) {
			this.ranges = ranges;
		}
		MultiplyRange.prototype.clone = function() {
			return new MultiplyRange(this.ranges.slice());
		};
		MultiplyRange.prototype.union2 = function(multiplyRange) {
			if (!multiplyRange || !multiplyRange.ranges) {
				return;
			}
			for (var i = 0; i < multiplyRange.ranges.length; i++) {
				this.ranges.push(multiplyRange.ranges[i]);
			}
		};
		MultiplyRange.prototype.isIntersect = function(range) {
			for (var i = 0; i < this.ranges.length; ++i) {
				if (range.isIntersect(this.ranges[i])) {
					return true;
				}
			}
			return false;
		};
		MultiplyRange.prototype.contains = function(c, r) {
			for (var i = 0; i < this.ranges.length; ++i) {
				if (this.ranges[i].contains(c, r)) {
					return true;
				}
			}
			return false;
		};
		MultiplyRange.prototype.getUnionRange = function() {
			if (0 === this.ranges.length) {
				return null;
			}
			var res = this.ranges[0].clone();
			for (var i = 1; i < this.ranges.length; ++i) {
				res.union2(this.ranges[i]);
			}
			return res;
		};
		MultiplyRange.prototype.unionByRowCol = function(_bCol) {
			if (0 === this.ranges.length) {
				return null;
			}
			if (1 === this.ranges.length) {
				return this.ranges;
			}
			var _ranges = this.ranges;
			var map = {};
			var res = [];
			for (var i = 0; i < _ranges.length; i++) {
				for (var j = (_bCol ? _ranges[i].c1 : _ranges[i].r1); j <= (_bCol ? _ranges[i].c2 : _ranges[i].r2); j++) {
					if (!map[j]) {
						res.push(j);
						map[j] = 1;
					}
				}
			}

			res = res.sort(function (a, b) {
				return a - b;
			});

			//объединяем
			var unionRanges = [];
			var start = null;
			var end = null;
			for (var n = 0; n < res.length; n++) {
				if (start === null) {
					start = res[n];
					end = res[n];
				} else if (res[n - 1] === res[n] - 1) {
					end++;
				} else {
					unionRanges.push(Asc.Range(_bCol ? start : 0, !_bCol ? start : 0, _bCol ? end : gc_nMaxCol0, !_bCol ? end : gc_nMaxRow0));
					start = res[n];
					end = res[n];
				}

				if (n === res.length - 1) {
					unionRanges.push(Asc.Range(_bCol ? start : 0, !_bCol ? start : 0, _bCol ? end : gc_nMaxCol0, !_bCol ? end : gc_nMaxRow0));
				}
			}

			return unionRanges.sort(function(a, b) {
				return _bCol ? b.c1 - a.c2 : b.r1 - a.r2;
			});
		};

		MultiplyRange.prototype.isNull = function() {
			if (!this.ranges || 0 === this.ranges.length || (1 === this.ranges.length && this.ranges[0] == null)) {
				return true;
			}
			return false;
		};

		function VisibleRange(visibleRange, offsetX, offsetY) {
			this.visibleRange = visibleRange;
			this.offsetX = offsetX;
			this.offsetY = offsetY;
		}

		function RangeCache() {
			this.oCache = {};
		}

		RangeCache.prototype.getAscRange = function (sRange) {
			return this._getRange(sRange, 1);
		};
		RangeCache.prototype.getRange3D = function (sRange) {
			var res = AscCommon.parserHelp.parse3DRef(sRange);
			if (!res) {
				return null;
			}
			var range = this._getRange(res.range.toUpperCase(), 1);
			return range ? new Range3D(range, res.sheet, res.sheet2) : null;
		};
		RangeCache.prototype.getActiveRange = function (sRange) {
			return this._getRange(sRange, 2);
		};
		RangeCache.prototype.getRangesFromSqRef = function (sqRef) {
			var res = [];
			var refs = sqRef.split(' ');
			for (var i = 0; i < refs.length; ++i) {
				var ref = AscCommonExcel.g_oRangeCache.getAscRange(refs[i]);
				if (ref) {
					res.push(ref.clone());
				}
			}
			return res;
		};
		RangeCache.prototype.getFormulaRange = function (sRange) {
			return this._getRange(sRange, 3);
		};
		RangeCache.prototype._getRange = function (sRange, type) {
			if (AscCommonExcel.g_R1C1Mode) {
				var o = {
					Formula: sRange, pCurrPos: 0
				};
				if (AscCommon.parserHelp.isArea.call(o, o.Formula, o.pCurrPos)) {
					sRange = o.real_str;
				} else if (AscCommon.parserHelp.isRef.call(o, o.Formula, o.pCurrPos)) {
					sRange = o.real_str;
				}
			}
			var oRes = null;
			var oCacheVal = this.oCache[sRange];
			if (null == oCacheVal) {
				var oFirstAddr, oLastAddr;
				var bIsSingle = true;
				var nIndex = sRange.indexOf(":");
				if (-1 != nIndex) {
					bIsSingle = false;
					oFirstAddr = g_oCellAddressUtils.getCellAddress(sRange.substring(0, nIndex));
					oLastAddr = g_oCellAddressUtils.getCellAddress(sRange.substring(nIndex + 1));
				} else {
					oFirstAddr = oLastAddr = g_oCellAddressUtils.getCellAddress(sRange);
				}
				oCacheVal = {first: null, last: null, ascRange: null, formulaRange: null, activeRange: null};
				//последнее условие, чтобы не распознавалось "A", "1"(должно быть "A:A", "1:1")
				if (oFirstAddr.isValid() && oLastAddr.isValid() &&
					(!bIsSingle || (!oFirstAddr.getIsRow() && !oFirstAddr.getIsCol()))) {
					oCacheVal.first = oFirstAddr;
					oCacheVal.last = oLastAddr;
				}
				this.oCache[sRange] = oCacheVal;
			}
			if (1 == type) {
				oRes = oCacheVal.ascRange;
			} else if (2 == type) {
				oRes = oCacheVal.activeRange;
			} else {
				oRes = oCacheVal.formulaRange;
			}
			if (null == oRes && null != oCacheVal.first && null != oCacheVal.last) {
				var r1 = oCacheVal.first.getRow0(), r2 = oCacheVal.last.getRow0(), c1 = oCacheVal.first.getCol0(), c2 = oCacheVal.last.getCol0();
				var r1Abs = oCacheVal.first.getRowAbs(), r2Abs = oCacheVal.last.getRowAbs(),
					c1Abs = oCacheVal.first.getColAbs(), c2Abs = oCacheVal.last.getColAbs();
				if (oCacheVal.first.getIsRow() && oCacheVal.last.getIsRow()) {
					c1 = 0;
					c2 = gc_nMaxCol0;
				}
				if (oCacheVal.first.getIsCol() && oCacheVal.last.getIsCol()) {
					r1 = 0;
					r2 = gc_nMaxRow0;
				}
				if (r1 > r2) {
					var temp = r1;
					r1 = r2;
					r2 = temp;
					temp = r1Abs;
					r1Abs = r2Abs;
					r2Abs = temp;
				}
				if (c1 > c2) {
					var temp = c1;
					c1 = c2;
					c2 = temp;
					temp = c1Abs;
					c1Abs = c2Abs;
					c2Abs = temp;
				}

				if (1 == type) {
					if (null == oCacheVal.ascRange) {
						var oAscRange = new Range(c1, r1, c2, r2);
						oAscRange.setAbs(r1Abs, c1Abs, r2Abs, c2Abs);

						oCacheVal.ascRange = oAscRange;
					}
					oRes = oCacheVal.ascRange;
				} else if (2 == type) {
					if (null == oCacheVal.activeRange) {
						var oActiveRange = new ActiveRange(c1, r1, c2, r2);
						oActiveRange.setAbs(r1Abs, c1Abs, r2Abs, c2Abs);
						oActiveRange.startCol = oActiveRange.c1;
						oActiveRange.startRow = oActiveRange.r1;
						oCacheVal.activeRange = oActiveRange;
					}
					oRes = oCacheVal.activeRange;
				} else {
					if (null == oCacheVal.formulaRange) {
						var oFormulaRange = new FormulaRange(c1, r1, c2, r2);
						oFormulaRange.setAbs(r1Abs, c1Abs, r2Abs, c2Abs);

						oCacheVal.formulaRange = oFormulaRange;
					}
					oRes = oCacheVal.formulaRange;
				}
			}
			return oRes;
		};

		var g_oRangeCache = new RangeCache();
		/**
		 * @constructor
		 * @memberOf Asc
		 */
		function HandlersList(handlers) {
			if ( !(this instanceof HandlersList) ) {return new HandlersList(handlers);}
			this.handlers = handlers || {};
			return this;
		}

		HandlersList.prototype = {

			constructor: HandlersList,

			trigger: function (eventName) {
				var h = this.handlers[eventName], t = typeOf(h), a = Array.prototype.slice.call(arguments, 1), i;
				if (t === kFunctionL) {
					return h.apply(this, a);
				}
				if (t === kArrayL) {
					for (i = 0; i < h.length; i += 1) {
						if (typeOf(h[i]) === kFunctionL) {h[i].apply(this, a);}
					}
					return true;
				}
				return false;
			},

			add: function (eventName, eventHandler, replaceOldHandler) {
				var th = this.handlers, h, old, t;
				if (replaceOldHandler || !th.hasOwnProperty(eventName)) {
					th[eventName] = eventHandler;
				} else {
					old = h = th[eventName];
					t = typeOf(old);
					if (t !== kArrayL) {
						h = th[eventName] = [];
						if (t === kFunctionL) {h.push(old);}
					}
					h.push(eventHandler);
				}
			},

			remove: function (eventName, eventHandler) {
				var th = this.handlers, h = th[eventName], i;
				if (th.hasOwnProperty(eventName)) {
					if (typeOf(h) !== kArrayL || typeOf(eventHandler) !== kFunctionL) {
						delete th[eventName];
						return true;
					}
					for (i = h.length - 1; i >= 0; i -= 1) {
						if (h[i] === eventHandler) {
							delete h[i];
							return true;
						}
					}
				}
				return false;
			}

		};


		function outputDebugStr(channel) {
			var c = window.console;
			if (Asc.g_debug_mode && c && c[channel] && c[channel].apply) {
				c[channel].apply(this, Array.prototype.slice.call(arguments, 1));
			}
		}

		function trim(val)
		{
			if(String.prototype.trim)
				return val.trim();
			else
				return val.replace(/^\s+|\s+$/g,'');
		}

		function isNumberInfinity(val) {
		    var valTrim = trim(val);
		    var valInt = valTrim - 0;
		    return valInt == valTrim && valTrim.length > 0 && MIN_EXCEL_INT < valInt && valInt < MAX_EXCEL_INT;//
		}

		function arrayToLowerCase(array) {
			var result = [];
			for (var i = 0, length = array.length; i < length; ++i)
				result.push(array[i].toLowerCase());
			return result;
		}

		function isFixedWidthCell(frag) {
			for (var i = 0; i < frag.length; ++i) {
				var f = frag[i].format;
				if (f && f.getRepeat()) {return true;}
			}
			return false;
		}

		function dropDecimalAutofit(f) {
			var s = getFragmentsText(f);
			// Проверка scientific format
			if (s.search(/E/i) >= 0) {
				return f;
			}
			// Поиск десятичной точки
			var pos = s.indexOf(AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator);
			if (-1 !== pos) {
				f = [f[0].clone()];
				f[0].setFragmentText(s.slice(0, pos));
			}
			return f;
		}

		function getFragmentsText(f) {
			return f.reduce(function (pv, cv) {
				if (null === cv.getFragmentText()) {
					cv.initText();
				}
				return pv + cv.getFragmentText();
			}, "");
		}
		function getFragmentsLength(f) {
			return f.length > 0 ? f.reduce(function (pv, cv) {
				if (null === cv.getFragmentText()) {
					cv.initText();
				}
				return pv + cv.getFragmentText().length;
			}, 0) : 0;
		}
		function getFragmentsCharCodes(f) {
			return f.reduce(function (pv, cv) {
				return pv.concat(cv.getCharCodes());
			}, "");
		}
		function getFragmentsCharCodesLength(f) {
			return f.length > 0 ? f.reduce(function (pv, cv) {
				return pv + cv.getCharCodes().length;
			}, 0) : 0;
		}
		function getFragmentsTextFromCode(f) {
			return f.reduce(function (pv, cv) {
				if (null === cv.getFragmentText()) {
					cv.initText();
				}
				return pv + cv.getFragmentText();
			}, "");
		}
		function convertUnicodeToSimpleString(sUnicode)
		{
			var sUTF16 = "";
			var nLength = sUnicode.length;
			for (var nPos = 0; nPos < nLength; nPos++)
			{
				sUTF16 += String.fromCharCode(sUnicode[nPos]);
			}

			return sUTF16;
		}

		function executeInR1C1Mode(mode, runFunction) {
			var oldMode = AscCommonExcel.g_R1C1Mode;
			AscCommonExcel.g_R1C1Mode = mode;
			runFunction();
			AscCommonExcel.g_R1C1Mode = oldMode;
		}

		function checkFilteringMode(f, oThis, args) {
			if (!window['AscCommonExcel'].filteringMode) {
				AscCommon.History.LocalChange = true;
			}
			var ret = f.apply(oThis, args);
			if (!window['AscCommonExcel'].filteringMode) {
				AscCommon.History.LocalChange = false;
			}
			return ret;
		}

		function getEndValueRange(dx, v1, v2, coord1, coord2) {
			var leftDir = {x1: v2, x2: v1},
				rightDir = {x1: v1, x2: v2},
			    res;
			if (0 !== dx) {
				if (coord1 > v1 && coord2 < v2) {
					if (0 > dx) {
						res = coord1 === coord2 ? leftDir : rightDir;
					} else {
						res = coord1 === coord2 ? rightDir : leftDir;
					}
				} else if (coord1 === v1 && coord2 === v2) {
					if (0 > dx) {
						res = leftDir;
					} else {
						res = rightDir;
					}
				} else if (coord1 > v1 && coord2 === v2) {
					res = leftDir;
				} else if (coord1 === v1 && coord2 < v2) {
					res = rightDir;
				} else {
					res = rightDir;
				}
			} else {
				res = rightDir;
			}
			return res;
		}

		function checkStylesNames(cellStyles) {
			var oStyle, i;
			for (i = 0; i < cellStyles.DefaultStyles.length; ++i) {
				oStyle = cellStyles.DefaultStyles[i];
				AscFonts.FontPickerByCharacter.getFontsByString(oStyle.Name);
				AscFonts.FontPickerByCharacter.getFontsByString(AscCommon.translateManager.getValue(oStyle.Name));
			}
			for (i = 0; i < cellStyles.CustomStyles.length; ++i) {
				oStyle = cellStyles.CustomStyles[i];
				AscFonts.FontPickerByCharacter.getFontsByString(oStyle.Name);
			}
		}

		function getContext(w, h, wb) {
			var oCanvas = document.createElement('canvas');
			oCanvas.width = w;
			oCanvas.height = h;
			return new Asc.DrawingContext(
				{canvas: oCanvas, units: 0/*px*/, fmgrGraphics: wb.fmgrGraphics, font: wb.m_oFont});
		}
		function getGraphics(ctx) {
			var graphics = new AscCommon.CGraphics();
			graphics.init(ctx.ctx, ctx.getWidth(0), ctx.getHeight(0), ctx.getWidth(3), ctx.getHeight(3));
			graphics.m_oFontManager = AscCommon.g_fontManager;

			return graphics;
		}
		function generateCellStyles(w, h, wb) {
			var result = [];

			var widthWithRetina = AscCommon.AscBrowser.convertToRetinaValue(w, true);
			var heightWithRetina = AscCommon.AscBrowser.convertToRetinaValue(h, true);

			var ctx = getContext(widthWithRetina, heightWithRetina, wb);
			var oCanvas = ctx.getCanvas();
			var graphics = getGraphics(ctx);

			function addStyles(styles, type) {
				var oStyle, name, displayName;
				for (var i = 0; i < styles.length && i < 1000; ++i) {
					oStyle = styles[i];
					if (oStyle.Hidden) {
						continue;
					}
					name = displayName = oStyle.Name;
					if (type === AscCommon.c_oAscStyleImage.Default) {
						// ToDo Возможно стоит переписать немного, чтобы не пробегать каждый раз по массиву custom-стилей (нужно генерировать AllStyles)
						oStyle = cellStyles.getCustomStyleByBuiltinId(oStyle.BuiltinId) || oStyle;
						displayName = AscCommon.translateManager.getValue(name);
					} else if (null !== oStyle.BuiltinId) {
						continue;
					}

					if (window["IS_NATIVE_EDITOR"]) {
						window["native"]["BeginDrawStyle"](type, name);
					}
					drawStyle(ctx, graphics, wb.stringRender, oStyle, displayName, widthWithRetina, heightWithRetina);
					if (window["IS_NATIVE_EDITOR"]) {
						window["native"]["EndDrawStyle"]();
					} else {
						result.push(new AscCommon.CStyleImage(name, type, oCanvas.toDataURL("image/png")));
					}
				}
			}

			var cellStyles = wb.model.CellStyles;
			addStyles(cellStyles.CustomStyles, AscCommon.c_oAscStyleImage.Document);
			addStyles(cellStyles.DefaultStyles, AscCommon.c_oAscStyleImage.Default);

			return result;
		}

		function drawStyle(ctx, graphics, sr, oStyle, sStyleName, width, height, opt_cf_preview) {
			var bc = null, bs = Asc.c_oAscBorderStyles.None, isNotFirst = false; // cached border color
			ctx.clear();
			// Fill cell
			if (oStyle.ApplyFill) {
				var fill = oStyle.getFill();
				if (null !== fill) {
					AscCommonExcel.drawFillCell(ctx, graphics, fill, new AscCommon.asc_CRect(0, 0, width, height));
				}
			}

			function drawBorder(type, b, x1, y1, x2, y2) {
				if (b && b.w > 0) {
					var isStroke = false;
					var isNewColor = !AscCommonExcel.g_oColorManager.isEqual(bc, b.getColorOrDefault());
					var isNewStyle = bs !== b.s;
					if (isNotFirst && (isNewColor || isNewStyle)) {
						ctx.stroke();
						isStroke = true;
					}

					if (isNewColor) {
						bc = b.getColorOrDefault();
						ctx.setStrokeStyle(bc);
					}
					if (isNewStyle) {
						bs = b.s;
						ctx.setLineWidth(b.w);
						ctx.setLineDash(b.getDashSegments());
					}

					if (isStroke || false === isNotFirst) {
						isNotFirst = true;
						ctx.beginPath();
					}

					switch (type) {
						case AscCommon.c_oAscBorderType.Hor:
							ctx.lineHor(x1, y1, x2);
							break;
						case AscCommon.c_oAscBorderType.Ver:
							ctx.lineVer(x1, y1, y2);
							break;
						case AscCommon.c_oAscBorderType.Diag:
							ctx.lineDiag(x1, y1, x2, y2);
							break;
					}
				}
			}

			// Draw text
			var format = oStyle.getFont().clone();
			// Для размера шрифта делаем ограничение для превью в 16pt (у Excel 18pt, но и высота превью больше 22px)
			var nSize = format.getSize();
			if (16 < format.getSize()) {
				nSize = 16;
			}

			// рисуем в пикселях
			if (window["IS_NATIVE_EDITOR"]) {
				nSize *= AscCommon.AscBrowser.retinaPixelRatio;
			}

			format.setSize(nSize);

			var tm;
			if (!opt_cf_preview) {
				tm = sr.measureString(sStyleName);
			} else {
				var cellFlags = new AscCommonExcel.CellFlags();
				cellFlags.textAlign = oStyle.xfs.align && oStyle.xfs.align.hor;

				var fragments = [];
				var tempFragment = new AscCommonExcel.Fragment();
				tempFragment.setFragmentText(sStyleName);
				tempFragment.format = format;
				fragments.push(tempFragment);
				tm = sr.measureString(fragments, cellFlags, width);
			}

			var width_padding = 4;
			if (oStyle.xfs && oStyle.xfs.align && oStyle.xfs.align.hor === AscCommon.align_Center) {
				width_padding = Asc.round(0.5 * (width - tm.width));
			}

			// Текст будем рисовать по центру (в Excel чуть по другому реализовано, у них постоянный отступ снизу)
			var textY = Asc.round(0.5 * (height - tm.height));
			if (!opt_cf_preview) {
				ctx.setFont(format);
				ctx.setFillStyle(oStyle.getFontColor() || new AscCommon.CColor(0, 0, 0));
				ctx.fillText(sStyleName, width_padding, textY + tm.baseline);
			} else {
				sr.render(ctx, width_padding, textY, tm.width, oStyle.getFontColor() || new AscCommon.CColor(0, 0, 0));
			}

			if (oStyle.ApplyBorder) {
				// borders
				var oBorders = oStyle.getBorder();
				drawBorder(AscCommon.c_oAscBorderType.Ver, oBorders.l, 0, 0, 0, height);
				drawBorder(AscCommon.c_oAscBorderType.Hor, oBorders.b, 0, height - 1, width, height - 1);
				drawBorder(AscCommon.c_oAscBorderType.Ver, oBorders.r, width - 1, height, width - 1, 0);
				drawBorder(AscCommon.c_oAscBorderType.Hor, oBorders.t, width, 0, 0, 0);
				if (isNotFirst) {
					ctx.stroke();
				}
			}
		}

		function drawFillCell(ctx, graphics, fill, rect) {
			if (!fill.hasFill()) {
				return;
			}

			var solid = fill.getSolidFill();
			if (solid) {
				ctx.setFillStyle(solid).fillRect(rect._x, rect._y, rect._width, rect._height);
				return;
			}

			var vector_koef = AscCommonExcel.vector_koef / ctx.getZoom();
			if (AscCommon.AscBrowser.isCustomScaling()) {
				vector_koef /= AscCommon.AscBrowser.retinaPixelRatio;
			}
			rect._x *= vector_koef;
			rect._y *= vector_koef;
			rect._width *= vector_koef;
			rect._height *= vector_koef;
			AscFormat.ExecuteNoHistory(
				function () {
					var geometry = new AscFormat.CreateGeometry("rect");
					geometry.Recalculate(rect._width, rect._height, true);
					var oUniFill = AscCommonExcel.convertFillToUnifill(fill);

					if (ctx instanceof AscCommonExcel.CPdfPrinter) {
						graphics.SaveGrState();
						var _baseTransform;
						if (!ctx.Transform) {
							_baseTransform = new AscCommon.CMatrix();
						} else {
							_baseTransform = ctx.Transform;
						}
						graphics.SetBaseTransform(_baseTransform);
					}

					graphics.SaveGrState();
					var oMatrix = new AscCommon.CMatrix();
					oMatrix.tx = rect._x;
					oMatrix.ty = rect._y;
					graphics.transform3(oMatrix);
					var shapeDrawer = new AscCommon.CShapeDrawer();
					shapeDrawer.Graphics = graphics;

					shapeDrawer.fromShape2(new AscFormat.CColorObj(null, oUniFill, geometry), graphics, geometry);
					shapeDrawer.draw(geometry);
					graphics.RestoreGrState();

					if (ctx instanceof AscCommonExcel.CPdfPrinter) {
						graphics.SetBaseTransform(null);
						graphics.RestoreGrState();
					}
				}, this, []
			);
		}

		function generateSlicerStyles(w, h, wb) {
			var result = [];

			w = AscCommon.AscBrowser.convertToRetinaValue(w, true);
			h = AscCommon.AscBrowser.convertToRetinaValue(h, true);

			var ctx = getContext(w, h, wb);
			var oCanvas = ctx.getCanvas();
			var graphics = getGraphics(ctx);

			function addStyles(styles, type) {
				for(var sStyleName in styles) {
					if(styles.hasOwnProperty(sStyleName)) {
						var oSlicerStyle = styles[sStyleName];
						var oTableStyle = oAllTableStyles[sStyleName];
						if (oSlicerStyle && oTableStyle) {
							drawSlicerStyle(ctx, graphics, oSlicerStyle, oTableStyle, w, h);
							result.push(new AscCommon.CStyleImage(sStyleName, type, oCanvas.toDataURL("image/png")));
						}
					}
				}
			}

			var oAllTableStyles = wb.model.TableStyles.AllStyles;
			addStyles(wb.model.SlicerStyles.CustomStyles, AscCommon.c_oAscStyleImage.Document);
			addStyles(wb.model.SlicerStyles.DefaultStyles, AscCommon.c_oAscStyleImage.Default);

			return result;
		}

		function drawSlicerStyle(ctx, graphics, oSlicerStyle, oTableStyle, width, height) {
			ctx.clear();

			var dxf, dxfWhole;

			var nBH = (height / 6)  >> 0;

			var nIns = (height / 6 / 5) >> 0;
			var nHIns = (height / 6 / 5 + 0.5) >> 0;
			//whole
			dxfWhole = oTableStyle.wholeTable && oTableStyle.wholeTable.dxf;
			var oOldFill;
			if(dxfWhole) {
				oOldFill = dxfWhole.getFill();
				if(!oOldFill) {
					var oFill = new AscCommonExcel.Fill();
					oFill.fromColor(new AscCommonExcel.RgbColor(0xFFFFFF));
					dxfWhole.setFill(oFill);
				}
			}
			drawSlicerPreviewElement(dxfWhole, null, ctx, graphics, 0, 0, width, height);
			if(dxfWhole) {
				dxfWhole.setFill(oOldFill);
			}
			dxfWhole = dxfWhole || true;

			//header
			dxf = oTableStyle.headerRow && oTableStyle.headerRow.dxf;
			drawSlicerPreviewElement(dxf, dxfWhole, ctx, graphics, 0, 0, width, nBH);

			var nPos = nBH + nIns;
			var aBT = [
				Asc.ST_slicerStyleType.selectedItemWithData,
				Asc.ST_slicerStyleType.unselectedItemWithData,
				Asc.ST_slicerStyleType.selectedItemWithNoData,
				Asc.ST_slicerStyleType.unselectedItemWithNoData
			];
			for(var nType = 0; nType < aBT.length; ++nType) {
				dxf = oSlicerStyle[aBT[nType]];
				drawSlicerPreviewElement(dxf, dxfWhole, ctx, graphics, nHIns, nPos, width - nHIns, nPos + nBH);
				nPos += (nBH + nIns);
			}
		}
		function drawSlicerPreviewElement(dxf, dxfWhole, ctx, graphics, x0, y0, x1, y1) {
			var oFill = dxf && dxf.getFill();
			if(oFill) {
				AscCommonExcel.drawFillCell(ctx, graphics, oFill, new AscCommon.asc_CRect(x0, y0, x1 - x0, y1 - y0));
			}
			var oBorder = dxf && dxf.getBorder();
			if (oBorder) {
				var oS = oBorder.l;
				if(oS && !oS.isEmpty()) {
					ctx.setStrokeStyle(oS.getColorOrDefault()).setLineWidth(1).setLineDash(oS.getDashSegments()).beginPath();
					ctx.lineVer(x0, y0, y1);
					ctx.stroke();
				}
				oS = oBorder.t;
				if(oS && !oS.isEmpty()) {
					ctx.setStrokeStyle(oS.getColorOrDefault()).setLineWidth(1).setLineDash(oS.getDashSegments()).beginPath();
					ctx.lineHor(x0 + 1, y0, x1 - 1);
					ctx.stroke();
				}
				oS = oBorder.r;
				if(oS && !oS.isEmpty()) {
					ctx.setStrokeStyle(oS.getColorOrDefault()).setLineWidth(1).setLineDash(oS.getDashSegments()).beginPath();
					ctx.lineVer(x1 - 1, y0, y1);
					ctx.stroke();
				}
				oS = oBorder.b;
				if(oS && !oS.isEmpty()) {
					ctx.setStrokeStyle(oS.getColorOrDefault()).setLineWidth(1).setLineDash(oS.getDashSegments()).beginPath();
					ctx.lineHor(x0 + 1, y1 - 1, x1 - 1);
					ctx.stroke();
				}
			}
			if (dxfWhole) {
				var nTIns = 5;
				var nTW = 8;

				nTIns = AscCommon.AscBrowser.convertToRetinaValue(nTIns, true);
				nTW = AscCommon.AscBrowser.convertToRetinaValue(nTW, true);

				var oFont = dxf && dxf.getFont() || dxfWhole && dxfWhole.getFont && dxfWhole.getFont();
				var oColor = oFont ? oFont.getColor() : new AscCommon.CColor(0, 0, 0);
				ctx.setStrokeStyle(oColor);
				ctx.setLineWidth(1);
				ctx.setLineDash([]);
				ctx.beginPath();
				ctx.lineHor(nTIns, (y0 + (y1 - y0) / 2.0 + 0.5) >> 0, nTIns + nTW);
				ctx.stroke();
			}
		}

		function generateXfsStyle(w, h, wb, xfs, text) {
			if (AscCommon.AscBrowser.isRetina) {
				w = AscCommon.AscBrowser.convertToRetinaValue(w, true);
				h = AscCommon.AscBrowser.convertToRetinaValue(h, true);
			}

			var ctx = getContext(w, h, wb);
			var oCanvas = ctx.getCanvas();
			var graphics = getGraphics(ctx);

			var oStyle = new AscCommonExcel.CCellStyle();
			oStyle.xfs = xfs;

			drawStyle(ctx, graphics, wb.stringRender, oStyle, text, w, h);
			return new AscCommon.CStyleImage(text, null, oCanvas.toDataURL("image/png"));
		}

		function createAndPutCanvas(id) {
			var parent =  document.getElementById(id);
			if (!parent)
				return;

			var w = parent.clientWidth;
			var h = parent.clientHeight;
			if (!w || !h) {
				return;
			}

			var canvas = parent.firstChild;
			if (!canvas)
			{
				canvas = document.createElement('canvas');
				canvas.style.cssText = "pointer-events: none;padding:0;margin:0;user-select:none;";
				canvas.style.width = w + "px";
				canvas.style.height = h + "px";
				parent.appendChild(canvas);
			}

			canvas.width = AscCommon.AscBrowser.convertToRetinaValue(w, true);
			canvas.height = AscCommon.AscBrowser.convertToRetinaValue(h, true);

			return canvas;
		}

		//TODO рассмотреть объединение с generateXfsStyle
		function generateXfsStyle2(id, wb, xfs, text) {
			var canvas = createAndPutCanvas(id);
			if (!canvas) {
				return;
			}
			var w = canvas.width;
			var h = canvas.height;

			var ctx = new Asc.DrawingContext({canvas: canvas, units: 0/*px*/, fmgrGraphics: wb.fmgrGraphics, font: wb.m_oFont});
			var graphics = getGraphics(ctx);

			var oStyle = new AscCommonExcel.CCellStyle();
			oStyle.xfs = xfs;

			drawStyle(ctx, graphics, wb.stringRender, oStyle, text, w, h, true);
		}

		function drawGradientPreview(id, wb, colors, _colorBorderOut, _colorBorderIn, _realPercentWidth, _indent) {
			if (!colors || !colors.length) {
				return null;
			}

			var canvas = createAndPutCanvas(id);
			if (!canvas) {
				return;
			}
			var w = canvas.width;
			var h = canvas.height;

			var ctx = new Asc.DrawingContext({canvas: canvas, units: 0/*px*/, fmgrGraphics: wb.fmgrGraphics, font: wb.m_oFont});
			var graphics = getGraphics(ctx);

			var fill = new AscCommonExcel.Fill();
			if (colors.length === 1) {
				fill.fromColor(colors[0]);
			} else {
				fill.gradientFill = new AscCommonExcel.GradientFill();
				var arrColors = [];
				for (var i = 0; i < colors.length; i++) {
					var _stop = new AscCommonExcel.GradientStop();
					_stop.position = (i + 1)/colors.length;
					_stop.color = colors[i];
					arrColors.push(_stop);
				}
				fill.gradientFill.asc_putGradientStops(arrColors);
			}

			if (!_indent) {
				_indent = 0;
			}
			var rectX = _indent;
			var rectY = _indent;
			var rectW = w - _indent * 2;
			var rectH = h - _indent * 2;
			if (_realPercentWidth) {
				if (_realPercentWidth > 0) {
					rectW = rectW * _realPercentWidth;
				} else {
					rectX = rectW - rectW * Math.abs(_realPercentWidth) + 1;
					rectW = rectW * Math.abs(_realPercentWidth) + 1;
				}
			}
			AscCommonExcel.drawFillCell(ctx, graphics, fill,  new AscCommon.asc_CRect(rectX, rectY, rectW, rectH));

			if (_colorBorderIn) {
				ctx.setLineWidth(1).setStrokeStyle(_colorBorderIn).strokeRect(rectX, rectY, rectW - 1, rectH - 1);
			}
			if (_colorBorderOut) {
				ctx.setLineWidth(1).setStrokeStyle(_colorBorderOut).strokeRect(0, 0, w - 1, h - 1);
			}
		}

		function drawIconSetPreview(id, wb, iconImgs) {
			if (!iconImgs || !iconImgs.length) {
				return null;
			}

			var canvas = createAndPutCanvas(id);
			if (!canvas) {
				return;
			}

			var ctx = new Asc.DrawingContext({canvas: canvas, units: 0/*px*/, fmgrGraphics: wb.fmgrGraphics, font: wb.m_oFont});
			var graphics = getGraphics(ctx);

			var shapeDrawer = new AscCommon.CShapeDrawer();
			shapeDrawer.Graphics = graphics;


			AscFormat.ExecuteNoHistory(
				function () {
					for (var i = 0; i < iconImgs.length; i++) {
						var img = iconImgs[i];

						if (!img) {
							continue;
						}

						var geometry = new AscFormat.CreateGeometry("rect");
						geometry.Recalculate(5, 5, true);

						var oUniFill = new AscFormat.builder_CreateBlipFill(img, "stretch");
						graphics.save();
						var oMatrix = new AscCommon.CMatrix();
						oMatrix.tx = i*5;
						oMatrix.ty = 0;
						graphics.transform3(oMatrix);

						shapeDrawer.fromShape2(new AscFormat.CColorObj(null, oUniFill, geometry), graphics, geometry);
						shapeDrawer.draw(geometry);
						graphics.restore();
					}
				}, this, []
			);
		}

		//-----------------------------------------------------------------
		// События движения мыши
		//-----------------------------------------------------------------
		/** @constructor */
		function asc_CMouseMoveData (obj) {
			if ( !(this instanceof asc_CMouseMoveData) ) {
				return new asc_CMouseMoveData(obj);
			}

			if (obj) {
				this.type = obj.type;
				this.x = obj.x;
				this.reverseX = obj.reverseX;	// Отображать комментарий слева от ячейки
				this.y = obj.y;
				this.hyperlink = obj.hyperlink;
				this.aCommentIndexes = obj.aCommentIndexes;
				this.userId = obj.userId;
				this.lockedObjectType = obj.lockedObjectType;

				// Для resize
				this.sizeCCOrPt = obj.sizeCCOrPt;
				this.sizePx = obj.sizePx;

				//Filter
				this.filter = obj.filter;

				//Tooltip
				this.tooltip = obj.tooltip;

				this.color = obj.color;

				this.placeholderType = obj.placeholderType;
			}

			return this;
		}
		asc_CMouseMoveData.prototype = {
			constructor: asc_CMouseMoveData,
			asc_getType: function () { return this.type; },
			asc_getX: function () { return this.x; },
			asc_getReverseX: function () { return this.reverseX; },
			asc_getY: function () { return this.y; },
			asc_getHyperlink: function () { return this.hyperlink; },
			asc_getCommentIndexes: function () { return this.aCommentIndexes; },
			asc_getUserId: function () { return this.userId; },
			asc_getLockedObjectType: function () { return this.lockedObjectType; },
			asc_getSizeCCOrPt: function () { return this.sizeCCOrPt; },
			asc_getSizePx: function () { return this.sizePx; },
			asc_getFilter: function () { return this.filter; },
			asc_getTooltip: function () { return this.tooltip; },
			asc_getColor: function () { return this.color; },
			asc_getPlaceholderType: function () { return this.placeholderType; }
		};

		// Гиперссылка
		/** @constructor */
		function asc_CHyperlink (obj) {
			// Класс Hyperlink из модели
			this.hyperlinkModel = null != obj ? obj : new AscCommonExcel.Hyperlink();
			// Используется только для выдачи наружу и выставлении обратно
			this.text = null;

			return this;
		}

		asc_CHyperlink.prototype.asc_getType = function () {
			return this.hyperlinkModel.getHyperlinkType();
		};
		asc_CHyperlink.prototype.asc_getHyperlinkUrl = function () {
			return this.hyperlinkModel.Hyperlink;
		};
		asc_CHyperlink.prototype.asc_getTooltip = function () {
			return this.hyperlinkModel.Tooltip;
		};
		asc_CHyperlink.prototype.asc_getLocation = function () {
			return this.hyperlinkModel.getLocation();
		};
		asc_CHyperlink.prototype.asc_getSheet = function () {
			return this.hyperlinkModel.LocationSheet;
		};
		asc_CHyperlink.prototype.asc_getRange = function () {
			return this.hyperlinkModel.getLocationRange();
		};
		asc_CHyperlink.prototype.asc_getText = function () {
			return this.text;
		};

		asc_CHyperlink.prototype.asc_setType = function (val) {
			// В принципе эта функция избыточна
			switch (val) {
				case Asc.c_oAscHyperlinkType.WebLink:
					this.hyperlinkModel.setLocation(null);
					break;
				case Asc.c_oAscHyperlinkType.RangeLink:
					this.hyperlinkModel.Hyperlink = null;
					break;
			}
		};
		asc_CHyperlink.prototype.asc_setHyperlinkUrl = function (val) {
			this.hyperlinkModel.Hyperlink = val;
		};
		asc_CHyperlink.prototype.asc_setTooltip = function (val) {
			this.hyperlinkModel.Tooltip = val ? val.slice(0, Asc.c_oAscMaxTooltipLength) : val;
		};
		asc_CHyperlink.prototype.asc_setLocation = function (val) {
			this.hyperlinkModel.setLocation(val);
		};
		asc_CHyperlink.prototype.asc_setSheet = function (val) {
			this.hyperlinkModel.setLocationSheet(val);
		};
		asc_CHyperlink.prototype.asc_setRange = function (val) {
			this.hyperlinkModel.setLocationRange(val);
		};
		asc_CHyperlink.prototype.asc_setText = function (val) {
			this.text = val;
		};

		function CPagePrint() {
			this.pageWidth = 0;
			this.pageHeight = 0;

			this.pageClipRectLeft = 0;
			this.pageClipRectTop = 0;
			this.pageClipRectWidth = 0;
			this.pageClipRectHeight = 0;

			this.pageRange = null;

			this.leftFieldInPx = 0;
			this.topFieldInPx = 0;

			this.pageGridLines = false;
			this.pageHeadings = false;

			this.indexWorksheet = -1;

			this.startOffset = 0;
			this.startOffsetPx = 0;

			this.scale = null;

			this.titleRowRange = null;
			this.titleColRange = null;
			this.titleWidth = 0;
			this.titleHeight = 0;

			return this;
		}
		CPagePrint.prototype.clone = function () {
			var res = new CPagePrint();
			res.pageWidth = this.pageWidth;
			res.pageHeight = this.pageHeight;
			res.pageClipRectLeft = this.pageClipRectLeft;
			res.pageClipRectTop = this.pageClipRectTop;
			res.pageClipRectWidth = this.pageClipRectWidth;
			res.pageClipRectHeight = this.pageClipRectHeight;

			res.pageRange = this.pageRange ? this.pageRange.clone() : null;

			res.leftFieldInPx = this.leftFieldInPx;
			res.topFieldInPx = this.topFieldInPx;

			res.pageGridLines = this.pageGridLines;
			res.pageHeadings = this.pageHeadings;

			res.indexWorksheet = this.indexWorksheet;

			res.startOffset = this.startOffset;
			res.startOffsetPx = this.startOffsetPx;

			res.scale = this.scale;

			res.titleRowRange = this.titleRowRange;
			res.titleColRange = this.titleColRange;
			res.titleWidth = this.titleWidth;
			res.titleHeight = this.titleHeight;

			return res;
		};
		function CPrintPagesData () {
			this.arrPages = [];
			this.currentIndex = 0;

			return this;
		}
		/** @constructor */
		function asc_CAdjustPrint () {
			// Вид печати
			this.printType = Asc.c_oAscPrintType.ActiveSheets;
			this.pageOptionsMap = null;
			this.ignorePrintArea = null;

			this.isOnlyFirstPage = null;
			this.nativeOptions = undefined;
			this.activeSheetsArray = null;//массив с индексами листов, которые необходимо напечатать

			this.startPageIndex = null;
			this.endPageIndex = null;

			return this;
		}
		asc_CAdjustPrint.prototype.asc_getPrintType = function () { return this.printType; };
		asc_CAdjustPrint.prototype.asc_setPrintType = function (val) { this.printType = val; };
		asc_CAdjustPrint.prototype.asc_getPageOptionsMap = function () { return this.pageOptionsMap; };
		asc_CAdjustPrint.prototype.asc_setPageOptionsMap = function (val) { this.pageOptionsMap = val; };
		asc_CAdjustPrint.prototype.asc_getIgnorePrintArea = function () { return this.ignorePrintArea; };
		asc_CAdjustPrint.prototype.asc_setIgnorePrintArea = function (val) { this.ignorePrintArea = val; };
		asc_CAdjustPrint.prototype.asc_getNativeOptions = function () { return this.nativeOptions; };
		asc_CAdjustPrint.prototype.asc_setNativeOptions = function (val) { this.nativeOptions = val; };
		asc_CAdjustPrint.prototype.asc_getActiveSheetsArray = function () { return this.activeSheetsArray; };
		asc_CAdjustPrint.prototype.asc_setActiveSheetsArray = function (val) { this.activeSheetsArray = val; };
		asc_CAdjustPrint.prototype.asc_getStartPageIndex = function () { return this.startPageIndex; };
		asc_CAdjustPrint.prototype.asc_setStartPageIndex = function (val) { this.startPageIndex = val; };
		asc_CAdjustPrint.prototype.asc_getEndPageIndex = function () { return this.endPageIndex; };
		asc_CAdjustPrint.prototype.asc_setEndPageIndex = function (val) { this.endPageIndex = val; };

		/** @constructor */
		function asc_CLockInfo () {
			this["sheetId"] = null;
			this["type"] = null;
			this["subType"] = null;
			this["guid"] = null;
			this["rangeOrObjectId"] = null;
		}

		/** @constructor */
		function asc_CCollaborativeRange (c1, r1, c2, r2) {
			this["c1"] = c1;
			this["r1"] = r1;
			this["c2"] = c2;
			this["r2"] = r2;
		}

		/** @constructor */
		function asc_CSheetViewSettings () {
			// Показывать ли сетку
			this.showGridLines = null;
			// Показывать обозначения строк и столбцов
			this.showRowColHeaders = null;

			// Закрепление области
			this.pane = null;

			//current view zoom
			this.zoomScale = 100;

			this.showZeros = null;
			this.showFormulas = null;

			this.topLeftCell = null;
			this.view = null;

			return this;
		}

		asc_CSheetViewSettings.prototype = {
			constructor: asc_CSheetViewSettings,
			clone: function () {
				var result = new asc_CSheetViewSettings();
				result.showGridLines = this.showGridLines;
				result.showRowColHeaders = this.showRowColHeaders;
				result.zoom = this.zoom;
				if (this.pane) {
					result.pane = this.pane.clone();
				}
				result.showZeros = this.showZeros;
				result.topLeftCell = this.topLeftCell;
				result.showFormulas = this.showFormulas;
				return result;
			},
			isEqual: function (settings) {
				//TODO showzeros?
				return this.asc_getShowGridLines() === settings.asc_getShowGridLines() &&
					this.asc_getShowRowColHeaders() === settings.asc_getShowRowColHeaders();
			},
			asc_getShowGridLines: function () { return false !== this.showGridLines; },
			asc_getShowRowColHeaders: function () { return false !== this.showRowColHeaders; },
			asc_getZoomScale: function () { return this.zoomScale; },
			asc_getIsFreezePane: function () { return null !== this.pane && this.pane.isInit(); },
			asc_getShowZeros: function () { return false !== this.showZeros; },
			asc_getShowFormulas: function () { return false !== this.showFormulas; },
			asc_setShowGridLines: function (val) { this.showGridLines = val; },
			asc_setShowRowColHeaders: function (val) { this.showRowColHeaders = val; },
			asc_setZoomScale: function (val) { this.zoomScale = val; },
			asc_setShowZeros: function (val) { this.showZeros = val; },
			asc_setShowFormulas: function (val) { this.showFormulas = val; }
		};

		/** @constructor */
		function asc_CPane () {
			this.state = null;
			this.topLeftCell = null;
			this.xSplit = 0;
			this.ySplit = 0;
			// CellAddress для удобства
			this.topLeftFrozenCell = null;

			return this;
		}
		asc_CPane.prototype.isInit = function () {
			return null !== this.topLeftFrozenCell;
		};
		asc_CPane.prototype.clone = function() {
			var res = new asc_CPane();
			res.state = this.state;
			res.topLeftCell = this.topLeftCell;
			res.xSplit = this.xSplit;
			res.ySplit = this.ySplit;
			res.topLeftFrozenCell = this.topLeftFrozenCell ?
				new AscCommon.CellAddress(this.topLeftFrozenCell.row, this.topLeftFrozenCell.col) : null;
			return res;
		};
		asc_CPane.prototype.init = function() {
			// ToDo Обрабатываем пока только frozen и frozenSplit
			if ((AscCommonExcel.c_oAscPaneState.Frozen === this.state || AscCommonExcel.c_oAscPaneState.FrozenSplit === this.state) &&
				(0 < this.xSplit || 0 < this.ySplit)) {
				this.topLeftFrozenCell = new AscCommon.CellAddress(this.ySplit, this.xSplit, 0);
				if (!this.topLeftFrozenCell.isValid())
					this.topLeftFrozenCell = null;
			}
		};

		function RedoObjectParam () {
			this.bIsOn = false;
			this.oChangeWorksheetUpdate = {};
			this.bUpdateWorksheetByModel = false;
			this.bOnSheetsChanged = false;
			this.oOnUpdateTabColor = {};
			this.oOnUpdateSheetViewSettings = {};
			this.bAddRemoveRowCol = false;
			this.bChangeColorScheme = false;
			this.bChangeActive = false;
			this.activeSheet = null;
		}

		/** @constructor */
		function asc_CSheetPr() {
			if (!(this instanceof asc_CSheetPr)) {
				return new asc_CSheetPr();
			}

			this.CodeName = null;
			this.EnableFormatConditionsCalculation = null;
			this.FilterMode = null;
			this.Published = null;
			this.SyncHorizontal = null;
			this.SyncRef = null;
			this.SyncVertical = null;
			this.TransitionEntry = null;
			this.TransitionEvaluation = null;

			this.TabColor = null;
			this.AutoPageBreaks = true;
			this.FitToPage = false;
			this.ApplyStyles = false;
			this.ShowOutlineSymbols = true;
			this.SummaryBelow = true;
			this.SummaryRight = true;

			return this;
		}
		asc_CSheetPr.prototype.clone = function()  {
			var res = new asc_CSheetPr();

			res.CodeName = this.CodeName;
			res.EnableFormatConditionsCalculation = this.EnableFormatConditionsCalculation;
			res.FilterMode = this.FilterMode;
			res.Published = this.Published;
			res.SyncHorizontal = this.SyncHorizontal;
			res.SyncRef = this.SyncRef;
			res.SyncVertical = this.SyncVertical;
			res.TransitionEntry = this.TransitionEntry;
			res.TransitionEvaluation = this.TransitionEvaluation;
			if (this.TabColor)
				res.TabColor = this.TabColor.clone();

			res.FitToPage = this.FitToPage;

			res.SummaryBelow = this.SummaryBelow;
			res.SummaryRight = this.SummaryRight;

			return res;
		};

		// Математическая информация о выделении
		/** @constructor */
		function asc_CSelectionMathInfo() {
			this.count = 0;
			this.countNumbers = 0;
			this.sum = null;
			this.average = null;
			this.min = null;
			this.max = null;
		}

		asc_CSelectionMathInfo.prototype = {
			constructor: asc_CSelectionMathInfo,
			asc_getCount: function () { return this.count; },
			asc_getCountNumbers: function () { return this.countNumbers; },
			asc_getSum: function () { return this.sum; },
			asc_getAverage: function () { return this.average; },
			asc_getMin: function () { return this.min; },
			asc_getMax: function () { return this.max; }
		};

		/** @constructor */
		function asc_CFindOptions() {
			this.findWhat = "";							// текст, который ищем
			this.wordsIndex = 0;                         // индекс текущего слова
			this.scanByRows = true;						// просмотр по строкам/столбцам
			this.scanForward = true;					// поиск вперед/назад
			this.isMatchCase = false;					// учитывать регистр
			this.isWholeCell = false;
			this.isWholeWord = false;
			this.isSpellCheck = false;		    // изменение вызванное в проверке орфографии
			this.scanOnOnlySheet = Asc.c_oAscSearchBy.Sheet;				// искать только на листе/в книге c_oAscSearchBy
			this.lookIn = Asc.c_oAscFindLookIn.Formulas;	// искать в формулах/значениях/примечаниях

			this.findRegExp = null;
			this.replaceWith = "";						// текст, на который заменяем (если у нас замена)
			this.isReplaceAll = false;					// заменить все (если у нас замена)

			// внутренние переменные
			this.findInSelection = false;
			this.selectionRange = null;
			this.findRange = null;
			this.findResults = null;
			this.indexInArray = 0;
			this.countFind = 0;
			this.countReplace = 0;
			this.countFindAll = 0;
			this.countReplaceAll = 0;
			this.sheetIndex = -1;
			this.error = false;

			this.isNeedRecalc = null;

			this.specificRange = null;
			this.isForMacros = null;
			this.activeCell = null;

			//если запускаем новый поиск из-за измененного документа, то присылаем последний элемент, на который
			//кликнул пользователь и далее пытаемся найти следующий/предыдущий
			this.lastSearchElem = null;

			this.isNotSearchEmptyCells = null;
		}

		asc_CFindOptions.prototype.clone = function () {
			var result = new asc_CFindOptions();
			result.wordsIndex = this.wordsIndex;
			result.findWhat = this.findWhat;
			result.scanByRows = this.scanByRows;
			result.scanForward = this.scanForward;
			result.isMatchCase = this.isMatchCase;
			result.isWholeCell = this.isWholeCell;
			result.isWholeWord = this.isWholeWord;
			result.isSpellCheck = this.isSpellCheck;
			result.scanOnOnlySheet = this.scanOnOnlySheet;
			result.lookIn = this.lookIn;


			result.replaceWith = this.replaceWith;
			result.isReplaceAll = this.isReplaceAll;

			result.findInSelection = this.findInSelection;
			result.selectionRange = this.selectionRange ? this.selectionRange.clone() : null;
			result.findRange = this.findRange ? this.findRange.clone() : null;
			result.indexInArray = this.indexInArray;
			result.countFind = this.countFind;
			result.countReplace = this.countReplace;
			result.countFindAll = this.countFindAll;
			result.countReplaceAll = this.countReplaceAll;
			result.sheetIndex = this.sheetIndex;
			result.error = this.error;

			result.specificRange = this.specificRange;
			result.lastSearchElem = this.lastSearchElem;
			result.isNotSearchEmptyCells = this.isNotSearchEmptyCells;
			result.activeCell = this.activeCell;

			return result;
		};

		asc_CFindOptions.prototype.isEqual = function (obj) {
			return obj && this.isEqual2(obj) && this.scanForward === obj.scanForward &&
				this.scanOnOnlySheet === obj.scanOnOnlySheet;
		};
		asc_CFindOptions.prototype.isEqual2 = function (obj) {
			return obj && this.findWhat === obj.findWhat && this.scanByRows === obj.scanByRows && this.isMatchCase === obj.isMatchCase && this.isWholeCell === obj.isWholeCell &&
				this.lookIn === obj.lookIn && this.specificRange == obj.specificRange && this.isNotSearchEmptyCells == obj.isNotSearchEmptyCells && this.activeCell ==
				obj.activeCell;
		};
		asc_CFindOptions.prototype.clearFindAll = function () {
			this.countFindAll = 0;
			this.countReplaceAll = 0;
			this.error = false;
		};
		asc_CFindOptions.prototype.updateFindAll = function () {
			this.countFindAll += this.countFind;
			this.countReplaceAll += this.countReplace;
		};
		asc_CFindOptions.prototype.GetText = function () {
			return this.findWhat;
		};
		asc_CFindOptions.prototype.IsMatchCase = function () {
			return this.isMatchCase;
		};
		asc_CFindOptions.prototype.IsWholeWords = function () {
			return this.isWholeWord;
		};
		asc_CFindOptions.prototype.IsWholeWords = function () {
			return this.isWholeWord;
		};

		asc_CFindOptions.prototype.asc_setFindWhat = function (val) {this.findWhat = val;};
		asc_CFindOptions.prototype.asc_setScanByRows = function (val) {this.scanByRows = val;};
		asc_CFindOptions.prototype.asc_setScanForward = function (val) {this.scanForward = val;};
		asc_CFindOptions.prototype.asc_setIsMatchCase = function (val) {this.isMatchCase = val;};
		asc_CFindOptions.prototype.asc_setIsWholeCell = function (val) {this.isWholeCell = val;};
		asc_CFindOptions.prototype.asc_setIsWholeWord = function (val) {this.isWholeWord = val;};
		asc_CFindOptions.prototype.asc_changeSingleWord = function (val) { this.isChangeSingleWord = val; };
		asc_CFindOptions.prototype.asc_setScanOnOnlySheet = function (val) {
			//TODO не стал менять native.js, поставил условие для scanOnOnlySheet
			if (val === true) {
				this.scanOnOnlySheet = Asc.c_oAscSearchBy.Sheet;
			} else if (val === false) {
				this.scanOnOnlySheet = Asc.c_oAscSearchBy.Workbook;
			} else {
				this.scanOnOnlySheet = val;
			}
		};
		asc_CFindOptions.prototype.asc_setLookIn = function (val) {this.lookIn = val;};
		asc_CFindOptions.prototype.asc_setReplaceWith = function (val) {this.replaceWith = val;};
		asc_CFindOptions.prototype.asc_setIsReplaceAll = function (val) {this.isReplaceAll = val;};
		asc_CFindOptions.prototype.asc_setSpecificRange = function (val) {this.specificRange = val;};
		asc_CFindOptions.prototype.asc_setNeedRecalc = function (val) {this.isNeedRecalc = val;};
		asc_CFindOptions.prototype.asc_setLastSearchElem = function (val) {this.lastSearchElem = val;};
		asc_CFindOptions.prototype.asc_setNotSearchEmptyCells = function (val) {this.isNotSearchEmptyCells = val;};
		asc_CFindOptions.prototype.asc_setActiveCell = function (val) {this.activeCell = val;};
		asc_CFindOptions.prototype.asc_setIsForMacros = function (val) {this.isForMacros = val;};

		/** @constructor */
		function findResults() {
			this.values = {};

			this.currentKey1 = -1;
			this.currentKey2 = -1;
			this.currentKeys1 = null;
			this.currentKeys2 = null;
		}

		findResults.prototype.isNotEmpty = function () {
			return 0 !== Object.keys(this.values).length;
		};
		findResults.prototype.contains = function (key1, key2) {
			return this.values[key1] && this.values[key1][key2];
		};
		findResults.prototype.add = function (key1, key2, cell) {
			if (!this.values[key1]) {
				this.values[key1] = {};
			}
			this.values[key1][key2] = cell;
		};
		findResults.prototype._init = function (key1, key2) {
			this.currentKey1 = key1;
			this.currentKey2 = key2;
			this.currentKeyIndex1 = -1;
			this.currentKeyIndex2 = -1;

			this.currentKeys2 = null;
			this.currentKeys1 = Object.keys(this.values).sort(AscCommon.fSortAscending);
			this.currentKeyIndex1 = this._findKey(this.currentKey1, this.currentKeys1);
			if (0 === this.currentKeys1[this.currentKeyIndex1] - this.currentKey1) {
				this.currentKeys2 = Object.keys(this.values[this.currentKey1]).sort(AscCommon.fSortAscending);
				this.currentKeyIndex2 = this._findKey(this.currentKey2, this.currentKeys2);
			}
		};
		findResults.prototype.find = function (key1, key2, forward) {
			this.forward = forward;

			if (this.currentKey1 !== key1 || this.currentKey2 !== key2) {
				this._init(key1, key2);
			}

			if (0 === this.currentKeys1.length) {
				return false;
			}

			var step = this.forward ? +1 : -1;
			this.currentKeyIndex2 += step;
			if (!this.currentKeys2 || !this.currentKeys2[this.currentKeyIndex2]) {
				this.currentKeyIndex1 += step;
				if (!this.currentKeys1[this.currentKeyIndex1]) {
					this.currentKeyIndex1 = this.forward ? 0 : this.currentKeys1.length - 1;
				}
				this.currentKey1 = this.currentKeys1[this.currentKeyIndex1] >> 0;
				this.currentKeys2 = Object.keys(this.values[this.currentKey1]).sort(AscCommon.fSortAscending);
				this.currentKeyIndex2 = this.forward ? 0 : this.currentKeys2.length - 1;
			}
			this.currentKey2 = this.currentKeys2[this.currentKeyIndex2] >> 0;
			return true;
		};
		findResults.prototype._findKey = function (key, arrayKeys) {
			var i = this.forward ? 0 : arrayKeys.length - 1;
			var step = this.forward ? +1 : -1;
			var _key;
			while (_key = arrayKeys[i]) {
				_key = step * ((_key >> 0) - key);
				if (_key >= 0) {
					return 0 === _key ? i : (i - step);
				}
				i += step;
			}
			return -2;
		};

		function CSpellcheckState() {
			this.lastSpellInfo = null;
			this.lastIndex = 0;

			this.lockSpell = false;
			this.startCell = null;
			this.currentCell = null;
			this.iteration = false;
			this.ignoreWords = {};
			this.changeWords = {};
			this.cellsChange = [];
			this.newWord = null;
			this.cellText = null;
			this.newCellText = null;
			this.isStart = false;
			this.afterReplace = false;
			this.isIgnoreUppercase = false;
			this.isIgnoreNumbers = false;
		}

		CSpellcheckState.prototype.clean = function () {
			this.isStart = false;
			this.lastSpellInfo = null;
			this.lastIndex = 0;

			this.lockSpell = false;
			this.startCell = null;
			this.currentCell = null;
			this.iteration = false;
			this.ignoreWords = {};
			this.changeWords = {};
			this.cellsChange = [];
			this.newWord = null;
			this.cellText = null;
			this.newCellText = null;
			this.afterReplace = false;
		};
		CSpellcheckState.prototype.nextRow = function () {
			this.lastSpellInfo = null;
			this.lastIndex = 0;

			this.currentCell.row += 1;
			this.currentCell.col = 0;
		};

		/** @constructor */
		function asc_CCompleteMenu(name, type) {
			this.name = name;
			this.type = type;
		}
		asc_CCompleteMenu.prototype.asc_getName = function () {return this.name;};
		asc_CCompleteMenu.prototype.asc_getType = function () {return this.type;};

		function CCacheMeasureEmpty2() {
			this.cache = {};
		}
		CCacheMeasureEmpty2.prototype.getKey = function (elem) {
			return elem.getName() + (elem.getBold() ? 'B' : 'N') + (elem.getItalic() ? 'I' : 'N');
		};
		CCacheMeasureEmpty2.prototype.add = function (elem, val) {
			this.cache[this.getKey(elem)] = val;
		};
		CCacheMeasureEmpty2.prototype.get = function (elem) {
			return this.cache[this.getKey(elem)];
		};
		var g_oCacheMeasureEmpty2 = new CCacheMeasureEmpty2();

		function CCacheMeasureEmpty() {
			this.cache = {};
		}
		CCacheMeasureEmpty.prototype.add = function (elem, val) {
			var fn = elem.getName();
			var font = (this.cache[fn] || (this.cache[fn] = {}));
			font[elem.getSize()] = val;
		};
		CCacheMeasureEmpty.prototype.get = function (elem) {
			var font = this.cache[elem.getName()];
			return font ? font[elem.getSize()] : null;
		};
		var g_oCacheMeasureEmpty = new CCacheMeasureEmpty();

		/** @constructor */
		function asc_CFormatCellsInfo() {
			this.type = Asc.c_oAscNumFormatType.General;
			this.decimalPlaces = 2;
			this.separator = false;
			this.symbol = null;
			this.currency = null;
		}
		asc_CFormatCellsInfo.prototype.asc_setType = function (val) {this.type = val;};
		asc_CFormatCellsInfo.prototype.asc_setDecimalPlaces = function (val) {this.decimalPlaces = val;};
		asc_CFormatCellsInfo.prototype.asc_setSeparator = function (val) {this.separator = val;};
		asc_CFormatCellsInfo.prototype.asc_setSymbol = function (val) {this.symbol = val;};
		asc_CFormatCellsInfo.prototype.asc_setCurrencySymbol = function (val) {this.currency = val;};
		asc_CFormatCellsInfo.prototype.asc_getType = function () {return this.type;};
		asc_CFormatCellsInfo.prototype.asc_getDecimalPlaces = function () {return this.decimalPlaces;};
		asc_CFormatCellsInfo.prototype.asc_getSeparator = function () {return this.separator;};
		asc_CFormatCellsInfo.prototype.asc_getSymbol = function () {return this.symbol;};
		asc_CFormatCellsInfo.prototype.asc_getCurrencySymbol = function () {return this.currency;};

		/**
		 * передаём в меню для того, чтобы показать иконку опций авторавертывания таблиц
		 * @constructor
		 */
		function asc_CAutoCorrectOptions(){
			this.type = null;
			this.options = [];
			this.cellCoord = null;
		}
		asc_CAutoCorrectOptions.prototype.asc_setType = function (val) {this.type = val;};
		asc_CAutoCorrectOptions.prototype.asc_setOptions = function (val) {this.options = val;};
		asc_CAutoCorrectOptions.prototype.asc_setCellCoord = function(val) { this.cellCoord = val; };
		asc_CAutoCorrectOptions.prototype.asc_getType = function () {return this.type;};
		asc_CAutoCorrectOptions.prototype.asc_getOptions = function () {return this.options;};
		asc_CAutoCorrectOptions.prototype.asc_getCellCoord = function () {return this.cellCoord;};

		function CEditorEnterOptions() {
			this.cursorPos = null;
			this.eventPos = null;
			this.focus = false;
			this.newText = null;
			this.hideCursor = false;
			this.quickInput = false;
		}

		/** @constructor */
		function cDate() {
			var bind = Function.bind;
			var unbind = bind.bind(bind);
			var date = new (unbind(Date, null).apply(null, arguments));
			date.__proto__ = cDate.prototype;
			return date;
		}

		cDate.prototype = Object.create(Date.prototype);
		cDate.prototype.constructor = cDate;
		cDate.prototype.excelNullDate1900 = Date.UTC( 1899, 11, 30, 0, 0, 0 );
		cDate.prototype.excelNullDate1904 = Date.UTC( 1904, 0, 1, 0, 0, 0 );

		cDate.prototype.getExcelNullDate = function () {
			return AscCommon.bDate1904 ? cDate.prototype.excelNullDate1904 : cDate.prototype.excelNullDate1900;
		};

		cDate.prototype.isLeapYear = function () {
			var y = this.getUTCFullYear();
			return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
		};
		cDate.prototype.isLeapYear1900 = function () {
			var y = this.getUTCFullYear();
			return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0 || 1900 === y;
		};

		cDate.prototype.getDaysInMonth = function () {
//    return arguments.callee[this.isLeapYear() ? 'L' : 'R'][this.getMonth()];
			return this.isLeapYear() ? this.getDaysInMonth.L[this.getUTCMonth()] : this.getDaysInMonth.R[this.getUTCMonth()];
		};

		// durations of months for the regular year
		cDate.prototype.getDaysInMonth.R = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
		// durations of months for the leap year
		cDate.prototype.getDaysInMonth.L = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

		cDate.prototype.getDayOfYear = function () {
			let year = Date.prototype.getUTCFullYear.call(this);
			let month = Date.prototype.getUTCMonth.call(this);
			let date = Date.prototype.getUTCDate.call(this);
			if (1899 === year && 11 === month && 30 === date) {
				return 0;
			} else if (1899 === year && 11 === month && 31 === date) {
				return 1;
			}
			let dayCount = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334];
			let dayOfYear = dayCount[month] + date;
			if (month > 1 && this.isLeapYear1900()) dayOfYear++;
			if (1900 === year && month <= 1) {
				dayOfYear++;
			}
			return dayOfYear;
		};

		cDate.prototype.truncate = function () {
			this.setUTCHours( 0, 0, 0, 0 );
			return this;
		};

		cDate.prototype.getExcelDate = function (bLocal) {
			return Math.floor( this.getExcelDateWithTime(bLocal) );
		};

		cDate.prototype.getExcelDateWithTime = function (bLocal) {
			var year = this.getUTCFullYear(), month = this.getUTCMonth(), date = this.getUTCDate(), res;
			var timeZoneOffset = bLocal ? this.getTimezoneOffset() * 60 * 1000 : 0;

			if (1900 === year && 0 === month && 0 === date) {
				res = 0;
			} else if (1900 < year || (1900 === year && 1 < month)) {
				res = (Date.UTC(year, month, date, this.getUTCHours(), this.getUTCMinutes(), this.getUTCSeconds()) - this.getExcelNullDate() - timeZoneOffset) / c_msPerDay;
			} else if (1900 === year && 1 === month && 29 === date) {
				res = 60;
			} else {
				res = (Date.UTC(year, month, date, this.getUTCHours(), this.getUTCMinutes(), this.getUTCSeconds()) - this.getExcelNullDate() - timeZoneOffset) / c_msPerDay - 1;
			}

			return res;
		};

		cDate.prototype.getExcelDateWithTime2 = function () {
			var year = Date.prototype.getUTCFullYear.call(this);
			var month = Date.prototype.getUTCMonth.call(this);
			var date = Date.prototype.getUTCDate.call(this);

			return (Date.UTC(year, month, date, this.getUTCHours(), this.getUTCMinutes(), this.getUTCSeconds()) - this.getExcelNullDate() ) / c_msPerDay;
		};

		cDate.prototype.getDateFromExcel = function ( val ) {

			val = Math.floor( val );

			return this.getDateFromExcelWithTime(val);
		};

		cDate.prototype.getDateFromExcelWithTime = function ( val ) {
			if (AscCommon.bDate1904) {
				return new cDate( val * c_msPerDay + this.getExcelNullDate() );
			} else {
				if ( val < 60 ) {
					return new cDate( val * c_msPerDay + this.getExcelNullDate() );
				} else if (val === 60) {
					return new cDate( Date.UTC( 1900, 1, 29 ) );
				} else {
					return new cDate( val * c_msPerDay + this.getExcelNullDate() );
				}
			}
		};

		cDate.prototype.getDateFromExcelWithTime2 = function ( val ) {
			return new cDate( val * c_msPerDay + this.getExcelNullDate() );
		};

		cDate.prototype.addYears = function ( counts ) {
			this.setUTCFullYear( this.getUTCFullYear() + Math.floor( counts ) );
		};

		cDate.prototype.addMonths = function ( counts ) {
			if ( this.lastDayOfMonth() ) {
				this.setUTCDate( 1 );
				this.setUTCMonth( this.getUTCMonth() + Math.floor( counts ) );
				this.setUTCDate( this.getDaysInMonth() );
			} else {
				this.setUTCMonth( this.getUTCMonth() + Math.floor( counts ) );
			}
		};

		cDate.prototype.addDays = function ( counts ) {
			this.setUTCDate( this.getUTCDate() + Math.floor( counts ) );
		};
		cDate.prototype.addDays2 = function ( counts ) {
			Date.prototype.setUTCDate.call(this, Date.prototype.getUTCDate.call(this) + Math.floor( counts ) );
		};

		cDate.prototype.lastDayOfMonth = function () {
			return this.getDaysInMonth() == this.getUTCDate();
		};
		cDate.prototype.getUTCDate = function () {
			var year = Date.prototype.getUTCFullYear.call(this);
			var month = Date.prototype.getUTCMonth.call(this);
			var date = Date.prototype.getUTCDate.call(this);

			if(1899 == year && 11 == month && 31 == date) {
				return 0;
			} else {
				return date;
			}
		};

		cDate.prototype.getUTCMonth = function () {
			var year = Date.prototype.getUTCFullYear.call(this);
			var month = Date.prototype.getUTCMonth.call(this);
			var date = Date.prototype.getUTCDate.call(this);

			if(1899 == year && 11 == month && (30 === date || 31 === date)) {
				return 0;
			} else {
				return month;
			}
		};

		cDate.prototype.getUTCFullYear = function () {
			var year = Date.prototype.getUTCFullYear.call(this);
			var month = Date.prototype.getUTCMonth.call(this);
			var date = Date.prototype.getUTCDate.call(this);

			if(1899 == year && 11 == month && (30 === date || 31 === date)) {
				return 1900;
			} else {
				return year;
			}
		};

		cDate.prototype.getDateString = function (api) {
			return api.asc_getLocaleExample(AscCommon.getShortDateFormat(), this.getExcelDate());
		};
		cDate.prototype.getTimeString = function (api) {
			return api.asc_getLocaleExample(AscCommon.getShortTimeFormat(), this.getExcelDateWithTime() - this.getTimezoneOffset()/(60*24));
		};
		cDate.prototype.fromISO8601 = function (dateStr) {
			if (dateStr.endsWith("Z")) {
				return new cDate(dateStr);
			} else {
				return new cDate(dateStr + "Z");
			}
		};
		cDate.prototype.getCurrentDate = function () {
			return this;
		}

		function getIconsForLoad() {
			return AscCommonExcel.getCFIconsForLoad().concat(AscCommonExcel.getSlicerIconsForLoad()).concat(AscCommonExcel.getPivotButtonsForLoad());
		}


		/*
		 * Export
		 * -----------------------------------------------------------------------------
		 */
		var prot;
		window['Asc'] = window['Asc'] || {};
		window['AscCommonExcel'] = window['AscCommonExcel'] || {};
		window['AscCommonExcel'].g_ActiveCell = null; // Active Cell for calculate (in R1C1 mode for relative cell)
		window['AscCommonExcel'].g_R1C1Mode = false; // No calculate in R1C1 mode
		window["AscCommonExcel"].recalcType = recalcType;
		window["AscCommonExcel"].sizePxinPt = sizePxinPt;
		window['AscCommonExcel'].c_sPerDay = c_sPerDay;
		window['AscCommonExcel'].c_msPerDay = c_msPerDay;
		window["AscCommonExcel"].applyFunction = applyFunction;
		window['AscCommonExcel'].g_IncludeNewRowColInTable = true;
		window['AscCommonExcel'].g_AutoCorrectHyperlinks = true;

		window["Asc"]["cDate"] = window["Asc"].cDate = window['AscCommonExcel'].cDate = cDate;
		prot = cDate.prototype;
		prot["getExcelDateWithTime"] = prot.getExcelDateWithTime;

		window["Asc"].typeOf = typeOf;
		window["Asc"].lastIndexOf = lastIndexOf;
		window["Asc"].search = search;
		window["Asc"].getUniqueRangeColor = getUniqueRangeColor;
		window["Asc"].getMinValueOrNull = getMinValueOrNull;
		window["Asc"].round = round;
		window["Asc"].floor = floor;
		window["Asc"].ceil = ceil;
		window["Asc"].incDecFonSize = incDecFonSize;
		window["AscCommonExcel"].calcDecades = calcDecades;
		window["AscCommonExcel"].convertPtToPx = convertPtToPx;
		window["AscCommonExcel"].convertPxToPt = convertPxToPt;
		window["Asc"].profileTime = profileTime;
		window["AscCommonExcel"].getMatchingBorder = getMatchingBorder;
		window["AscCommonExcel"].WordSplitting = WordSplitting;
		window["AscCommonExcel"].getFindRegExp = getFindRegExp;
		window["AscCommonExcel"].convertFillToUnifill = convertFillToUnifill;
		window["AscCommonExcel"].replaceSpellCheckWords = replaceSpellCheckWords;
		window["AscCommonExcel"].getFullHyperlinkLength = getFullHyperlinkLength;
		window["Asc"].outputDebugStr = outputDebugStr;
		window["Asc"].isNumberInfinity = isNumberInfinity;
		window["Asc"].trim = trim;
		window["Asc"].arrayToLowerCase = arrayToLowerCase;
		window["Asc"].isFixedWidthCell = isFixedWidthCell;
		window["AscCommonExcel"].dropDecimalAutofit = dropDecimalAutofit;
		window["AscCommonExcel"].getFragmentsText = getFragmentsText;
		window['AscCommonExcel'].getFragmentsLength = getFragmentsLength;
		window["AscCommonExcel"].getFragmentsCharCodes = getFragmentsCharCodes;
		window["AscCommonExcel"].getFragmentsCharCodesLength = getFragmentsCharCodesLength;
		window["AscCommonExcel"].convertUnicodeToSimpleString = convertUnicodeToSimpleString;
		window['AscCommonExcel'].executeInR1C1Mode = executeInR1C1Mode;
		window['AscCommonExcel'].checkFilteringMode = checkFilteringMode;
		window["Asc"].getEndValueRange = getEndValueRange;
		window["AscCommonExcel"].checkStylesNames = checkStylesNames;
		window["AscCommonExcel"].generateCellStyles = generateCellStyles;
		window["AscCommonExcel"].generateSlicerStyles = generateSlicerStyles;
		window["AscCommonExcel"].generateXfsStyle = generateXfsStyle;
		window["AscCommonExcel"].generateXfsStyle2 = generateXfsStyle2;
		window["AscCommonExcel"].getIconsForLoad = getIconsForLoad;
		window["AscCommonExcel"].drawGradientPreview = drawGradientPreview;
		window["AscCommonExcel"].drawIconSetPreview = drawIconSetPreview;

		window["Asc"]["referenceType"] = window["AscCommonExcel"].referenceType = referenceType;
		prot = referenceType;
		prot['A'] = prot.A;
		prot['ARRC'] = prot.ARRC;
		prot['RRAC'] = prot.RRAC;
		prot['R'] = prot.R;
		window["Asc"].Range = Range;
		window["AscCommonExcel"].Range3D = Range3D;
		window["AscCommonExcel"].SelectionRange = SelectionRange;
		window["AscCommonExcel"].ActiveRange = ActiveRange;
		window["AscCommonExcel"].FormulaRange = FormulaRange;
		window["AscCommonExcel"].MultiplyRange = MultiplyRange;
		window["AscCommonExcel"].VisibleRange = VisibleRange;
		window["AscCommonExcel"].g_oRangeCache = g_oRangeCache;

		window["AscCommonExcel"].HandlersList = HandlersList;

		window["AscCommonExcel"].RedoObjectParam = RedoObjectParam;

		window["AscCommonExcel"].asc_CMouseMoveData = asc_CMouseMoveData;
		prot = asc_CMouseMoveData.prototype;
		prot["asc_getType"] = prot.asc_getType;
		prot["asc_getX"] = prot.asc_getX;
		prot["asc_getReverseX"] = prot.asc_getReverseX;
		prot["asc_getY"] = prot.asc_getY;
		prot["asc_getHyperlink"] = prot.asc_getHyperlink;
		prot["asc_getCommentIndexes"] = prot.asc_getCommentIndexes;
		prot["asc_getUserId"] = prot.asc_getUserId;
		prot["asc_getLockedObjectType"] = prot.asc_getLockedObjectType;
		prot["asc_getSizeCCOrPt"] = prot.asc_getSizeCCOrPt;
		prot["asc_getSizePx"] = prot.asc_getSizePx;
		prot["asc_getFilter"] = prot.asc_getFilter;
		prot["asc_getTooltip"] = prot.asc_getTooltip;
		prot["asc_getColor"] = prot.asc_getColor;
		prot["asc_getPlaceholderType"] = prot.asc_getPlaceholderType;

		window["Asc"]["asc_CHyperlink"] = window["Asc"].asc_CHyperlink = asc_CHyperlink;
		prot = asc_CHyperlink.prototype;
		prot["asc_getType"] = prot.asc_getType;
		prot["asc_getHyperlinkUrl"] = prot.asc_getHyperlinkUrl;
		prot["asc_getTooltip"] = prot.asc_getTooltip;
		prot["asc_getLocation"] = prot.asc_getLocation;
		prot["asc_getSheet"] = prot.asc_getSheet;
		prot["asc_getRange"] = prot.asc_getRange;
		prot["asc_getText"] = prot.asc_getText;
		prot["asc_setType"] = prot.asc_setType;
		prot["asc_setHyperlinkUrl"] = prot.asc_setHyperlinkUrl;
		prot["asc_setTooltip"] = prot.asc_setTooltip;
		prot["asc_setLocation"] = prot.asc_setLocation;
		prot["asc_setSheet"] = prot.asc_setSheet;
		prot["asc_setRange"] = prot.asc_setRange;
		prot["asc_setText"] = prot.asc_setText;

		window["AscCommonExcel"].CPagePrint = CPagePrint;
		window["AscCommonExcel"].CPrintPagesData = CPrintPagesData;

		window["Asc"]["asc_CAdjustPrint"] = window["Asc"].asc_CAdjustPrint = asc_CAdjustPrint;
		prot = asc_CAdjustPrint.prototype;
		prot["asc_getPrintType"] = prot.asc_getPrintType;
		prot["asc_setPrintType"] = prot.asc_setPrintType;
		prot["asc_getPageOptionsMap"] = prot.asc_getPageOptionsMap;
		prot["asc_setPageOptionsMap"] = prot.asc_setPageOptionsMap;
		prot["asc_getIgnorePrintArea"] = prot.asc_getIgnorePrintArea;
		prot["asc_setIgnorePrintArea"] = prot.asc_setIgnorePrintArea;
		prot["asc_getNativeOptions"] = prot.asc_getNativeOptions;
		prot["asc_setNativeOptions"] = prot.asc_setNativeOptions;
		prot["asc_getActiveSheetsArray"] = prot.asc_getActiveSheetsArray;
		prot["asc_setActiveSheetsArray"] = prot.asc_setActiveSheetsArray;
		prot["asc_getStartPageIndex"] = prot.asc_getStartPageIndex;
		prot["asc_setStartPageIndex"] = prot.asc_setStartPageIndex;
		prot["asc_getEndPageIndex"] = prot.asc_getEndPageIndex;
		prot["asc_setEndPageIndex"] = prot.asc_setEndPageIndex;

		window["AscCommonExcel"].asc_CLockInfo = asc_CLockInfo;

		window["AscCommonExcel"].asc_CCollaborativeRange = asc_CCollaborativeRange;

		window["AscCommonExcel"].asc_CSheetViewSettings = asc_CSheetViewSettings;
		prot = asc_CSheetViewSettings.prototype;
		prot["asc_getShowGridLines"] = prot.asc_getShowGridLines;
		prot["asc_getShowRowColHeaders"] = prot.asc_getShowRowColHeaders;
		prot["asc_getIsFreezePane"] = prot.asc_getIsFreezePane;
		prot["asc_getShowZeros"] = prot.asc_getShowZeros;
		prot["asc_getShowFormulas"] = prot.asc_getShowFormulas;
		prot["asc_setShowGridLines"] = prot.asc_setShowGridLines;
		prot["asc_setShowRowColHeaders"] = prot.asc_setShowRowColHeaders;
		prot["asc_setShowZeros"] = prot.asc_setShowZeros;
		prot["asc_setShowFormulas"] = prot.asc_setShowFormulas;

		window["AscCommonExcel"].asc_CPane = asc_CPane;
		window["AscCommonExcel"].asc_CSheetPr = asc_CSheetPr;

		window["AscCommonExcel"].asc_CSelectionMathInfo = asc_CSelectionMathInfo;
		prot = asc_CSelectionMathInfo.prototype;
		prot["asc_getCount"] = prot.asc_getCount;
		prot["asc_getCountNumbers"] = prot.asc_getCountNumbers;
		prot["asc_getSum"] = prot.asc_getSum;
		prot["asc_getAverage"] = prot.asc_getAverage;
		prot["asc_getMin"] = prot.asc_getMin;
		prot["asc_getMax"] = prot.asc_getMax;

		window["Asc"]["asc_CFindOptions"] = window["Asc"].asc_CFindOptions = asc_CFindOptions;
		prot = asc_CFindOptions.prototype;
		prot["asc_setFindWhat"] = prot.asc_setFindWhat;
		prot["asc_setScanByRows"] = prot.asc_setScanByRows;
		prot["asc_setScanForward"] = prot.asc_setScanForward;
		prot["asc_setIsMatchCase"] = prot.asc_setIsMatchCase;
		prot["asc_setIsWholeCell"] = prot.asc_setIsWholeCell;
		prot["asc_setScanOnOnlySheet"] = prot.asc_setScanOnOnlySheet;
		prot["asc_setLookIn"] = prot.asc_setLookIn;
		prot["asc_setReplaceWith"] = prot.asc_setReplaceWith;
		prot["asc_setIsReplaceAll"] = prot.asc_setIsReplaceAll;
		prot["asc_setSpecificRange"] = prot.asc_setSpecificRange;
		prot["asc_setNeedRecalc"] = prot.asc_setNeedRecalc;
		prot["asc_setLastSearchElem"] = prot.asc_setLastSearchElem;
		prot["asc_setNotSearchEmptyCells"] = prot.asc_setNotSearchEmptyCells;
		prot["asc_setActiveCell"] = prot.asc_setActiveCell;
		prot["asc_setIsForMacros"] = prot.asc_setIsForMacros;


		window["AscCommonExcel"].findResults = findResults;

		window["AscCommonExcel"].CSpellcheckState = CSpellcheckState;

		window["AscCommonExcel"].asc_CCompleteMenu = asc_CCompleteMenu;
		prot = asc_CCompleteMenu.prototype;
		prot["asc_getName"] = prot.asc_getName;
		prot["asc_getType"] = prot.asc_getType;

		window["AscCommonExcel"].g_oCacheMeasureEmpty = g_oCacheMeasureEmpty;
		window["AscCommonExcel"].g_oCacheMeasureEmpty2 = g_oCacheMeasureEmpty2;

		window["Asc"]["asc_CFormatCellsInfo"] = window["Asc"].asc_CFormatCellsInfo = asc_CFormatCellsInfo;
		prot = asc_CFormatCellsInfo.prototype;
		prot["asc_setType"] = prot.asc_setType;
		prot["asc_setDecimalPlaces"] = prot.asc_setDecimalPlaces;
		prot["asc_setSeparator"] = prot.asc_setSeparator;
		prot["asc_setSymbol"] = prot.asc_setSymbol;
		prot["asc_setCurrencySymbol"] = prot.asc_setCurrencySymbol;
		prot["asc_getType"] = prot.asc_getType;
		prot["asc_getDecimalPlaces"] = prot.asc_getDecimalPlaces;
		prot["asc_getSeparator"] = prot.asc_getSeparator;
		prot["asc_getSymbol"] = prot.asc_getSymbol;
		prot["asc_getCurrencySymbol"] = prot.asc_getCurrencySymbol;

		window["Asc"]["asc_CAutoCorrectOptions"] = window["Asc"].asc_CAutoCorrectOptions = asc_CAutoCorrectOptions;
		prot = asc_CAutoCorrectOptions.prototype;
		prot["asc_getType"] = prot.asc_getType;
		prot["asc_getOptions"] = prot.asc_getOptions;
		prot["asc_getCellCoord"] = prot.asc_getCellCoord;

		window['AscCommonExcel'].CEditorEnterOptions = CEditorEnterOptions;
		window['AscCommonExcel'].drawFillCell = drawFillCell;
		window['AscCommonExcel'].getContext = getContext;
		window['AscCommonExcel'].getGraphics = getGraphics;
	})(window);
