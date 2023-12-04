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

(function(window, undefined)
{

	// Import
	var g_fontApplication = AscFonts.g_fontApplication;

	function CGrRFonts()
	{
		this.Ascii    = {Name : "Empty", Index : -1};
		this.EastAsia = {Name : "Empty", Index : -1};
		this.HAnsi    = {Name : "Empty", Index : -1};
		this.CS       = {Name : "Empty", Index : -1};
	}

	CGrRFonts.prototype =
	{
		checkFromTheme : function(fontScheme, rFonts)
		{
			this.Ascii.Name    = fontScheme.checkFont(rFonts.Ascii.Name);
			this.EastAsia.Name = fontScheme.checkFont(rFonts.EastAsia.Name);
			this.HAnsi.Name    = fontScheme.checkFont(rFonts.HAnsi.Name);
			this.CS.Name       = fontScheme.checkFont(rFonts.CS.Name);

			this.Ascii.Index    = -1;
			this.EastAsia.Index = -1;
			this.HAnsi.Index    = -1;
			this.CS.Index       = -1;
		},
        fromRFonts : function(rFonts)
        {
			if (rFonts.Ascii)
			{
				this.Ascii.Name = rFonts.Ascii.Name;
			}
			else
			{
				this.Ascii.Name = "Empty";
			}
			if (rFonts.EastAsia)
			{
				this.EastAsia.Name = rFonts.EastAsia.Name;
			}
			else
			{
				this.EastAsia.Name = "Empty";
			}
			if (rFonts.HAnsi)
			{
				this.HAnsi.Name = rFonts.HAnsi.Name;
			}
			else
			{
				this.HAnsi.Name = "Empty"
			}
			if (rFonts.CS)
			{
				this.CS.Name = rFonts.CS.Name;
			}
			else
			{
				this.CS.Name = "Empty";
			}

            this.Ascii.Index    = -1;
            this.EastAsia.Index = -1;
            this.HAnsi.Index    = -1;
            this.CS.Index       = -1;
        }
	};

	var gr_state_pen       = 0;
	var gr_state_brush     = 1;
	var gr_state_pen_brush = 2;
	var gr_state_state     = 3;
	var gr_state_all       = 4;

	function CFontSetup()
	{
		this.Name   = "";
		this.Index  = -1;
		this.Size   = 12;
		this.Bold   = false;
		this.Italic = false;

		this.SetUpName  = "";
		this.SetUpIndex = -1;
		this.SetUpSize  = 12;
		this.SetUpStyle = -1;

		this.SetUpMatrix = new CMatrix();
	}

	CFontSetup.prototype =
	{
		Clear : function()
		{
			this.Name   = "";
			this.Index  = -1;
			this.Size   = 12;
			this.Bold   = false;
			this.Italic = false;

			this.SetUpName  = "";
			this.SetUpIndex = -1;
			this.SetUpSize  = 12;
			this.SetUpStyle = -1;

			this.SetUpMatrix = new CMatrix();
		}
	};

	function CGrState_Pen()
	{
		this.Type = gr_state_pen;
		this.Pen  = null;
	}

	CGrState_Pen.prototype =
	{
		Init : function(_pen)
		{
			if (_pen !== undefined)
				this.Pen = _pen.CreateDublicate();
		}
	};

	function CGrState_Brush()
	{
		this.Type  = gr_state_brush;
		this.Brush = null;
	}

	CGrState_Brush.prototype =
	{
		Init : function(_brush)
		{
			if (undefined !== _brush)
				this.Brush = _brush.CreateDublicate();
		}
	};

	function CGrState_PenBrush()
	{
		this.Type  = gr_state_pen_brush;
		this.Pen   = null;
		this.Brush = null;
	}

	CGrState_PenBrush.prototype =
	{
		Init : function(_pen, _brush)
		{
			if (undefined !== _pen && undefined !== _brush)
			{
				this.Pen   = _pen.CreateDublicate();
				this.Brush = _brush.CreateDublicate();
			}
		}
	};

	function CHist_Clip()
	{
		this.Path          = null;   // clipPath
		this.Rect          = null;   // clipRect. clipRect - is a simple clipPath.
		this.IsIntegerGrid = false;
		this.Transform     = new CMatrix();
	}

	CHist_Clip.prototype =
	{
		Init : function(path, rect, isIntegerGrid, transform)
		{
			this.Path = path;

			if (rect !== undefined)
			{
				this.Rect   = new _rect();
				this.Rect.x = rect.x;
				this.Rect.y = rect.y;
				this.Rect.w = rect.w;
				this.Rect.h = rect.h;
			}

			if (undefined !== isIntegerGrid)
				this.IsIntegerGrid = isIntegerGrid;

			if (undefined !== transform)
				this.Transform = transform.CreateDublicate();
		},

		ToRenderer : function(renderer)
		{
			if (this.Rect != null)
			{
				var r = this.Rect;
				renderer.StartClipPath();
				renderer.rect(r.x, r.y, r.w, r.h);
				renderer.EndClipPath();
			}
			else
			{
				// TODO: пока не используется
			}
		}
	};

	function CGrState_State()
	{
		this.Type = gr_state_state;

		this.Transform     = null;
		this.IsIntegerGrid = false;
		this.Clips         = null;
	}

	CGrState_State.prototype =
	{
		Init : function(_transform, _isIntegerGrid, _clips)
		{
			if (undefined !== _transform)
				this.Transform = _transform.CreateDublicate();

			if (undefined !== _isIntegerGrid)
				this.IsIntegerGrid = _isIntegerGrid;

			if (undefined !== _clips)
				this.Clips = _clips;
		},

		ApplyClips : function(renderer)
		{
			var _len = this.Clips.length;
			for (var i = 0; i < _len; i++)
			{
				this.Clips[i].ToRenderer(renderer);
			}
		},

		Apply : function(parent)
		{
			for (var i = 0, len = this.Clips.length; i < len; i++)
			{
				parent.transform3(this.Clips[i].Transform);
				parent.SetIntegerGrid(this.Clips[i].IsIntegerGrid);

				var _r = this.Clips[i].Rect;

				parent.StartClipPath();

				parent._s();
				parent._m(_r.x, _r.y);
				parent._l(_r.x + _r.w, _r.y);
				parent._l(_r.x + _r.w, _r.y + _r.h);
				parent._l(_r.x, _r.y + _r.h);
				parent._l(_r.x, _r.y);

				parent.EndClipPath();
			}
		}
	};

	function CGrState()
	{
		this.Parent = null;
		this.States = [];

		this.Clips = [];
	}

	CGrState.prototype =
	{
		SavePen : function()
		{
			if (null == this.Parent)
				return;

			var _state = new CGrState_Pen();
			_state.Init(this.Parent.m_oPen);
			this.States.push(_state);
		},

		SaveBrush : function()
		{
			if (null == this.Parent)
				return;

			var _state = new CGrState_Brush();
			_state.Init(this.Parent.m_oBrush);
			this.States.push(_state);
		},

		SavePenBrush : function()
		{
			if (null == this.Parent)
				return;

			var _state = new CGrState_PenBrush();
			_state.Init(this.Parent.m_oPen, this.Parent.m_oBrush);
			this.States.push(_state);
		},

		RestorePen : function()
		{
			var _ind = this.States.length - 1;
			if (null == this.Parent || -1 == _ind)
				return;

			var _state = this.States[_ind];
			if (_state.Type == gr_state_pen)
			{
				this.States.splice(_ind, 1);
				var _c = _state.Pen.Color;
				this.Parent.p_color(_c.R, _c.G, _c.B, _c.A);
			}
		},

		RestoreBrush : function()
		{
			var _ind = this.States.length - 1;
			if (null == this.Parent || -1 == _ind)
				return;

			var _state = this.States[_ind];
			if (_state.Type == gr_state_brush)
			{
				this.States.splice(_ind, 1);
				var _c = _state.Brush.Color1;
				this.Parent.b_color1(_c.R, _c.G, _c.B, _c.A);
			}
		},

		RestorePenBrush : function()
		{
			var _ind = this.States.length - 1;
			if (null == this.Parent || -1 == _ind)
				return;

			var _state = this.States[_ind];
			if (_state.Type == gr_state_pen_brush)
			{
				this.States.splice(_ind, 1);
				var _cb = _state.Brush.Color1;
				var _cp = _state.Pen.Color;
				this.Parent.b_color1(_cb.R, _cb.G, _cb.B, _cb.A);
				this.Parent.p_color(_cp.R, _cp.G, _cp.B, _cp.A);
			}
		},

		SaveGrState : function()
		{
			if (null == this.Parent)
				return;

			var _state = new CGrState_State();
			_state.Init(this.Parent.m_oTransform, !!this.Parent.m_bIntegerGrid, this.Clips);
			this.States.push(_state);
			this.Clips = [];
		},

		RestoreGrState : function()
		{
			var _ind = this.States.length - 1;
			if (null == this.Parent || -1 == _ind)
				return;

			var _state = this.States[_ind];
			if (_state.Type === gr_state_state)
			{
				if (this.Clips.length > 0)
				{
					// значит клипы были, и их нужно обновить
					this.Parent.RemoveClip();

					for (var i = 0; i <= _ind; i++)
					{
						if (this.States[i].Type === gr_state_state)
							this.States[i].Apply(this.Parent);
					}
				}

				this.Clips = _state.Clips;
				this.States.splice(_ind, 1);

				this.Parent.transform3(_state.Transform);
				this.Parent.SetIntegerGrid(_state.IsIntegerGrid);
			}
		},

		RemoveLastClip : function()
		{
			// цель - убрать примененные this.Clips
			if (this.Clips.length === 0)
				return;

			this.lastState = new CGrState_State();
			this.lastState.Init(this.Parent.m_oTransform, !!this.Parent.m_bIntegerGrid, this.Clips);

			this.Parent.RemoveClip();
			for (var i = 0, len = this.States.length; i < len; i++)
			{
				if (this.States[i].Type === gr_state_state)
					this.States[i].Apply(this.Parent);
			}

			this.Clips = [];
			this.Parent.transform3(this.lastState.Transform);
			this.Parent.SetIntegerGrid(this.lastState.IsIntegerGrid);
		},
		RestoreLastClip : function()
		{
			// цель - вернуть примененные this.lastState.Clips
			if (!this.lastState)
				return;

			this.lastState.Apply(this.Parent);

			this.Clips = this.lastState.Clips;
			this.Parent.transform3(this.lastState.Transform);
			this.Parent.SetIntegerGrid(this.lastState.IsIntegerGrid);

			this.lastState = null;
		},

		Save : function()
		{
			this.SavePen();
			this.SaveBrush();
			this.SaveGrState();
		},

		Restore : function()
		{
			this.RestoreGrState();
			this.RestoreBrush();
			this.RestorePen();
		},

		StartClipPath : function()
		{
			// реализовать, как понадобится
		},

		EndClipPath : function()
		{
			// реализовать, как понадобится
		},

		AddClipRect : function(_r)
		{
			var _histClip           = new CHist_Clip();
			_histClip.Transform     = this.Parent.m_oTransform.CreateDublicate();
			_histClip.IsIntegerGrid = !!this.Parent.m_bIntegerGrid;
			_histClip.Rect          = new _rect();
			_histClip.Rect.x        = _r.x;
			_histClip.Rect.y        = _r.y;
			_histClip.Rect.w        = _r.w;
			_histClip.Rect.h        = _r.h;

			this.Clips.push(_histClip);

			this.Parent.StartClipPath();

			this.Parent._s();
			this.Parent._m(_r.x, _r.y);
			this.Parent._l(_r.x + _r.w, _r.y);
			this.Parent._l(_r.x + _r.w, _r.y + _r.h);
			this.Parent._l(_r.x, _r.y + _r.h);
			this.Parent._l(_r.x, _r.y);

			this.Parent.EndClipPath();

			//this.Parent._e();
		}
	};

	function CMemory(bIsNoInit)
	{
		this.Init = function(len)
		{
			var _canvas = document.createElement('canvas');
			var _ctx    = _canvas.getContext('2d');
			this.len    = (len === undefined) ? 1024 * 1024 * 5 : len;
			this.ImData = _ctx.createImageData(this.len / 4, 1);
			this.data   = this.ImData.data;
			this.pos    = 0;
		}

		this.ImData = null;
		this.data   = null;
		this.len    = 0;
		this.pos    = 0;

		this.context = null;

		if (true !== bIsNoInit)
			this.Init();

		this.Copy = function(oMemory, nPos, nLen)
		{
			for (var Index = 0; Index < nLen; Index++)
			{
				this.CheckSize(1);
				this.data[this.pos++] = oMemory.data[Index + nPos];
			}
		};

		this.CheckSize          = function(count)
		{
			if (this.pos + count >= this.len)
			{
				var _canvas = document.createElement('canvas');
				var _ctx    = _canvas.getContext('2d');

				var oldImData = this.ImData;
				var oldData   = this.data;
				var oldPos    = this.pos;

				this.len = Math.max(this.len * 2, this.pos + ((3 * count / 2) >> 0));

				this.ImData = _ctx.createImageData(this.len / 4, 1);
				this.data   = this.ImData.data;
				var newData = this.data;

				for (var i = 0; i < this.pos; i++)
					newData[i] = oldData[i];
			}
		}
		this.GetBase64Memory    = function()
		{
			return AscCommon.Base64.encode(this.data, 0, this.pos);
		}
		this.GetBase64Memory2   = function(nPos, nLen)
		{
			return AscCommon.Base64.encode(this.data, nPos, nLen);
		}
		this.sha256 = function()
		{
			let sha256 = AscCommon.Digest.sha256(this.data, 0, this.pos);
			return AscCommon.Hex.encode(sha256);
		};
		this.GetData   = function(nPos, nLen)
		{
			var _canvas = document.createElement('canvas');
			var _ctx    = _canvas.getContext('2d');

			var len = this.GetCurPosition();

			//todo ImData.data.length multiple of 4
			var ImData = _ctx.createImageData(Math.ceil(len / 4), 1);
			var res = ImData.data;

			for (var i = 0; i < len; i++)
				res[i] = this.data[i];
			return res;
		}
		this.GetDataUint8   = function(pos, len)
		{
			if (undefined === pos) {
				pos = 0;
			}
			if (undefined === len) {
				len = this.GetCurPosition() - pos;
			}
			return this.data.slice(pos, pos + len);
		}
		this.GetCurPosition     = function()
		{
			return this.pos;
		}
		this.Seek               = function(nPos)
		{
			this.pos = nPos;
		}
		this.Skip               = function(nDif)
		{
			this.pos += nDif;
		}
		this.WriteWithLen = function(_this, callback)
		{
			let oldPos = this.GetCurPosition();
			this.WriteULong(0);
			callback.call(_this, this);
			let curPos = this.GetCurPosition();
			let len = curPos - oldPos;
			this.Seek(oldPos);
			this.WriteULong(len - 4);
			this.Seek(curPos);
			return len;
		};
		this.WriteBool          = function(val)
		{
			this.CheckSize(1);
			if (false == val)
				this.data[this.pos++] = 0;
			else
				this.data[this.pos++] = 1;
		}
		this.WriteByte          = function(val)
		{
			this.CheckSize(1);
			this.data[this.pos++] = val;
		}
		this.WriteSByte         = function(val)
		{
			this.CheckSize(1);
			if (val < 0)
				val += 256;
			this.data[this.pos++] = val;
		}
		this.WriteShort          = function(val)
		{
			this.CheckSize(2);
			this.data[this.pos++] = (val) & 0xFF;
			this.data[this.pos++] = (val >>> 8) & 0xFF;
		}
		this.WriteUShort          = function(val)
		{
			this.WriteShort(AscFonts.FT_Common.UShort_To_Short(val));
		}
		this.WriteLong          = function(val)
		{
			this.CheckSize(4);
			this.data[this.pos++] = (val) & 0xFF;
			this.data[this.pos++] = (val >>> 8) & 0xFF;
			this.data[this.pos++] = (val >>> 16) & 0xFF;
			this.data[this.pos++] = (val >>> 24) & 0xFF;
		}
		this.WriteULong          = function(val)
		{
			this.WriteLong(AscFonts.FT_Common.UintToInt(val));
		}
		this.WriteDouble        = function(val)
		{
			this.CheckSize(4);
			var lval              = ((val * 100000) >> 0) & 0xFFFFFFFF; // спасаем пять знаков после запятой.
			this.data[this.pos++] = (lval) & 0xFF;
			this.data[this.pos++] = (lval >>> 8) & 0xFF;
			this.data[this.pos++] = (lval >>> 16) & 0xFF;
			this.data[this.pos++] = (lval >>> 24) & 0xFF;
		}
		var tempHelp = new ArrayBuffer(8);
		var tempHelpUnit = new Uint8Array(tempHelp);
		var tempHelpFloat = new Float64Array(tempHelp);
		this.WriteDouble2       = function(val)
		{
			this.CheckSize(8);
			tempHelpFloat[0] = val;
			this.data[this.pos++] = tempHelpUnit[0];
			this.data[this.pos++] = tempHelpUnit[1];
			this.data[this.pos++] = tempHelpUnit[2];
			this.data[this.pos++] = tempHelpUnit[3];
			this.data[this.pos++] = tempHelpUnit[4];
			this.data[this.pos++] = tempHelpUnit[5];
			this.data[this.pos++] = tempHelpUnit[6];
			this.data[this.pos++] = tempHelpUnit[7];
		}
		this._doubleEncodeLE754 = function(v)
		{
			//код взят из jspack.js на основе стандарта Little-endian N-bit IEEE 754 floating point
			var s, e, m, i, d, c, mLen, eLen, eBias, eMax;
			var el = {len : 8, mLen : 52, rt : 0};
			mLen = el.mLen, eLen = el.len * 8 - el.mLen - 1, eMax = (1 << eLen) - 1, eBias = eMax >> 1;

			s = v < 0 ? 1 : 0;
			v = Math.abs(v);
			if (isNaN(v) || (v == Infinity))
			{
				m = isNaN(v) ? 1 : 0;
				e = eMax;
			}
			else
			{
				e = Math.floor(Math.log(v) / Math.LN2);            // Calculate log2 of the value
				if (v * (c = Math.pow(2, -e)) < 1)
				{
					e--;
					c *= 2;
				}        // Math.log() isn't 100% reliable

				// Round by adding 1/2 the significand's LSD
				if (e + eBias >= 1)
				{
					v += el.rt / c;
				}            // Normalized:  mLen significand digits
				else
				{
					v += el.rt * Math.pow(2, 1 - eBias);
				}         // Denormalized:  <= mLen significand digits
				if (v * c >= 2)
				{
					e++;
					c /= 2;
				}                // Rounding can increment the exponent

				if (e + eBias >= eMax)
				{
					// Overflow
					m = 0;
					e = eMax;
				}
				else if (e + eBias >= 1)
				{
					// Normalized - term order matters, as Math.pow(2, 52-e) and v*Math.pow(2, 52) can overflow
					m = (v * c - 1) * Math.pow(2, mLen);
					e = e + eBias;
				}
				else
				{
					// Denormalized - also catches the '0' case, somewhat by chance
					m = v * Math.pow(2, eBias - 1) * Math.pow(2, mLen);
					e = 0;
				}
			}
			var a = new Array(8);
			for (i = 0, d = 1; mLen >= 8; a[i] = m & 0xff, i += d, m /= 256, mLen -= 8);
			for (e = (e << mLen) | m, eLen += mLen; eLen > 0; a[i] = e & 0xff, i += d, e /= 256, eLen -= 8);
			a[i - d] |= s * 128;
			return a;
		}
        this.WriteStringBySymbol = function(code)
        {
        	if (code < 0xFFFF)
			{
				this.CheckSize(4);
                this.data[this.pos++] = 1;
                this.data[this.pos++] = 0;
                this.data[this.pos++] = code & 0xFF;
                this.data[this.pos++] = (code >>> 8) & 0xFF;
			}
			else
			{
                this.CheckSize(6);
                this.data[this.pos++] = 2;
                this.data[this.pos++] = 0;

                var codePt = code - 0x10000;
                var c1 = 0xD800 | (codePt >> 10);
                var c2 = 0xDC00 | (codePt & 0x3FF);
                this.data[this.pos++] = c1 & 0xFF;
                this.data[this.pos++] = (c1 >>> 8) & 0xFF;
                this.data[this.pos++] = c2 & 0xFF;
                this.data[this.pos++] = (c2 >>> 8) & 0xFF;
			}
        }
		this.WriteString        = function(text)
		{
			if ("string" != typeof text)
				text = text + "";

			var count = text.length & 0xFFFF;
			this.CheckSize(2 * count + 2);
			this.data[this.pos++] = count & 0xFF;
			this.data[this.pos++] = (count >>> 8) & 0xFF;
			for (var i = 0; i < count; i++)
			{
				var c                 = text.charCodeAt(i) & 0xFFFF;
				this.data[this.pos++] = c & 0xFF;
				this.data[this.pos++] = (c >>> 8) & 0xFF;
			}
		}
		this.WriteString2       = function(text)
		{
			if ("string" != typeof text)
				text = text + "";

			var count      = text.length & 0x7FFFFFFF;
			var countWrite = 2 * count;
			this.WriteLong(countWrite);
			this.CheckSize(countWrite);
			for (var i = 0; i < count; i++)
			{
				var c                 = text.charCodeAt(i) & 0xFFFF;
				this.data[this.pos++] = c & 0xFF;
				this.data[this.pos++] = (c >>> 8) & 0xFF;
			}
		}
		this.WriteString3       = function(text)
		{
			if ("string" != typeof text)
				text = text + "";

			var count      = text.length & 0x7FFFFFFF;
			var countWrite = 2 * count;
			this.CheckSize(countWrite);
			for (var i = 0; i < count; i++)
			{
				var c                 = text.charCodeAt(i) & 0xFFFF;
				this.data[this.pos++] = c & 0xFF;
				this.data[this.pos++] = (c >>> 8) & 0xFF;
			}
		}
		this.WriteString4       = function(text)
		{
			if ("string" != typeof text)
				text = text + "";

			var count      = text.length & 0x7FFFFFFF;
			this.WriteLong(count);
			this.CheckSize(2 * count);
			for (var i = 0; i < count; i++)
			{
				var c                 = text.charCodeAt(i) & 0xFFFF;
				this.data[this.pos++] = c & 0xFF;
				this.data[this.pos++] = (c >>> 8) & 0xFF;
			}
		}
		this.WriteStringA = function(text)
		{
			var count = text.length & 0xFFFF;
			this.WriteULong(count);
			this.CheckSize(count);
			for (var i=0;i<count;i++)
			{
				var c = text.charCodeAt(i) & 0xFF;
				this.data[this.pos++] = c;
			}
		};
		this.ClearNoAttack      = function()
		{
			this.pos = 0;
		}

		this.WriteLongAt = function(_pos, val)
		{
			this.data[_pos++] = (val) & 0xFF;
			this.data[_pos++] = (val >>> 8) & 0xFF;
			this.data[_pos++] = (val >>> 16) & 0xFF;
			this.data[_pos++] = (val >>> 24) & 0xFF;
		}

		this.WriteBuffer = function(data, _pos, count)
		{
			this.CheckSize(count);
			for (var i = 0; i < count; i++)
			{
				this.data[this.pos++] = data[_pos + i];
			}
		}

		this.WriteUtf8Char = function(code)
		{
			this.CheckSize(1);
			if (code < 0x80) {
				this.data[this.pos++] = code;
			}
			else if (code < 0x0800) {
				this.data[this.pos++] = (0xC0 | (code >> 6));
				this.data[this.pos++] = (0x80 | (code & 0x3F));
			}
			else if (code < 0x10000) {
				this.data[this.pos++] = (0xE0 | (code >> 12));
				this.data[this.pos++] = (0x80 | ((code >> 6) & 0x3F));
				this.data[this.pos++] = (0x80 | (code & 0x3F));
			}
			else if (code < 0x1FFFFF) {
				this.data[this.pos++] = (0xF0 | (code >> 18));
				this.data[this.pos++] = (0x80 | ((code >> 12) & 0x3F));
				this.data[this.pos++] = (0x80 | ((code >> 6) & 0x3F));
				this.data[this.pos++] = (0x80 | (code & 0x3F));
			}
			else if (code < 0x3FFFFFF) {
				this.data[this.pos++] = (0xF8 | (code >> 24));
				this.data[this.pos++] = (0x80 | ((code >> 18) & 0x3F));
				this.data[this.pos++] = (0x80 | ((code >> 12) & 0x3F));
				this.data[this.pos++] = (0x80 | ((code >> 6) & 0x3F));
				this.data[this.pos++] = (0x80 | (code & 0x3F));
			}
			else if (code < 0x7FFFFFFF) {
				this.data[this.pos++] = (0xFC | (code >> 30));
				this.data[this.pos++] = (0x80 | ((code >> 24) & 0x3F));
				this.data[this.pos++] = (0x80 | ((code >> 18) & 0x3F));
				this.data[this.pos++] = (0x80 | ((code >> 12) & 0x3F));
				this.data[this.pos++] = (0x80 | ((code >> 6) & 0x3F));
				this.data[this.pos++] = (0x80 | (code & 0x3F));
			}
		};

		this.WriteXmlString = function(val)
		{
			var pCur = 0;
			var pEnd = val.length;
			while (pCur < pEnd)
			{
				var code = val.charCodeAt(pCur++);
				if (code >= 0xD800 && code <= 0xDFFF && pCur < pEnd)
				{
					code = 0x10000 + (((code & 0x3FF) << 10) | (0x03FF & val.charCodeAt(pCur++)));
				}
				this.WriteUtf8Char(code);
			}
		};
		this.WriteXmlStringEncode = function(val)
		{
			var pCur = 0;
			var pEnd = val.length;
			while (pCur < pEnd)
			{
				var code = val.charCodeAt(pCur++);
				if (code >= 0xD800 && code <= 0xDFFF && pCur < pEnd)
				{
					code = 0x10000 + (((code & 0x3FF) << 10) | (0x03FF & val.charCodeAt(pCur++)));
				}
				this.WriteXmlCharCode(code);
			}
		};
		this.WriteXmlCharCode = function(code)
		{
			switch (code)
			{
				case 0x26:
					//&
					this.WriteUtf8Char(0x26);
					this.WriteUtf8Char(0x61);
					this.WriteUtf8Char(0x6d);
					this.WriteUtf8Char(0x70);
					this.WriteUtf8Char(0x3b);
					break;
				case 0x27:
					//'
					this.WriteUtf8Char(0x26);
					this.WriteUtf8Char(0x61);
					this.WriteUtf8Char(0x70);
					this.WriteUtf8Char(0x6f);
					this.WriteUtf8Char(0x73);
					this.WriteUtf8Char(0x3b);
					break;
				case 0x3c:
					//<
					this.WriteUtf8Char(0x26);
					this.WriteUtf8Char(0x6c);
					this.WriteUtf8Char(0x74);
					this.WriteUtf8Char(0x3b);
					break;
				case 0x3e:
					//>
					this.WriteUtf8Char(0x26);
					this.WriteUtf8Char(0x67);
					this.WriteUtf8Char(0x74);
					this.WriteUtf8Char(0x3b);
					break;
				case 0x22:
					//"
					this.WriteUtf8Char(0x26);
					this.WriteUtf8Char(0x71);
					this.WriteUtf8Char(0x75);
					this.WriteUtf8Char(0x6f);
					this.WriteUtf8Char(0x74);
					this.WriteUtf8Char(0x3b);
					break;
				default:
					this.WriteUtf8Char(code);
					break;
			}
		};
		this.WriteXmlBool = function(val)
		{
			this.WriteXmlString(val ? '1' : '0');
		};
		this.WriteXmlByte = function(val)
		{
			this.WriteXmlInt(val);
		};
		this.WriteXmlSByte = function(val)
		{
			this.WriteXmlInt(val);
		};
		this.WriteXmlInt = function(val)
		{
			this.WriteXmlString(val.toFixed(0));
		};
		this.WriteXmlUInt = function(val)
		{
			this.WriteXmlInt(val);
		};
		this.WriteXmlInt64 = function(val)
		{
			this.WriteXmlInt(val);
		};
		this.WriteXmlUInt64 = function(val)
		{
			this.WriteXmlInt(val);
		};
		this.WriteXmlDouble = function(val)
		{
			this.WriteXmlNumber(val);
		};
		this.WriteXmlNumber = function(val)
		{
			this.WriteXmlString(val.toString());
		};
		this.WriteXmlNodeStart = function(name)
		{
			this.WriteUtf8Char(0x3c);
			this.WriteXmlString(name);
		};
		this.WriteXmlNodeEnd = function(name)
		{
			this.WriteUtf8Char(0x3c);
			this.WriteUtf8Char(0x2f);
			this.WriteXmlString(name);
			this.WriteUtf8Char(0x3e);
		};
		this.WriteXmlNodeWithText = function(name, text)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd(false);
			if (text)
				this.WriteXmlStringEncode(text.toString());
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlAttributesEnd = function(isEnd)
		{
			if (isEnd)
				this.WriteUtf8Char(0x2f);
			this.WriteUtf8Char(0x3e);
		};
		this.WriteXmlAttributeString = function(name, val)
		{
			this.WriteUtf8Char(0x20);
			this.WriteXmlString(name);
			this.WriteUtf8Char(0x3d);
			this.WriteUtf8Char(0x22);
			this.WriteXmlString(val.toString());
			this.WriteUtf8Char(0x22);
		};
		this.WriteXmlAttributeStringEncode = function(name, val)
		{
			this.WriteUtf8Char(0x20);
			this.WriteXmlString(name);
			this.WriteUtf8Char(0x3d);
			this.WriteUtf8Char(0x22);
			this.WriteXmlStringEncode(val.toString());
			this.WriteUtf8Char(0x22);
		};
		this.WriteXmlAttributeBool = function(name, val)
		{
			this.WriteXmlAttributeString(name, val ? '1' : '0');
		};
		this.WriteXmlAttributeByte = function(name, val)
		{
			this.WriteXmlAttributeInt(name, val);
		};
		this.WriteXmlAttributeSByte = function(name, val)
		{
			this.WriteXmlAttributeInt(name, val);
		};
		this.WriteXmlAttributeInt = function(name, val)
		{
			this.WriteXmlAttributeString(name, val.toFixed(0));
		};
		this.WriteXmlAttributeUInt = function(name, val)
		{
			this.WriteXmlAttributeInt(name, val);
		};
		this.WriteXmlAttributeInt64 = function(name, val)
		{
			this.WriteXmlAttributeInt(name, val);
		};
		this.WriteXmlAttributeUInt64 = function(name, val)
		{
			this.WriteXmlAttributeInt(name, val);
		};
		this.WriteXmlAttributeDouble = function(name, val)
		{
			this.WriteXmlAttributeNumber(name, val);
		};
		this.WriteXmlAttributeNumber = function(name, val)
		{
			this.WriteXmlAttributeString(name, val.toString());
		};
		this.WriteXmlNullable = function(val, name)
		{
			if (val) {
				val.toXml(this, name);
			}
		};
		//пересмотреть, куча аргументов
		this.WriteXmlArray = function(val, name, opt_parentName, needWriteCount, ns, childns)
		{
			if (!ns) {
				ns = "";
			}
			if (!childns) {
				childns = "";
			}
			if(val && val.length > 0) {
				if(opt_parentName) {
					this.WriteXmlNodeStart(ns + opt_parentName);
					if (needWriteCount) {
						this.WriteXmlNullableAttributeNumber("count", val.length);
					}
					this.WriteXmlAttributesEnd();
				}
				val.forEach(function(elem, index){
					elem.toXml(this, name, childns, childns, index);
				}, this);
				if(opt_parentName) {
					this.WriteXmlNodeEnd(ns + opt_parentName);
				}
			}
		};
		this.WriteXmlNullableAttributeString = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeString(name, val)
			}
		};
		this.WriteXmlNullableAttributeStringEncode = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeStringEncode(name, val)
			}
		};
		this.WriteXmlNonEmptyAttributeStringEncode = function(name, val)
		{
			if (val) {
				this.WriteXmlAttributeStringEncode(name, val)
			}
		};
		this.WriteXmlNullableAttributeBool = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeBool(name, val)
			}
		};
		this.WriteXmlNullableAttributeBool2 = function(name, val)
		{
			//добавлюя по аналогии с x2t
			if (null !== val && undefined !== val) {
				this.WriteXmlNullableAttributeString(name, val ? "1": "0")
			}
		};
		this.WriteXmlNullableAttributeByte = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeByte(name, val)
			}
		};
		this.WriteXmlNullableAttributeSByte = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeSByte(name, val)
			}
		};
		this.WriteXmlNullableAttributeInt = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeInt(name, val)
			}
		};
		this.WriteXmlNullableAttributeUInt = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeUInt(name, val)
			}
		};
		this.WriteXmlNullableAttributeInt64 = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeInt64(name, val)
			}
		};
		this.WriteXmlNullableAttributeUInt64 = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeUInt64(name, val)
			}
		};
		this.WriteXmlNullableAttributeDouble = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeDouble(name, val)
			}
		};
		this.WriteXmlNullableAttributeNumber = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeNumber(name, val)
			}
		};
		this.WriteXmlNullableAttributeIntWithKoef = function(name, val, koef)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeInt(name, val * koef)
			}
		};
		this.WriteXmlNullableAttributeUIntWithKoef = function(name, val, koef)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlAttributeUInt(name, val * koef)
			}
		};
		this.WriteXmlAttributeBoolIfTrue = function(name, val)
		{
			if (val) {
				this.WriteXmlAttributeBool(name, val)
			}
		};
		this.WriteXmlValueString = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlString(val.toString());
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueStringEncode = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributeString("xml:space", "preserve");
			this.WriteXmlAttributesEnd();
			this.WriteXmlStringEncode(val.toString());
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueStringEncode2 = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlStringEncode(val.toString());
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueBool = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlBool(val);
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueByte = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlByte(val);
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueSByte = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlSByte(val);
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueInt = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlInt(val);
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueUInt = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlUInt(val);
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueInt64 = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlInt64(val);
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueUInt64 = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlUInt64(val);
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueDouble = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlDouble(val);
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlValueNumber = function(name, val)
		{
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributesEnd();
			this.WriteXmlNumber(val);
			this.WriteXmlNodeEnd(name);
		};
		this.WriteXmlNullableValueString = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueString(name, val)
			}
		};
		this.WriteXmlNullableValueStringEncode = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueStringEncode(name, val)
			}
		};
		this.WriteXmlNullableValueStringEncode2 = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueStringEncode2(name, val)
			}
		};
		this.WriteXmlNullableValueBool = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueBool(name, val)
			}
		};
		this.WriteXmlNullableValueByte = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueByte(name, val)
			}
		};
		this.WriteXmlNullableValueSByte = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueSByte(name, val)
			}
		};
		this.WriteXmlNullableValueInt = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueInt(name, val)
			}
		};
		this.WriteXmlNullableValueUInt = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueUInt(name, val)
			}
		};
		this.WriteXmlNullableValueInt64 = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueInt64(name, val)
			}
		};
		this.WriteXmlNullableValueUInt64 = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueUInt64(name, val)
			}
		};
		this.WriteXmlNullableValueDouble = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueDouble(name, val)
			}
		};
		this.WriteXmlNullableValueNumber = function(name, val)
		{
			if (null !== val && undefined !== val) {
				this.WriteXmlValueNumber(name, val)
			}
		};
		this.XlsbStartRecord = function(type, len) {
			//Type
			if (type < 0x80) {
				this.WriteByte(type);
			}
			else {
				this.WriteByte((type & 0x7F) | 0x80);
				this.WriteByte(type >> 7);
			}
			//Len
			for (var i = 0; i < 4; ++i) {
				var part = len & 0x7F;
				len = len >> 7;
				if (len === 0) {
					this.WriteByte(part);
					break;
				}
				else {
					this.WriteByte(part | 0x80);
				}
			}
		};
		this.XlsbEndRecord = function() {
		};
		//все аргументы сохраняю как в x2t, ns - префикс пока не использую
		this.WritingValNode = function(ns, name, val) {
			this.WriteXmlNodeStart(name);
			this.WriteXmlAttributeString("val", val);
			this.WriteXmlAttributesEnd(true);
		};
		this.WritingValNodeEncodeXml = function(ns, name, val) {
			this.WriteXmlNodeStart(name);
			this.WriteXmlNullableAttributeStringEncode("val", val);
			this.WriteXmlAttributesEnd(true);
		};
		this.WritingValNodeIf = function(ns, name, cond, val) {
			this.WriteXmlNodeStart(name);
			if (cond) {
				this.WriteXmlAttributeString("val", val);
			}
			this.WriteXmlAttributesEnd(true);
		};
		this.WriteXmlHeader = function()
		{
			this.WriteXmlString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
		};
		this.WriteXmlRelationshipsNS = function()
		{
			this.WriteXmlAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		};
	}

	function CCommandsType()
	{
		this.ctPenXML              = 0;
		this.ctPenColor            = 1;
		this.ctPenAlpha            = 2;
		this.ctPenSize             = 3;
		this.ctPenDashStyle        = 4;
		this.ctPenLineStartCap     = 5;
		this.ctPenLineEndCap       = 6;
		this.ctPenLineJoin         = 7;
		this.ctPenDashPatern       = 8;
		this.ctPenDashPatternCount = 9;
		this.ctPenDashOffset       = 10;
		this.ctPenAlign            = 11;
		this.ctPenMiterLimit       = 12;

		// brush
		this.ctBrushXML             = 20;
		this.ctBrushType            = 21;
		this.ctBrushColor1          = 22;
		this.ctBrushColor2          = 23;
		this.ctBrushAlpha1          = 24;
		this.ctBrushAlpha2          = 25;
		this.ctBrushTexturePathOld  = 26;
		this.ctBrushTextureAlpha    = 27;
		this.ctBrushTextureMode     = 28;
		this.ctBrushRectable        = 29;
		this.ctBrushRectableEnabled = 30;
		this.ctBrushGradient        = 31;
		this.ctBrushTexturePath     = 32;

		// font
		this.ctFontXML       = 40;
		this.ctFontName      = 41;
		this.ctFontSize      = 42;
		this.ctFontStyle     = 43;
		this.ctFontPath      = 44;
		this.ctFontGID       = 45;
		this.ctFontCharSpace = 46;

		// shadow
		this.ctShadowXML       = 50;
		this.ctShadowVisible   = 51;
		this.ctShadowDistanceX = 52;
		this.ctShadowDistanceY = 53;
		this.ctShadowBlurSize  = 54;
		this.ctShadowColor     = 55;
		this.ctShadowAlpha     = 56;

		// edge
		this.ctEdgeXML      = 70;
		this.ctEdgeVisible  = 71;
		this.ctEdgeDistance = 72;
		this.ctEdgeColor    = 73;
		this.ctEdgeAlpha    = 74;

		// text
		this.ctDrawText        = 80;
		this.ctDrawTextEx      = 81;
		this.ctDrawTextCode    = 82;
		this.ctDrawTextCodeGid = 83;

		// pathcommands
		this.ctPathCommandMoveTo          = 91;
		this.ctPathCommandLineTo          = 92;
		this.ctPathCommandLinesTo         = 93;
		this.ctPathCommandCurveTo         = 94;
		this.ctPathCommandCurvesTo        = 95;
		this.ctPathCommandArcTo           = 96;
		this.ctPathCommandClose           = 97;
		this.ctPathCommandEnd             = 98;
		this.ctDrawPath                   = 99;
		this.ctPathCommandStart           = 100;
		this.ctPathCommandGetCurrentPoint = 101;
		this.ctPathCommandText            = 102;
		this.ctPathCommandTextEx          = 103;

		// image
		this.ctDrawImage         = 110;
		this.ctDrawImageFromFile = 111;

		this.ctSetParams = 120;

		this.ctBeginCommand = 121;
		this.ctEndCommand   = 122;

		this.ctSetTransform   = 130;
		this.ctResetTransform = 131;

		this.ctClipMode = 140;

		this.ctCommandLong1   = 150;
		this.ctCommandDouble1 = 151;
		this.ctCommandString1 = 152;
		this.ctCommandLong2   = 153;
		this.ctCommandDouble2 = 154;
		this.ctCommandString2 = 155;

		this.ctHyperlink		= 160;
		this.ctLink				= 161;
		this.ctFormField		= 162;
		this.ctDocInfo			= 163;
		this.ctAnnotField		= 164;
		this.ctAnnotFieldDelete	= 165;

		this.ctPageWidth  = 200;
		this.ctPageHeight = 201;

		this.ctPageStart		= 202;
		this.ctPageEnd			= 203;
		this.ctDocumentEdit		= 204;
		this.ctDocumentClose	= 205;
		this.ctPageEdit			= 206;

		this.ctError = 255;
	}

	var CommandType = new CCommandsType();

	var MetaBrushType = {
		Solid    : 0,
		Gradient : 1,
		Texture  : 2
	};

	// 0 - dash
	// 1 - dashDot
	// 2 - dot
	// 3 - lgDash
	// 4 - lgDashDot
	// 5 - lgDashDotDot
	// 6 - solid
	// 7 - sysDash
	// 8 - sysDashDot
	// 9 - sysDashDotDot
	// 10- sysDot

	var DashPatternPresets = [
		[4, 3],
		[4, 3, 1, 3],
		[1, 3],
		[8, 3],
		[8, 3, 1, 3],
		[8, 3, 1, 3, 1, 3],
		undefined,
		[3, 1],
		[3, 1, 1, 1],
		[3, 1, 1, 1, 1, 1],
		[1, 1]
	];

	function CMetafileFontPicker(manager)
	{
		this.Manager = manager; 	// в идеале - кэш измерятеля. тогда ни один шрифт не будет загружен заново
		if (!this.Manager)
		{
			this.Manager = new AscFonts.CFontManager();
			this.Manager.Initialize(false)
		}

		this.FontsInCache = {};
		this.LastPickFont = null;
		this.LastPickFontNameOrigin = "";
		this.LastPickFontName = "";
		this.Metafile = null; 												// класс, которому будет подменяться шрифт

		this.SetFont = function(setFont)
		{
			var name = setFont.FontFamily.Name;
			var size = setFont.FontSize;

            var style = 0;
            if (setFont.Italic == true)
                style += 2;
            if (setFont.Bold == true)
                style += 1;

            var name_check = name + "_" + style;
			if (this.FontsInCache[name_check])
			{
				this.LastPickFont = this.FontsInCache[name_check];
			}
			else
			{
				var font = g_fontApplication.GetFontFileWeb(name, style);
                var font_name_index = AscFonts.g_map_font_index[font.m_wsFontName];
                var fontId = AscFonts.g_font_infos[font_name_index].GetFontID(AscCommon.g_font_loader, style);
                var test_id = fontId.id + fontId.faceIndex + size;

                var cache = this.Manager.m_oFontsCache;
                this.LastPickFont = cache.Fonts[test_id];
                if (!this.LastPickFont)
                	this.LastPickFont = cache.Fonts[test_id + "nbold"];
                if (!this.LastPickFont)
                    this.LastPickFont = cache.Fonts[test_id + "nitalic"];
                if (!this.LastPickFont)
                    this.LastPickFont = cache.Fonts[test_id + "nboldnitalic"];

                if (!this.LastPickFont)
				{
					// такого при правильном кэше быть не должно
					if (window["NATIVE_EDITOR_ENJINE"] && fontId.file.Status != 0)
					{
						fontId.file.LoadFontNative();
					}
					this.LastPickFont = cache.LockFont(fontId.file.stream_index, fontId.id, fontId.faceIndex, size, "", this.Manager);
				}

				this.FontsInCache[name_check] = this.LastPickFont;
			}

            this.LastPickFontNameOrigin = name;
			this.LastPickFontName = name;
			this.Metafile.SetFont(setFont, true);
		};

		this.FillTextCode = function(glyph)
		{
			if (this.LastPickFont && this.LastPickFont.GetGIDByUnicode(glyph))
			{
				if (this.LastPickFontName != this.LastPickFontNameOrigin)
				{
                    this.LastPickFontName = this.LastPickFontNameOrigin;
					this.Metafile.SetFontName(this.LastPickFontName);
				}
			}
			else
			{
                var name = AscFonts.FontPickerByCharacter.getFontBySymbol(glyph);
                if (name != this.LastPickFontName)
				{
					this.LastPickFontName = name;
                    this.Metafile.SetFontName(this.LastPickFontName);
                }
			}
		};
	}

	function isCloudPrintingUrl()
	{
		if (window["AscDesktopEditor"])
		{
			if ((undefined !== window["AscDesktopEditor"]["CryptoMode"]) && (0 < window["AscDesktopEditor"]["CryptoMode"]))
				return false;

			if (window["AscDesktopEditor"]["IsLocalFile"] && window["AscDesktopEditor"]["IsFilePrinting"])
			{
				if (!window["AscDesktopEditor"]["IsLocalFile"]() && window["AscDesktopEditor"]["IsFilePrinting"]())
					return true;
			}
		}
		return false;
	}

	function getCloudPrintingUrl(url)
	{
		var urlLocal = AscCommon.g_oDocumentUrls.getImageLocal(url);
		if (urlLocal && urlLocal.endsWith(".svg"))
		{
			let localWithoutExt = urlLocal.slice(0, urlLocal.length - 3);
			let urlRes = AscCommon.g_oDocumentUrls.getImageUrl(localWithoutExt + "wmf");
			if (urlRes)
				return urlRes;
			urlRes = AscCommon.g_oDocumentUrls.getImageUrl(localWithoutExt + "emf");
			if (urlRes)
				return urlRes;
		}
		return url;
	}

	function CMetafile(width, height)
	{
		this.Width  = width;
		this.Height = height;

		this.m_oPen   = new CPen();
		this.m_oBrush = new CBrush();

		this.m_oFont =
		{
			Name     : "",
			FontSize : -1,
			Style    : -1
		};

		// чтобы выставилось в первый раз
		this.m_oPen.Color.R    = -1;
		this.m_oBrush.Color1.R = -1;
		this.m_oBrush.Color2.R = -1;

		this.m_oTransform    = new CMatrix();
		this.m_arrayCommands = [];

		this.Memory               = null;
		this.VectorMemoryForPrint = null;

		this.BrushType = MetaBrushType.Solid;

		// RFonts
		this.m_oTextPr  = null;
		this.m_oGrFonts = new CGrRFonts();

		// просто чтобы не создавать каждый раз
		this.m_oFontSlotFont    = new CFontSetup();
		this.LastFontOriginInfo = {Name : "", Replace : null};
		this.m_oFontTmp = { FontFamily : { Name : "arial" }, Bold : false, Italic : false };

		this.StartOffset = 0;

		this.m_bIsPenDash = false;

		this.FontPicker = null;
	}

	CMetafile.prototype =
	{
		// pen methods
		p_color  : function(r, g, b, a)
		{
			if (this.m_oPen.Color.R != r || this.m_oPen.Color.G != g || this.m_oPen.Color.B != b)
			{
				this.m_oPen.Color.R = r;
				this.m_oPen.Color.G = g;
				this.m_oPen.Color.B = b;

				var value = b << 16 | g << 8 | r;
				this.Memory.WriteByte(CommandType.ctPenColor);
				this.Memory.WriteLong(value);
			}
			if (this.m_oPen.Color.A != a)
			{
				this.m_oPen.Color.A = a;
				this.Memory.WriteByte(CommandType.ctPenAlpha);
				this.Memory.WriteByte(a);
			}
		},
		p_width  : function(w)
		{
			var val = w / 1000;
			if (this.m_oPen.Size != val)
			{
				this.m_oPen.Size = val;
				this.Memory.WriteByte(CommandType.ctPenSize);
				this.Memory.WriteDouble(val);
			}
		},
		p_dash : function(params)
		{
			var bIsDash = (params && (params.length > 0)) ? true : false;

			if (false == this.m_bIsPenDash && bIsDash == this.m_bIsPenDash)
				return;
			this.m_bIsPenDash = bIsDash;

			if (!this.m_bIsPenDash)
			{
				this.Memory.WriteByte(CommandType.ctPenDashStyle);
				this.Memory.WriteByte(0);
			}
			else
			{
				this.Memory.WriteByte(CommandType.ctPenDashStyle);
				this.Memory.WriteByte(5);

				this.Memory.WriteLong(params.length);
				for (var i = 0; i < params.length; i++)
				{
					this.Memory.WriteDouble(params[i]);
				}
			}
		},
		// brush methods
		b_color1 : function(r, g, b, a)
		{
			if (this.BrushType != MetaBrushType.Solid)
			{
				this.Memory.WriteByte(CommandType.ctBrushType);
				this.Memory.WriteLong(1000);
				this.BrushType = MetaBrushType.Solid;
			}

			if (this.m_oBrush.Color1.R != r || this.m_oBrush.Color1.G != g || this.m_oBrush.Color1.B != b)
			{
				this.m_oBrush.Color1.R = r;
				this.m_oBrush.Color1.G = g;
				this.m_oBrush.Color1.B = b;

				var value = b << 16 | g << 8 | r;
				this.Memory.WriteByte(CommandType.ctBrushColor1);
				this.Memory.WriteLong(value);
			}
			if (this.m_oBrush.Color1.A != a)
			{
				this.m_oBrush.Color1.A = a;
				this.Memory.WriteByte(CommandType.ctBrushAlpha1);
				this.Memory.WriteByte(a);
			}
		},
		b_color2 : function(r, g, b, a)
		{
			if (this.m_oBrush.Color2.R != r || this.m_oBrush.Color2.G != g || this.m_oBrush.Color2.B != b)
			{
				this.m_oBrush.Color2.R = r;
				this.m_oBrush.Color2.G = g;
				this.m_oBrush.Color2.B = b;

				var value = b << 16 | g << 8 | r;
				this.Memory.WriteByte(CommandType.ctBrushColor2);
				this.Memory.WriteLong(value);
			}
			if (this.m_oBrush.Color2.A != a)
			{
				this.m_oBrush.Color2.A = a;
				this.Memory.WriteByte(CommandType.ctBrushAlpha2);
				this.Memory.WriteByte(a);
			}
		},

		put_brushTexture : function(src, mode)
		{
			var isCloudPrinting = isCloudPrintingUrl();

			if (this.BrushType !== MetaBrushType.Texture)
			{
				this.Memory.WriteByte(CommandType.ctBrushType);
				this.Memory.WriteLong(3008);
				this.BrushType = MetaBrushType.Texture;
			}

			this.m_oBrush.Color1.R = -1;
			this.m_oBrush.Color1.G = -1;
			this.m_oBrush.Color1.B = -1;
			this.m_oBrush.Color1.A = -1;
			this.Memory.WriteByte(CommandType.ctBrushTexturePath);

			var _src = src;
			if (isCloudPrinting)
			{
				_src = getCloudPrintingUrl(src)
			}
			else
			{
				var srcLocal = AscCommon.g_oDocumentUrls.getLocal(_src);
				if (srcLocal)
					_src = srcLocal;
			}

			this.Memory.WriteString2(_src);
			this.Memory.WriteByte(CommandType.ctBrushTextureMode);
			this.Memory.WriteByte(mode);
		},

		put_BrushTextureAlpha : function(alpha)
		{
			var write = alpha;
			if (null == alpha || undefined == alpha)
				write = 255;

			this.Memory.WriteByte(CommandType.ctBrushTextureAlpha);
			this.Memory.WriteByte(write);
		},

		put_BrushGradient : function(gradFill, points, transparent)
		{
			this.BrushType = MetaBrushType.Gradient;

			this.Memory.WriteByte(CommandType.ctBrushGradient);

			this.Memory.WriteByte(AscCommon.g_nodeAttributeStart);

			if (gradFill.path != null && (gradFill.lin == null || gradFill.lin == undefined))
			{
				this.Memory.WriteByte(1);
				this.Memory.WriteByte(gradFill.path);

				this.Memory.WriteDouble(points.x0);
				this.Memory.WriteDouble(points.y0);
				this.Memory.WriteDouble(points.x1);
				this.Memory.WriteDouble(points.y1);
				this.Memory.WriteDouble(points.r0);
				this.Memory.WriteDouble(points.r1);
			}
			else
			{
				this.Memory.WriteByte(0);
				if (null == gradFill.lin)
				{
					this.Memory.WriteLong(90 * 60000);
					this.Memory.WriteBool(false);
				}
				else
				{
					this.Memory.WriteLong(gradFill.lin.angle);
					this.Memory.WriteBool(gradFill.lin.scale);
				}

				this.Memory.WriteDouble(points.x0);
				this.Memory.WriteDouble(points.y0);
				this.Memory.WriteDouble(points.x1);
				this.Memory.WriteDouble(points.y1);
			}

			var _colors = gradFill.colors;
			var firstColor = null;
			var lastColor = null;

			if (_colors.length > 0)
			{
				if (_colors[0].pos > 0)
				{
					firstColor = {
						color : {
							RGBA : {
								R : _colors[0].color.RGBA.R,
								G : _colors[0].color.RGBA.G,
								B : _colors[0].color.RGBA.B,
								A : _colors[0].color.RGBA.A
							}
						},
						pos : 0
					};
					_colors.unshift(firstColor);
				}

				var posLast = _colors.length - 1;
				if (_colors[posLast].pos < 100000)
				{
					lastColor = {
						color : {
							RGBA : {
								R : _colors[posLast].color.RGBA.R,
								G : _colors[posLast].color.RGBA.G,
								B : _colors[posLast].color.RGBA.B,
								A : _colors[posLast].color.RGBA.A
							}
						},
						pos : 100000
					};
					_colors.push(lastColor);
				}
			}

			this.Memory.WriteByte(2);
			this.Memory.WriteLong(_colors.length);

			for (var i = 0; i < _colors.length; i++)
			{
				this.Memory.WriteLong(_colors[i].pos);

				this.Memory.WriteByte(_colors[i].color.RGBA.R);
				this.Memory.WriteByte(_colors[i].color.RGBA.G);
				this.Memory.WriteByte(_colors[i].color.RGBA.B);

				if (null == transparent)
					this.Memory.WriteByte(_colors[i].color.RGBA.A);
				else
					this.Memory.WriteByte(transparent);
			}

			if (firstColor)
				_colors.shift();
			if (lastColor)
				_colors.pop();

			this.Memory.WriteByte(AscCommon.g_nodeAttributeEnd);
		},

		transform                 : function(sx, shy, shx, sy, tx, ty)
		{
			if (this.m_oTransform.sx != sx || this.m_oTransform.shx != shx || this.m_oTransform.shy != shy ||
				this.m_oTransform.sy != sy || this.m_oTransform.tx != tx || this.m_oTransform.ty != ty)
			{
				this.m_oTransform.sx  = sx;
				this.m_oTransform.shx = shx;
				this.m_oTransform.shy = shy;
				this.m_oTransform.sy  = sy;
				this.m_oTransform.tx  = tx;
				this.m_oTransform.ty  = ty;

				this.Memory.WriteByte(CommandType.ctSetTransform);
				this.Memory.WriteDouble(sx);
				this.Memory.WriteDouble(shy);
				this.Memory.WriteDouble(shx);
				this.Memory.WriteDouble(sy);
				this.Memory.WriteDouble(tx);
				this.Memory.WriteDouble(ty);
			}
		},
		// path commands
		_s                        : function()
		{
			if (this.VectorMemoryForPrint != null)
				this.VectorMemoryForPrint.ClearNoAttack();

			var _memory = (null == this.VectorMemoryForPrint) ? this.Memory : this.VectorMemoryForPrint;
			_memory.WriteByte(CommandType.ctPathCommandStart);
		},
		_e                        : function()
		{
			// тут всегда напрямую в Memory
			this.Memory.WriteByte(CommandType.ctPathCommandEnd);
		},
		_z                        : function()
		{
			var _memory = (null == this.VectorMemoryForPrint) ? this.Memory : this.VectorMemoryForPrint;
			_memory.WriteByte(CommandType.ctPathCommandClose);
		},
		_m                        : function(x, y)
		{
			var _memory = (null == this.VectorMemoryForPrint) ? this.Memory : this.VectorMemoryForPrint;
			_memory.WriteByte(CommandType.ctPathCommandMoveTo);
			_memory.WriteDouble(x);
			_memory.WriteDouble(y);
		},
		_l                        : function(x, y)
		{
			var _memory = (null == this.VectorMemoryForPrint) ? this.Memory : this.VectorMemoryForPrint;
			_memory.WriteByte(CommandType.ctPathCommandLineTo);
			_memory.WriteDouble(x);
			_memory.WriteDouble(y);
		},
		_c                        : function(x1, y1, x2, y2, x3, y3)
		{
			var _memory = (null == this.VectorMemoryForPrint) ? this.Memory : this.VectorMemoryForPrint;
			_memory.WriteByte(CommandType.ctPathCommandCurveTo);
			_memory.WriteDouble(x1);
			_memory.WriteDouble(y1);
			_memory.WriteDouble(x2);
			_memory.WriteDouble(y2);
			_memory.WriteDouble(x3);
			_memory.WriteDouble(y3);
		},
		_c2                       : function(x1, y1, x2, y2)
		{
			var _memory = (null == this.VectorMemoryForPrint) ? this.Memory : this.VectorMemoryForPrint;
			_memory.WriteByte(CommandType.ctPathCommandCurveTo);
			_memory.WriteDouble(x1);
			_memory.WriteDouble(y1);
			_memory.WriteDouble(x1);
			_memory.WriteDouble(y1);
			_memory.WriteDouble(x2);
			_memory.WriteDouble(y2);
		},
		ds                        : function()
		{
			if (null == this.VectorMemoryForPrint)
			{
				this.Memory.WriteByte(CommandType.ctDrawPath);
				this.Memory.WriteLong(1);
			}
			else
			{
				this.Memory.Copy(this.VectorMemoryForPrint, 0, this.VectorMemoryForPrint.pos);
				this.Memory.WriteByte(CommandType.ctDrawPath);
				this.Memory.WriteLong(1);
			}
		},
		df                        : function()
		{
			if (null == this.VectorMemoryForPrint)
			{
				this.Memory.WriteByte(CommandType.ctDrawPath);
				this.Memory.WriteLong(256);
			}
			else
			{
				this.Memory.Copy(this.VectorMemoryForPrint, 0, this.VectorMemoryForPrint.pos);
				this.Memory.WriteByte(CommandType.ctDrawPath);
				this.Memory.WriteLong(256);
			}
		},
		WriteVectorMemoryForPrint : function()
		{
			if (null != this.VectorMemoryForPrint)
			{
				this.Memory.Copy(this.VectorMemoryForPrint, 0, this.VectorMemoryForPrint.pos);
			}
		},
		drawpath                  : function(type)
		{
			if (null == this.VectorMemoryForPrint)
			{
				this.Memory.WriteByte(CommandType.ctDrawPath);
				this.Memory.WriteLong(type);
			}
			else
			{
				this.Memory.Copy(this.VectorMemoryForPrint, 0, this.VectorMemoryForPrint.pos);
				this.Memory.WriteByte(CommandType.ctDrawPath);
				this.Memory.WriteLong(type);
			}
		},

		// canvas state
		save    : function()
		{
		},
		restore : function()
		{
		},
		clip    : function()
		{
		},

		// images
		drawImage : function(img, x, y, w, h, isUseOriginUrl)
		{
			var isCloudPrinting = isCloudPrintingUrl();

			if (!window.editor)
			{
				// excel
				this.Memory.WriteByte(CommandType.ctDrawImageFromFile);

				let _img = img;
				if (isCloudPrinting)
				{
					_img = getCloudPrintingUrl(_img);
				}
				else
				{
					var imgLocal = AscCommon.g_oDocumentUrls.getLocal(img);
					if (imgLocal && (true !== isUseOriginUrl))
					{
						_img = imgLocal;
					}
				}

				this.Memory.WriteString2(_img);
				this.Memory.WriteDouble(x);
				this.Memory.WriteDouble(y);
				this.Memory.WriteDouble(w);
				this.Memory.WriteDouble(h);
				return;
			}

			var _src = "";
			if (!window["NATIVE_EDITOR_ENJINE"] && (true !== isUseOriginUrl))
			{
				var _img = window.editor.ImageLoader.map_image_index[img];
				if (_img == undefined || _img.Image == null)
					return;

				_src = _img.src;
			}
			else
			{
				_src = img;
			}

			if (isCloudPrinting)
			{
				_src = getCloudPrintingUrl(_src)
			}
			else
			{
				var srcLocal = AscCommon.g_oDocumentUrls.getLocal(_src);
				if (srcLocal)
					_src = srcLocal;
			}

			this.Memory.WriteByte(CommandType.ctDrawImageFromFile);
			this.Memory.WriteString2(_src);
			this.Memory.WriteDouble(x);
			this.Memory.WriteDouble(y);
			this.Memory.WriteDouble(w);
			this.Memory.WriteDouble(h);
		},

		SetFontName : function(name)
		{
            var fontinfo = g_fontApplication.GetFontInfo(name, 0, this.LastFontOriginInfo);
            if (this.m_oFont.Name != fontinfo.Name)
            {
                this.m_oFont.Name = fontinfo.Name;
                this.Memory.WriteByte(CommandType.ctFontName);
                this.Memory.WriteString(this.m_oFont.Name);
            }
		},

		SetFont      : function(font, isFromPicker)
		{
			if (this.FontPicker && !isFromPicker)
				return this.FontPicker.SetFont(font);

			if (null == font)
				return;

			var style = 0;
			if (font.Italic == true)
				style += 2;
			if (font.Bold == true)
				style += 1;

			var fontinfo = g_fontApplication.GetFontInfo(font.FontFamily.Name, style, this.LastFontOriginInfo);
			//style        = fontinfo.GetBaseStyle(style);

			if (this.m_oFont.Name != fontinfo.Name)
			{
				this.m_oFont.Name = fontinfo.Name;
				this.Memory.WriteByte(CommandType.ctFontName);
				this.Memory.WriteString(this.m_oFont.Name);
			}
			if (this.m_oFont.FontSize != font.FontSize)
			{
				this.m_oFont.FontSize = font.FontSize;
				this.Memory.WriteByte(CommandType.ctFontSize);
				this.Memory.WriteDouble(this.m_oFont.FontSize);
			}
			if (this.m_oFont.Style != style)
			{
				this.m_oFont.Style = style;
				this.Memory.WriteByte(CommandType.ctFontStyle);
				this.Memory.WriteLong(style);
			}
		},
		FillText     : function(x, y, text)
		{
			if (1 == text.length)
				return this.FillTextCode(x, y, text.charCodeAt(0));

			this.Memory.WriteByte(CommandType.ctDrawText);

			this.Memory.WriteString(text);
			this.Memory.WriteDouble(x);
			this.Memory.WriteDouble(y);
		},
		FillTextCode : function(x, y, code)
		{
			var _code = code;
            if (null != this.LastFontOriginInfo.Replace)
				_code = g_fontApplication.GetReplaceGlyph(_code, this.LastFontOriginInfo.Replace);

			if (this.FontPicker)
				this.FontPicker.FillTextCode(_code);

            this.Memory.WriteByte(CommandType.ctDrawText);
			this.Memory.WriteStringBySymbol(_code);
            this.Memory.WriteDouble(x);
            this.Memory.WriteDouble(y);
		},
		tg           : function(gid, x, y, codepoints)
		{
			/*
			var _old_pos = this.Memory.pos;
			g_fontApplication.LoadFont(this.m_oFont.Name, AscCommon.g_font_loader, AscCommon.g_oTextMeasurer.m_oManager, this.m_oFont.FontSize, Math.max(this.m_oFont.Style, 0), 72, 72);
			AscCommon.g_oTextMeasurer.m_oManager.LoadStringPathCode(gid, true, x, y, this);
			// start (1) + draw(1) + typedraw(4) + end(1) = 7!
			if ((this.Memory.pos - _old_pos) < 8)
				this.Memory.pos = _old_pos;
			*/

			this.Memory.WriteByte(CommandType.ctDrawTextCodeGid);
			this.Memory.WriteLong(gid);
			this.Memory.WriteDouble(x);
			this.Memory.WriteDouble(y);
			var count = codepoints ? codepoints.length : 0;
			this.Memory.WriteLong(count);
			for (var i = 0; i < count; i++)
				this.Memory.WriteLong(codepoints[i]);
		},
		charspace    : function(space)
		{
		},

		beginCommand : function(command)
		{
			this.Memory.WriteByte(CommandType.ctBeginCommand);
			this.Memory.WriteLong(command);
		},
		endCommand   : function(command)
		{
			if (32 == command)
			{
				if (null == this.VectorMemoryForPrint)
				{
					this.Memory.WriteByte(CommandType.ctEndCommand);
					this.Memory.WriteLong(command);
				}
				else
				{
					this.Memory.Copy(this.VectorMemoryForPrint, 0, this.VectorMemoryForPrint.pos);
					this.Memory.WriteByte(CommandType.ctEndCommand);
					this.Memory.WriteLong(command);
				}
				return;
			}
			this.Memory.WriteByte(CommandType.ctEndCommand);
			this.Memory.WriteLong(command);
		},

		put_PenLineJoin          : function(_join)
		{
			this.Memory.WriteByte(CommandType.ctPenLineJoin);
			this.Memory.WriteByte(_join & 0xFF);
		},
		put_TextureBounds        : function(x, y, w, h)
		{
			this.Memory.WriteByte(CommandType.ctBrushRectable);
			this.Memory.WriteDouble(x);
			this.Memory.WriteDouble(y);
			this.Memory.WriteDouble(w);
			this.Memory.WriteDouble(h);
		},
		put_TextureBoundsEnabled : function(bIsEnabled)
		{
			this.Memory.WriteByte(CommandType.ctBrushRectableEnabled);
			this.Memory.WriteBool(bIsEnabled);
		},

		SetFontInternal : function(name, size, style)
		{
			// TODO: remove m_oFontSlotFont
			var _lastFont = this.m_oFontSlotFont;
			_lastFont.Name = name;
			_lastFont.Size = size;
			_lastFont.Bold = (style & AscFonts.FontStyle.FontStyleBold) ? true : false;
			_lastFont.Italic = (style & AscFonts.FontStyle.FontStyleItalic) ? true : false;

			this.m_oFontTmp.FontFamily.Name = _lastFont.Name;
			this.m_oFontTmp.Bold = _lastFont.Bold;
			this.m_oFontTmp.Italic = _lastFont.Italic;
			this.m_oFontTmp.FontSize = _lastFont.Size;
			this.SetFont(this.m_oFontTmp);
		},

		SetFontSlot : function(slot, fontSizeKoef)
		{
			var _rfonts   = this.m_oGrFonts;
			var _lastFont = this.m_oFontSlotFont;

			switch (slot)
			{
				case fontslot_ASCII:
				{
					_lastFont.Name   = _rfonts.Ascii.Name;
					_lastFont.Size   = this.m_oTextPr.FontSize;
					_lastFont.Bold   = this.m_oTextPr.Bold;
					_lastFont.Italic = this.m_oTextPr.Italic;

					break;
				}
				case fontslot_CS:
				{
					_lastFont.Name   = _rfonts.CS.Name;
					_lastFont.Size   = this.m_oTextPr.FontSizeCS;
					_lastFont.Bold   = this.m_oTextPr.BoldCS;
					_lastFont.Italic = this.m_oTextPr.ItalicCS;

					break;
				}
				case fontslot_EastAsia:
				{
					_lastFont.Name   = _rfonts.EastAsia.Name;
					_lastFont.Size   = this.m_oTextPr.FontSize;
					_lastFont.Bold   = this.m_oTextPr.Bold;
					_lastFont.Italic = this.m_oTextPr.Italic;

					break;
				}
				case fontslot_HAnsi:
				default:
				{
					_lastFont.Name   = _rfonts.HAnsi.Name;
					_lastFont.Size   = this.m_oTextPr.FontSize;
					_lastFont.Bold   = this.m_oTextPr.Bold;
					_lastFont.Italic = this.m_oTextPr.Italic;

					break;
				}
			}

			if (undefined !== fontSizeKoef)
				_lastFont.Size *= fontSizeKoef;

            this.m_oFontTmp.FontFamily.Name = _lastFont.Name;
            this.m_oFontTmp.Bold = _lastFont.Bold;
            this.m_oFontTmp.Italic = _lastFont.Italic;
            this.m_oFontTmp.FontSize = _lastFont.Size;
            this.SetFont(this.m_oFontTmp);
		},

		AddHyperlink : function(x, y, w, h, url, tooltip)
		{
			this.Memory.WriteByte(CommandType.ctHyperlink);
			this.Memory.WriteDouble(x);
			this.Memory.WriteDouble(y);
			this.Memory.WriteDouble(w);
			this.Memory.WriteDouble(h);
			this.Memory.WriteString(url);
			this.Memory.WriteString(tooltip);
		},

		AddLink : function(x, y, w, h, dx, dy, dPage)
		{
			this.Memory.WriteByte(CommandType.ctLink);
			this.Memory.WriteDouble(x);
			this.Memory.WriteDouble(y);
			this.Memory.WriteDouble(w);
			this.Memory.WriteDouble(h);
			this.Memory.WriteDouble(dx);
			this.Memory.WriteDouble(dy);
			this.Memory.WriteLong(dPage);
		},

		AddFormField : function(nX, nY, nW, nH, nBaseLineOffset, oForm)
		{
			if (!oForm)
				return;

			this.Memory.WriteByte(CommandType.ctFormField);

			var nStartPos = this.Memory.GetCurPosition();
			this.Memory.Skip(4);

			this.Memory.WriteDouble(nX);
			this.Memory.WriteDouble(nY);
			this.Memory.WriteDouble(nW);
			this.Memory.WriteDouble(nH);
			this.Memory.WriteDouble(nBaseLineOffset);

			var nFlagPos = this.Memory.GetCurPosition();
			this.Memory.Skip(4);
			var nFlag = 0;

			var oFormPr = oForm.GetFormPr();
			
			let formKey = null;
			if (!oForm.IsMainForm())
			{
				let mainForm = oForm.GetMainForm();
				let subIndex = oForm.GetSubFormIndex();
				formKey = mainForm.GetFormKey() + "_" + subIndex;
			}
			else
			{
				formKey = oFormPr.GetKey();
			}
			
			if (formKey)
			{
				nFlag |= 1;
				this.Memory.WriteString(formKey);
			}

			var sHelpText = oFormPr.GetHelpText();
			if (sHelpText)
			{
				nFlag |= (1 << 1);
				this.Memory.WriteString(sHelpText);
			}

			if (oFormPr.GetRequired())
				nFlag |= (1 << 2);

			if (oForm.IsPlaceHolder())
				nFlag |= (1 << 3);

			// 7-ой и 8-ой биты зарезервированы для бордера
			var oBorder = oFormPr.GetBorder();
			if (oBorder && !oBorder.IsNone())
			{
				nFlag |= (1 << 6);

				var oColor = oBorder.GetColor();
				this.Memory.WriteLong(1);
				this.Memory.WriteDouble(oBorder.GetWidth());
				this.Memory.WriteByte(oColor.r);
				this.Memory.WriteByte(oColor.g);
				this.Memory.WriteByte(oColor.b);
				this.Memory.WriteByte(0x255);
			}

			var oParagraph = oForm.GetParagraph();

			var oShd = oFormPr.GetShd();
			if (oParagraph && oShd && !oShd.IsNil())
			{
				nFlag |= (1 << 9);

				var oColor = oShd.GetSimpleColor(oParagraph.GetTheme(), oParagraph.GetColorMap());
				this.Memory.WriteByte(oColor.r);
				this.Memory.WriteByte(oColor.g);
				this.Memory.WriteByte(oColor.b);
				this.Memory.WriteByte(0x255);
			}

			if (oParagraph && AscCommon.align_Left !== oParagraph.GetParagraphAlign())
			{
				nFlag |= (1 << 10);
				this.Memory.WriteByte(oParagraph.GetParagraphAlign());
			}

			// 0 - Unknown
			// 1 - Text
			// 2 - ComboBox/DropDownList
			// 3 - CheckBox/RadioButton
			// 4 - Picture
			// 5 - Signature
			// 6 - DateTime

			if (oForm.IsTextForm())
			{
				this.Memory.WriteLong(1);
				var oTextFormPr = oForm.GetTextFormPr();

				if (oTextFormPr.Comb)
					nFlag |= (1 << 20);

				if (oTextFormPr.MaxCharacters > 0)
				{
					nFlag |= (1 << 21);
					this.Memory.WriteLong(oTextFormPr.MaxCharacters);
				}

				let sValue = oForm.GetSelectedText(true, false, {NewLine : true});
				if (sValue)
				{
					nFlag |= (1 << 22);
					this.Memory.WriteString(sValue);
				}

				if (oTextFormPr.MultiLine && oForm.IsFixedForm())
					nFlag |= (1 << 23);

				if (oTextFormPr.AutoFit)
					nFlag |= (1 << 24);

				let sPlaceHolderText = oForm.GetPlaceholderText();
				if (sPlaceHolderText)
				{
					nFlag |= (1 << 25);
					this.Memory.WriteString(sPlaceHolderText);
				}

				let format = oTextFormPr.GetFormat();
				if (!format.IsEmpty())
				{
					nFlag |= (1 << 26);

					this.Memory.WriteByte(format.GetType());

					let formatSymbols = format.GetSymbols(false);

					this.Memory.WriteLong(formatSymbols.length);
					for (let index = 0, count = formatSymbols.length; index < count; ++index)
					{
						this.Memory.WriteLong(formatSymbols[index]);
					}

					let mask = "";

					if (format.IsMask())
						mask = format.GetMask();
					else if (format.IsRegExp())
						mask = format.GetRegExp();

					this.Memory.WriteString(mask);
				}
			}
			else if (oForm.IsComboBox() || oForm.IsDropDownList())
			{
				this.Memory.WriteLong(2);
				var isComboBox = oForm.IsComboBox();

				var oFormPr = isComboBox ? oForm.GetComboBoxPr() : oForm.GetDropDownListPr();

				if (isComboBox)
					nFlag |= (1 << 20);

				var sValue         = oForm.GetSelectedText(true);
				var nSelectedIndex = -1;

				// Обработка "Choose an item"
				var nItemsCount = oFormPr.GetItemsCount();
				if (nItemsCount > 0 && AscCommon.translateManager.getValue("Choose an item") === oFormPr.GetItemDisplayText(0))
				{
					this.Memory.WriteLong(nItemsCount - 1);
					for (var nIndex = 1; nIndex < nItemsCount; ++nIndex)
					{
						var sItemValue = oFormPr.GetItemDisplayText(nIndex);
						if (sItemValue === sValue)
							nSelectedIndex = nIndex;

						this.Memory.WriteString(sItemValue);
					}
				}
				else
				{
					this.Memory.WriteLong(nItemsCount);
					for (var nIndex = 0; nIndex < nItemsCount; ++nIndex)
					{
						var sItemValue = oFormPr.GetItemDisplayText(nIndex);
						if (sItemValue === sValue)
							nSelectedIndex = nIndex;

						this.Memory.WriteString(sItemValue);
					}
				}

				this.Memory.WriteLong(nSelectedIndex);

				if (sValue)
				{
					nFlag |= (1 << 22);
					this.Memory.WriteString(sValue);
				}

				var sPlaceHolderText = oForm.GetPlaceholderText();
				if (sPlaceHolderText)
				{
					nFlag |= (1 << 23);
					this.Memory.WriteString(sPlaceHolderText);
				}
			}
			else if (oForm.IsCheckBox())
			{
				this.Memory.WriteLong(3);

				var oCheckBoxPr = oForm.GetCheckBoxPr();

				if (oCheckBoxPr.GetChecked())
					nFlag |= (1 << 20);

				var nCheckedSymbol   = oCheckBoxPr.GetCheckedSymbol();
				var nUncheckedSymbol = oCheckBoxPr.GetUncheckedSymbol();

				var nType = 0x0000;
				if (0x2611 === nCheckedSymbol && 0x2610 === nUncheckedSymbol)
					nType = 0x0001;
				else if (0x25C9 === nCheckedSymbol && 0x25CB === nUncheckedSymbol)
					nType = 0x0002;

				var sCheckedFont = oCheckBoxPr.GetCheckedFont();
				if (AscCommon.IsAscFontSupport(sCheckedFont, nCheckedSymbol))
					sCheckedFont = "ASCW3";

				var sUncheckedFont = oCheckBoxPr.GetUncheckedFont();
				if (AscCommon.IsAscFontSupport(sUncheckedFont, nUncheckedSymbol))
					sUncheckedFont = "ASCW3";

				this.Memory.WriteLong(nType);
				this.Memory.WriteLong(nCheckedSymbol);
				this.Memory.WriteString(sCheckedFont);
				this.Memory.WriteLong(nUncheckedSymbol);
				this.Memory.WriteString(sUncheckedFont);

				var sGroupName = oCheckBoxPr.GetGroupKey();
				if (sGroupName)
				{
					nFlag |= (1 << 21);
					this.Memory.WriteString(sGroupName);
				}
			}
			else if (oForm.IsPicture())
			{
				this.Memory.WriteLong(4);

				var oPicturePr = oForm.GetPictureFormPr();

				if (oPicturePr.IsConstantProportions())
					nFlag |= (1 << 20);

				if (oPicturePr.IsRespectBorders())
					nFlag |= (1 << 21);

				nFlag |= ((oPicturePr.GetScaleFlag() & 0xF) << 24);
				this.Memory.WriteLong(oPicturePr.GetShiftX() * 1000);
				this.Memory.WriteLong(oPicturePr.GetShiftY() * 1000);

				if (!oForm.IsPlaceHolder())
				{
					var arrDrawings = oForm.GetAllDrawingObjects();
					if (arrDrawings.length > 0 && arrDrawings[0].IsPicture() && arrDrawings[0].GraphicObj.blipFill)
					{
						var _src = AscCommon.getFullImageSrc2(arrDrawings[0].GraphicObj.blipFill.RasterImageId);
						var isCloudPrinting = isCloudPrintingUrl();

						if (isCloudPrinting)
						{
							_src = getCloudPrintingUrl(_src)
						}
						else
						{
							var srcLocal = AscCommon.g_oDocumentUrls.getLocal(_src);
							if (srcLocal)
								_src = srcLocal;
						}

						nFlag |= (1 << 22);
						this.Memory.WriteString(_src);
					}
				}
			}
			else if (oForm.IsDatePicker())
			{
				this.Memory.WriteLong(6);
				let dateTimePr = oForm.GetDatePickerPr();
				
				let value = oForm.GetSelectedText(true, false, {NewLine : true});
				if (value)
				{
					nFlag |= (1 << 22);
					this.Memory.WriteString(value);
				}
				
				let placeholderText = oForm.GetPlaceholderText();
				if (placeholderText)
				{
					nFlag |= (1 << 25);
					this.Memory.WriteString(placeholderText);
				}

				let dateFormat = dateTimePr.GetDateFormat();
				if (dateFormat)
				{
					nFlag |= (1 << 26);
					this.Memory.WriteString(dateFormat);
				}
			}
			else
			{
				this.Memory.WriteLong(0);
			}

			var nEndPos = this.Memory.GetCurPosition();
			this.Memory.Seek(nFlagPos);
			this.Memory.WriteLong(nFlag);

			this.Memory.Seek(nStartPos);
			this.Memory.WriteLong(nEndPos - nStartPos);
			this.Memory.Seek(nEndPos);
		}
	};

	function CDocumentRenderer()
	{
		this.m_arrayPages         = [];
		this.m_lPagesCount        = 0;
		//this.DocumentInfo = "";
		this.Memory               = new CMemory();
		this.VectorMemoryForPrint = null;

		this.ClipManager            = new CClipManager();
		this.ClipManager.BaseObject = this;

		this.RENDERER_PDF_FLAG = true;
		this.ArrayPoints       = null;

		this.GrState        = new CGrState();
		this.GrState.Parent = this;

		this.m_oPen       = null;
		this.m_oBrush     = null;
		this.m_oTransform = null;

		this._restoreDumpedVectors = null;

		this.m_oBaseTransform = null;

		this.UseOriginImageUrl = false;

        this.FontPicker = null;

        this.isPrintMode = false;
	}

	CDocumentRenderer.prototype =
	{
		InitPicker : function(_manager)
		{
			this.FontPicker = new CMetafileFontPicker(_manager);
		},

		SetBaseTransform : function(_matrix)
		{
			this.m_oBaseTransform = _matrix;
		},
		BeginPage        : function(width, height)
		{
			this.m_arrayPages[this.m_arrayPages.length]                    = new CMetafile(width, height);
			this.m_lPagesCount                                             = this.m_arrayPages.length;
			this.m_arrayPages[this.m_lPagesCount - 1].Memory               = this.Memory;
			this.m_arrayPages[this.m_lPagesCount - 1].StartOffset          = this.Memory.pos;
			this.m_arrayPages[this.m_lPagesCount - 1].VectorMemoryForPrint = this.VectorMemoryForPrint;
            this.m_arrayPages[this.m_lPagesCount - 1].FontPicker		   = this.FontPicker;

            if (this.FontPicker)
            	this.m_arrayPages[this.m_lPagesCount - 1].FontPicker.Metafile  = this.m_arrayPages[this.m_lPagesCount - 1];

			this.Memory.WriteByte(CommandType.ctPageStart);

			this.Memory.WriteByte(CommandType.ctPageWidth);
			this.Memory.WriteDouble(width);
			this.Memory.WriteByte(CommandType.ctPageHeight);
			this.Memory.WriteDouble(height);

			var _page         = this.m_arrayPages[this.m_lPagesCount - 1];
			this.m_oPen       = _page.m_oPen;
			this.m_oBrush     = _page.m_oBrush;
			this.m_oTransform = _page.m_oTransform;
		},
		EndPage          : function()
		{
			this.Memory.WriteByte(CommandType.ctPageEnd);
		},

		p_color    : function(r, g, b, a)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].p_color(r, g, b, a);
		},
		p_width    : function(w)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].p_width(w);
		},
		p_dash : function(params)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].p_dash(params);
		},
		// brush methods
		b_color1   : function(r, g, b, a)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].b_color1(r, g, b, a);
		},
		b_color2   : function(r, g, b, a)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].b_color2(r, g, b, a);
		},
		transform  : function(sx, shy, shx, sy, tx, ty)
		{
			if (0 != this.m_lPagesCount)
			{
				if (null == this.m_oBaseTransform)
					this.m_arrayPages[this.m_lPagesCount - 1].transform(sx, shy, shx, sy, tx, ty);
				else
				{
					var _transform = new CMatrix();
					_transform.sx  = sx;
					_transform.shy = shy;
					_transform.shx = shx;
					_transform.sy  = sy;
					_transform.tx  = tx;
					_transform.ty  = ty;
					AscCommon.global_MatrixTransformer.MultiplyAppend(_transform, this.m_oBaseTransform);
					this.m_arrayPages[this.m_lPagesCount - 1].transform(_transform.sx, _transform.shy, _transform.shx, _transform.sy, _transform.tx, _transform.ty);
				}
			}
		},
		transform3 : function(m)
		{
			this.transform(m.sx, m.shy, m.shx, m.sy, m.tx, m.ty);
		},
		reset      : function()
		{
			this.transform(1, 0, 0, 1, 0, 0);
		},
		// path commands
		_s         : function()
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1]._s();
		},
		_e         : function()
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1]._e();
		},
		_z         : function()
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1]._z();
		},
		_m         : function(x, y)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1]._m(x, y);

			if (this.ArrayPoints != null)
				this.ArrayPoints[this.ArrayPoints.length] = {x : x, y : y};
		},
		_l         : function(x, y)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1]._l(x, y);

			if (this.ArrayPoints != null)
				this.ArrayPoints[this.ArrayPoints.length] = {x : x, y : y};
		},
		_c         : function(x1, y1, x2, y2, x3, y3)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1]._c(x1, y1, x2, y2, x3, y3);

			if (this.ArrayPoints != null)
			{
				this.ArrayPoints[this.ArrayPoints.length] = {x : x1, y : y1};
				this.ArrayPoints[this.ArrayPoints.length] = {x : x2, y : y2};
				this.ArrayPoints[this.ArrayPoints.length] = {x : x3, y : y3};
			}
		},
		_c2        : function(x1, y1, x2, y2)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1]._c2(x1, y1, x2, y2);

			if (this.ArrayPoints != null)
			{
				this.ArrayPoints[this.ArrayPoints.length] = {x : x1, y : y1};
				this.ArrayPoints[this.ArrayPoints.length] = {x : x2, y : y2};
			}
		},
		ds         : function()
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].ds();
		},
		df         : function()
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].df();
		},
		drawpath   : function(type)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].drawpath(type);
		},

		// canvas state
		save    : function()
		{
		},
		restore : function()
		{
		},
		clip    : function()
		{
		},

		// images
		drawImage : function(img, x, y, w, h, alpha, srcRect)
		{
			if (img == null || img == undefined || img == "")
				return;

			if (0 != this.m_lPagesCount)
			{
				if (!srcRect)
					this.m_arrayPages[this.m_lPagesCount - 1].drawImage(img, x, y, w, h, this.UseOriginImageUrl);
				else
				{
					/*
					 if (!window.editor)
					 {
					 this.m_arrayPages[this.m_lPagesCount - 1].drawImage(img,x,y,w,h);
					 return;
					 }
					 */

					/*
					var _img = undefined;
					if (window.editor)
						_img = window.editor.ImageLoader.map_image_index[img];
					else if (window["Asc"]["editor"])
						_img = window["Asc"]["editor"].ImageLoader.map_image_index[img];

					var w0 = 0;
					var h0 = 0;
					if (_img != undefined && _img.Image != null)
					{
						w0 = _img.Image.width;
						h0 = _img.Image.height;
					}

					if (w0 == 0 || h0 == 0)
					{
						this.m_arrayPages[this.m_lPagesCount - 1].drawImage(img, x, y, w, h);
						return;
					}
					*/

					var bIsClip = false;
					if (srcRect.l > 0 || srcRect.t > 0 || srcRect.r < 100 || srcRect.b < 100)
						bIsClip = true;

					if (bIsClip)
					{
						this.SaveGrState();
						this.AddClipRect(x, y, w, h);
					}

					var __w   = w;
					var __h   = h;
					var _delW = Math.max(0, -srcRect.l) + Math.max(0, srcRect.r - 100) + 100;
					var _delH = Math.max(0, -srcRect.t) + Math.max(0, srcRect.b - 100) + 100;

					if (srcRect.l < 0)
					{
						var _off = ((-srcRect.l / _delW) * __w);
						x += _off;
						w -= _off;
					}
					if (srcRect.t < 0)
					{
						var _off = ((-srcRect.t / _delH) * __h);
						y += _off;
						h -= _off;
					}
					if (srcRect.r > 100)
					{
						var _off = ((srcRect.r - 100) / _delW) * __w;
						w -= _off;
					}
					if (srcRect.b > 100)
					{
						var _off = ((srcRect.b - 100) / _delH) * __h;
						h -= _off;
					}

					var _wk = 100;
					if (srcRect.l > 0)
						_wk -= srcRect.l;
					if (srcRect.r < 100)
						_wk -= (100 - srcRect.r);
					_wk = 100 / _wk;

					var _hk = 100;
					if (srcRect.t > 0)
						_hk -= srcRect.t;
					if (srcRect.b < 100)
						_hk -= (100 - srcRect.b);
					_hk = 100 / _hk;

					var _r = x + w;
					var _b = y + h;

					if (srcRect.l > 0)
					{
						x -= ((srcRect.l * _wk * w) / 100);
					}
					if (srcRect.t > 0)
					{
						y -= ((srcRect.t * _hk * h) / 100);
					}
					if (srcRect.r < 100)
					{
						_r += (((100 - srcRect.r) * _wk * w) / 100);
					}
					if (srcRect.b < 100)
					{
						_b += (((100 - srcRect.b) * _hk * h) / 100);
					}

					this.m_arrayPages[this.m_lPagesCount - 1].drawImage(img, x, y, _r - x, _b - y);

					if (bIsClip)
					{
						this.RestoreGrState();
					}
				}
			}
		},

		SetFont                  : function(font)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].SetFont(font);
		},
		FillText                 : function(x, y, text, cropX, cropW)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].FillText(x, y, text);
		},
		FillTextCode             : function(x, y, text, cropX, cropW)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].FillTextCode(x, y, text);
		},
		tg                       : function(gid, x, y, codePoints)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].tg(gid, x, y, codePoints);
		},
		FillText2                : function(x, y, text)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].FillText(x, y, text);
		},
		charspace                : function(space)
		{
		},
		SetIntegerGrid           : function(param)
		{
		},
		GetIntegerGrid           : function()
		{
		},
		GetFont                  : function()
		{
			if (0 != this.m_lPagesCount)
				return this.m_arrayPages[this.m_lPagesCount - 1].m_oFont;
			return null;
		},
		put_GlobalAlpha          : function(enable, alpha)
		{
		},
		Start_GlobalAlpha        : function()
		{
		},
		End_GlobalAlpha          : function()
		{
		},
		DrawHeaderEdit           : function(yPos)
		{
		},
		DrawFooterEdit           : function(yPos)
		{
		},
		drawCollaborativeChanges : function(x, y, w, h)
		{
		},
		drawSearchResult         : function(x, y, w, h)
		{
		},

		DrawEmptyTableLine : function(x1, y1, x2, y2)
		{
			// эта функция не для печати или сохранения вообще
		},
		DrawLockParagraph  : function(lock_type, x, y1, y2)
		{
			// эта функция не для печати или сохранения вообще
		},

		DrawLockObjectRect : function(lock_type, x, y, w, h)
		{
			// эта функция не для печати или сохранения вообще
		},

		DrawSpellingLine : function(y0, x0, x1, w)
		{
		},

		// smart methods for horizontal / vertical lines
		drawHorLine  : function(align, y, x, r, penW)
		{
			this.p_width(1000 * penW);
			this._s();

			var _y = y;
			switch (align)
			{
				case 0:
				{
					_y = y + penW / 2;
					break;
				}
				case 1:
				{
					break;
				}
				case 2:
				{
					_y = y - penW / 2;
				}
			}
			this._m(x, y);
			this._l(r, y);

			this.ds();

			this._e();
		},
		drawHorLine2 : function(align, y, x, r, penW)
		{
			this.p_width(1000 * penW);

			var _y = y;
			switch (align)
			{
				case 0:
				{
					_y = y + penW / 2;
					break;
				}
				case 1:
				{
					break;
				}
				case 2:
				{
					_y = y - penW / 2;
					break;
				}
			}

			this._s();
			this._m(x, (_y - penW));
			this._l(r, (_y - penW));
			this.ds();

			this._s();
			this._m(x, (_y + penW));
			this._l(r, (_y + penW));
			this.ds();

			this._e();
		},

		drawVerLine : function(align, x, y, b, penW)
		{
			this.p_width(1000 * penW);
			this._s();

			var _x = x;
			switch (align)
			{
				case 0:
				{
					_x = x + penW / 2;
					break;
				}
				case 1:
				{
					break;
				}
				case 2:
				{
					_x = x - penW / 2;
				}
			}
			this._m(_x, y);
			this._l(_x, b);

			this.ds();
		},
		
		DrawPolygon : function(oPath, lineWidth, shift)
		{
			this.p_width(lineWidth);
			this._s();
			
			var Points = oPath.Points;
			var nCount = Points.length;
			// берем предпоследнюю точку, т.к. последняя совпадает с первой
			var PrevX = Points[nCount - 2].X, PrevY = Points[nCount - 2].Y;
			var _x    = Points[nCount - 2].X,    _y = Points[nCount - 2].Y;
			var StartX, StartY;
			
			for (var nIndex = 0; nIndex < nCount; nIndex++)
			{
				if(PrevX > Points[nIndex].X)
				{
					_y = Points[nIndex].Y - shift;
				}
				else if(PrevX < Points[nIndex].X)
				{
					_y  = Points[nIndex].Y + shift;
				}
				
				if(PrevY < Points[nIndex].Y)
				{
					_x = Points[nIndex].X - shift;
				}
				else if(PrevY > Points[nIndex].Y)
				{
					_x = Points[nIndex].X + shift;
				}
				
				PrevX = Points[nIndex].X;
				PrevY = Points[nIndex].Y;
				
				if(nIndex > 0)
				{
					if (1 == nIndex)
					{
						StartX = _x;
						StartY = _y;
						this._m(_x, _y);
					}
					else
					{
						this._l(_x, _y);
					}
				}
			}
			
			this._l(StartX, StartY);
			this._z();
			this.ds();
		},

		// мега крутые функции для таблиц
		drawHorLineExt : function(align, y, x, r, penW, leftMW, rightMW)
		{
			this.drawHorLine(align, y, x + leftMW, r + rightMW, penW);
		},

		rect : function(x, y, w, h)
		{
			var _x = x;
			var _y = y;
			var _r = (x + w);
			var _b = (y + h);

			this._s();
			this._m(_x, _y);
			this._l(_r, _y);
			this._l(_r, _b);
			this._l(_x, _b);
			this._l(_x, _y);
		},

		TableRect : function(x, y, w, h)
		{
			this.rect(x, y, w, h);
			this.df();
		},

		put_PenLineJoin          : function(_join)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].put_PenLineJoin(_join);
		},
		put_TextureBounds        : function(x, y, w, h)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].put_TextureBounds(x, y, w, h);
		},
		put_TextureBoundsEnabled : function(val)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].put_TextureBoundsEnabled(val);
		},
		put_brushTexture         : function(src, mode)
		{
			if (src == null || src == undefined)
				src = "";

			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].put_brushTexture(src, mode);
		},
		put_BrushTextureAlpha    : function(alpha)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].put_BrushTextureAlpha(alpha);
		},
		put_BrushGradient        : function(gradFill, points, transparent)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].put_BrushGradient(gradFill, points, transparent);
		},

		// функции клиппирования
		AddClipRect    : function(x, y, w, h)
		{
			/*
			 this.b_color1(0, 0, 0, 255);
			 this.rect(x, y, w, h);
			 this.df();
			 return;
			 */

			var __rect = new _rect();
			__rect.x   = x;
			__rect.y   = y;
			__rect.w   = w;
			__rect.h   = h;
			this.GrState.AddClipRect(__rect);
			//this.ClipManager.AddRect(x, y, w, h);
		},
		RemoveClipRect : function()
		{
			//this.ClipManager.RemoveRect();
		},

		SetClip    : function(r)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].beginCommand(32);

			this.rect(r.x, r.y, r.w, r.h);

			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].endCommand(32);

			//this._s();
		},
		RemoveClip : function()
		{
			if (0 != this.m_lPagesCount)
			{
				this.m_arrayPages[this.m_lPagesCount - 1].beginCommand(64);
				this.m_arrayPages[this.m_lPagesCount - 1].endCommand(64);
			}
		},

		GetTransform : function()
		{
			if (0 != this.m_lPagesCount)
			{
				return this.m_arrayPages[this.m_lPagesCount - 1].m_oTransform;
			}
			return null;
		},
		GetLineWidth : function()
		{
			if (0 != this.m_lPagesCount)
			{
				return this.m_arrayPages[this.m_lPagesCount - 1].m_oPen.Size;
			}
			return 0;
		},
		GetPen       : function()
		{
			if (0 != this.m_lPagesCount)
			{
				return this.m_arrayPages[this.m_lPagesCount - 1].m_oPen;
			}
			return 0;
		},
		GetBrush     : function()
		{
			if (0 != this.m_lPagesCount)
			{
				return this.m_arrayPages[this.m_lPagesCount - 1].m_oBrush;
			}
			return 0;
		},

		drawFlowAnchor : function(x, y)
		{
		},

		SavePen    : function()
		{
			this.GrState.SavePen();
		},
		RestorePen : function()
		{
			this.GrState.RestorePen();
		},

		SaveBrush    : function()
		{
			this.GrState.SaveBrush();
		},
		RestoreBrush : function()
		{
			this.GrState.RestoreBrush();
		},

		SavePenBrush    : function()
		{
			this.GrState.SavePenBrush();
		},
		RestorePenBrush : function()
		{
			this.GrState.RestorePenBrush();
		},

		SaveGrState    : function()
		{
			this.GrState.SaveGrState();
		},
		RestoreGrState : function()
		{
			var _t                = this.m_oBaseTransform;
			this.m_oBaseTransform = null;
			this.GrState.RestoreGrState();
			this.m_oBaseTransform = _t;
		},

		RemoveLastClip : function()
		{
			var _t                = this.m_oBaseTransform;
			this.m_oBaseTransform = null;
			this.GrState.RemoveLastClip();
			this.m_oBaseTransform = _t;
		},
		RestoreLastClip : function()
		{
			var _t                = this.m_oBaseTransform;
			this.m_oBaseTransform = null;
			this.GrState.RestoreLastClip();
			this.m_oBaseTransform = _t;
		},

		StartClipPath : function()
		{
			this.private_removeVectors();

			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].beginCommand(32);
		},

		EndClipPath : function()
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].endCommand(32);

			this.private_restoreVectors();
		},

		SetTextPr : function(textPr, theme)
		{
			if (0 != this.m_lPagesCount)
			{
				var _page = this.m_arrayPages[this.m_lPagesCount - 1];

				if (theme && textPr && textPr.ReplaceThemeFonts)
					textPr.ReplaceThemeFonts(theme.themeElements.fontScheme);

				_page.m_oTextPr = textPr;
				if (theme)
					_page.m_oGrFonts.checkFromTheme(theme.themeElements.fontScheme, _page.m_oTextPr.RFonts);
				else
					_page.m_oGrFonts = _page.m_oTextPr.RFonts;
			}
		},

		SetFontInternal : function(name, size, style)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].SetFontInternal(name, size, style);
		},

		SetFontSlot : function(slot, fontSizeKoef)
		{
			if (0 != this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].SetFontSlot(slot, fontSizeKoef);
		},

		GetTextPr : function()
		{
			if (0 != this.m_lPagesCount)
				return this.m_arrayPages[this.m_lPagesCount - 1].m_oTextPr;
			return null;
		},

		DrawPresentationComment : function(type, x, y, w, h)
		{

		},

		private_removeVectors : function()
		{
			this._restoreDumpedVectors = this.VectorMemoryForPrint;

			if (this._restoreDumpedVectors != null)
			{
				this.VectorMemoryForPrint = null;
				if (0 != this.m_lPagesCount)
					this.m_arrayPages[this.m_lPagesCount - 1].VectorMemoryForPrint = null;
			}
		},

		private_restoreVectors : function()
		{
			if (null != this._restoreDumpedVectors)
			{
				this.VectorMemoryForPrint = this._restoreDumpedVectors;
				if (0 != this.m_lPagesCount)
					this.m_arrayPages[this.m_lPagesCount - 1].VectorMemoryForPrint = this._restoreDumpedVectors;
			}
			this._restoreDumpedVectors = null;
		},

		AddHyperlink : function(x, y, w, h, url, tooltip)
		{
			if (0 !== this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].AddHyperlink(x, y, w, h, url, tooltip);
		},

		AddLink : function(x, y, w, h, dx, dy, dPage)
		{
			if (0 !== this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].AddLink(x, y, w, h, dx, dy, dPage);
		},

		AddFormField : function(nX, nY, nW, nH, nBaseLineOffset, oForm)
		{
			if (0 !== this.m_lPagesCount)
				this.m_arrayPages[this.m_lPagesCount - 1].AddFormField(nX, nY, nW, nH, nBaseLineOffset, oForm);
		},

		DocInfo : function(props)
		{
			if (props)
			{
				this.Memory.WriteByte(CommandType.ctDocInfo);

				var nFlagPos = this.Memory.GetCurPosition();
				this.Memory.Skip(4);
				var nFlag = 0;

				if (props.asc_getTitle())
				{
					nFlag |= 1;
					this.Memory.WriteString(props.asc_getTitle());
				}
				if (props.asc_getCreator())
				{
					nFlag |= (1 << 1);
					this.Memory.WriteString(props.asc_getCreator());
				}
				if (props.asc_getSubject())
				{
					nFlag |= (1 << 2);
					this.Memory.WriteString(props.asc_getSubject());
				}
				if (props.asc_getKeywords())
				{
					nFlag |= (1 << 3);
					this.Memory.WriteString(props.asc_getKeywords());
				}

				var nEndPos = this.Memory.GetCurPosition();
				this.Memory.Seek(nFlagPos);
				this.Memory.WriteLong(nFlag);
				this.Memory.Seek(nEndPos);
			}
		},
		
		IsPdfRenderer : function()
		{
			return this.RENDERER_PDF_FLAG;
		}
	};

	var MATRIX_ORDER_PREPEND = 0;
	var MATRIX_ORDER_APPEND  = 1;


	function CMatrix()
	{
		this.sx  = 1.0;
		this.shx = 0.0;
		this.shy = 0.0;
		this.sy  = 1.0;
		this.tx  = 0.0;
		this.ty  = 0.0;
	}

	CMatrix.prototype =
	{
		Reset           : function()
		{
			this.sx  = 1.0;
			this.shx = 0.0;
			this.shy = 0.0;
			this.sy  = 1.0;
			this.tx  = 0.0;
			this.ty  = 0.0;
		},
		// трансформ
		Multiply        : function(matrix, order)
		{
			if (MATRIX_ORDER_PREPEND == order)
			{
				var m = new CMatrix();
				m.sx  = matrix.sx;
				m.shx = matrix.shx;
				m.shy = matrix.shy;
				m.sy  = matrix.sy;
				m.tx  = matrix.tx;
				m.ty  = matrix.ty;
				m.Multiply(this, MATRIX_ORDER_APPEND);
				this.sx  = m.sx;
				this.shx = m.shx;
				this.shy = m.shy;
				this.sy  = m.sy;
				this.tx  = m.tx;
				this.ty  = m.ty;
			}
			else
			{
				var t0   = this.sx * matrix.sx + this.shy * matrix.shx;
				var t2   = this.shx * matrix.sx + this.sy * matrix.shx;
				var t4   = this.tx * matrix.sx + this.ty * matrix.shx + matrix.tx;
				this.shy = this.sx * matrix.shy + this.shy * matrix.sy;
				this.sy  = this.shx * matrix.shy + this.sy * matrix.sy;
				this.ty  = this.tx * matrix.shy + this.ty * matrix.sy + matrix.ty;
				this.sx  = t0;
				this.shx = t2;
				this.tx  = t4;
			}
			return this;
		},
		// а теперь частные случаи трансформа (для удобного пользования)
		Translate       : function(x, y, order)
		{
			var m = new CMatrix();
			m.tx  = x;
			m.ty  = y;
			this.Multiply(m, order);
		},
		Scale           : function(x, y, order)
		{
			var m = new CMatrix();
			m.sx  = x;
			m.sy  = y;
			this.Multiply(m, order);
		},
		Rotate          : function(a, order)
		{
			var m   = new CMatrix();
			var rad = AscCommon.deg2rad(a);
			m.sx    = Math.cos(rad);
			m.shx   = Math.sin(rad);
			m.shy   = -Math.sin(rad);
			m.sy    = Math.cos(rad);
			this.Multiply(m, order);
		},
		RotateAt        : function(a, x, y, order)
		{
			this.Translate(-x, -y, order);
			this.Rotate(a, order);
			this.Translate(x, y, order);
		},
		// determinant
		Determinant     : function()
		{
			return this.sx * this.sy - this.shy * this.shx;
		},
		// invert
		Invert          : function()
		{
			var det = this.Determinant();
			if (0.0001 > Math.abs(det))
				return;
			var d = 1 / det;

			var t0   = this.sy * d;
			this.sy  = this.sx * d;
			this.shy = -this.shy * d;
			this.shx = -this.shx * d;

			var t4  = -this.tx * t0 - this.ty * this.shx;
			this.ty = -this.tx * this.shy - this.ty * this.sy;

			this.sx = t0;
			this.tx = t4;
			return this;
		},
		// transform point
		TransformPointX : function(x, y)
		{
			return x * this.sx + y * this.shx + this.tx;
		},
		TransformPointY : function(x, y)
		{
			return x * this.shy + y * this.sy + this.ty;
		},
		// calculate rotate angle
		GetRotation     : function()
		{
			var x1  = 0.0;
			var y1  = 0.0;
			var x2  = 1.0;
			var y2  = 0.0;
			var _x1 = this.TransformPointX(x1, y1);
			var _y1 = this.TransformPointY(x1, y1);
			var _x2 = this.TransformPointX(x2, y2);
			var _y2 = this.TransformPointY(x2, y2);

			var _y = _y2 - _y1;
			var _x = _x2 - _x1;

			if (Math.abs(_y) < 0.001)
			{
				if (_x > 0)
					return 0;
				else
					return 180;
			}
			if (Math.abs(_x) < 0.001)
			{
				if (_y > 0)
					return 90;
				else
					return 270;
			}

			var a = Math.atan2(_y, _x);
			a     = AscCommon.rad2deg(a);
			if (a < 0)
				a += 360;
			return a;
		},
		// сделать дубликата
		CreateDublicate : function()
		{
			var m = new CMatrix();
			m.sx  = this.sx;
			m.shx = this.shx;
			m.shy = this.shy;
			m.sy  = this.sy;
			m.tx  = this.tx;
			m.ty  = this.ty;
			return m;
		},

		IsIdentity  : function()
		{
			if (this.sx == 1.0 &&
				this.shx == 0.0 &&
				this.shy == 0.0 &&
				this.sy == 1.0 &&
				this.tx == 0.0 &&
				this.ty == 0.0)
			{
				return true;
			}
			return false;
		},
		IsIdentity2 : function()
		{
			if (this.sx == 1.0 &&
				this.shx == 0.0 &&
				this.shy == 0.0 &&
				this.sy == 1.0)
			{
				return true;
			}
			return false;
		},

		GetScaleValue : function()
		{
			var x1 = this.TransformPointX(0, 0);
			var y1 = this.TransformPointY(0, 0);
			var x2 = this.TransformPointX(1, 1);
			var y2 = this.TransformPointY(1, 1);
			return Math.sqrt(((x2-x1)*(x2-x1)+(y2-y1)*(y2-y1))/2);
		}
	};

	function GradientGetAngleNoRotate(_angle, _transform)
	{
		var x1 = 0.0;
		var y1 = 0.0;
		var x2 = 1.0;
		var y2 = 0.0;

		var _matrixRotate = new CMatrix();
		_matrixRotate.Rotate(-_angle / 60000);

		var _x11 = _matrixRotate.TransformPointX(x1, y1);
		var _y11 = _matrixRotate.TransformPointY(x1, y1);
		var _x22 = _matrixRotate.TransformPointX(x2, y2);
		var _y22 = _matrixRotate.TransformPointY(x2, y2);

		_matrixRotate = global_MatrixTransformer.Invert(_transform);

		var _x1 = _matrixRotate.TransformPointX(_x11, _y11);
		var _y1 = _matrixRotate.TransformPointY(_x11, _y11);
		var _x2 = _matrixRotate.TransformPointX(_x22, _y22);
		var _y2 = _matrixRotate.TransformPointY(_x22, _y22);

		var _y = _y2 - _y1;
		var _x = _x2 - _x1;

		var a = 0;
		if (Math.abs(_y) < 0.001)
		{
			if (_x > 0)
				a = 0;
			else
				a = 180;
		}
		else if (Math.abs(_x) < 0.001)
		{
			if (_y > 0)
				a = 90;
			else
				a = 270;
		}
		else
		{
			a = Math.atan2(_y, _x);
			a = AscCommon.rad2deg(a);
		}

		if (a < 0)
			a += 360;

		//console.log(a);
		return a * 60000;
	};

	var CMatrixL = CMatrix;

	function CGlobalMatrixTransformer()
	{
		this.TranslateAppend = function(m, _tx, _ty)
		{
			m.tx += _tx;
			m.ty += _ty;
		}
		this.ScaleAppend     = function(m, _sx, _sy)
		{
			m.sx *= _sx;
			m.shx *= _sx;
			m.shy *= _sy;
			m.sy *= _sy;
			m.tx *= _sx;
			m.ty *= _sy;
		}
		this.RotateRadAppend = function(m, _rad)
		{
			var _sx  = Math.cos(_rad);
			var _shx = Math.sin(_rad);
			var _shy = -Math.sin(_rad);
			var _sy  = Math.cos(_rad);

			var t0 = m.sx * _sx + m.shy * _shx;
			var t2 = m.shx * _sx + m.sy * _shx;
			var t4 = m.tx * _sx + m.ty * _shx;
			m.shy  = m.sx * _shy + m.shy * _sy;
			m.sy   = m.shx * _shy + m.sy * _sy;
			m.ty   = m.tx * _shy + m.ty * _sy;
			m.sx   = t0;
			m.shx  = t2;
			m.tx   = t4;
		}

		this.MultiplyAppend = function(m1, m2)
		{
			var t0 = m1.sx * m2.sx + m1.shy * m2.shx;
			var t2 = m1.shx * m2.sx + m1.sy * m2.shx;
			var t4 = m1.tx * m2.sx + m1.ty * m2.shx + m2.tx;
			m1.shy = m1.sx * m2.shy + m1.shy * m2.sy;
			m1.sy  = m1.shx * m2.shy + m1.sy * m2.sy;
			m1.ty  = m1.tx * m2.shy + m1.ty * m2.sy + m2.ty;
			m1.sx  = t0;
			m1.shx = t2;
			m1.tx  = t4;
		}

		this.Invert = function(m)
		{
			var newM = m.CreateDublicate();
			var det  = newM.sx * newM.sy - newM.shy * newM.shx;
			if (0.0001 > Math.abs(det))
				return newM;

			var d = 1 / det;

			var t0   = newM.sy * d;
			newM.sy  = newM.sx * d;
			newM.shy = -newM.shy * d;
			newM.shx = -newM.shx * d;

			var t4  = -newM.tx * t0 - newM.ty * newM.shx;
			newM.ty = -newM.tx * newM.shy - newM.ty * newM.sy;

			newM.sx = t0;
			newM.tx = t4;
			return newM;
		}

		this.MultiplyAppendInvert = function(m1, m2)
		{
			var m = this.Invert(m2);
			this.MultiplyAppend(m1, m);
		}

		this.MultiplyPrepend = function(m1, m2)
		{
			var m = new CMatrixL();
			m.sx  = m2.sx;
			m.shx = m2.shx;
			m.shy = m2.shy;
			m.sy  = m2.sy;
			m.tx  = m2.tx;
			m.ty  = m2.ty;
			this.MultiplyAppend(m, m1);
			m1.sx  = m.sx;
			m1.shx = m.shx;
			m1.shy = m.shy;
			m1.sy  = m.sy;
			m1.tx  = m.tx;
			m1.ty  = m.ty;
		}

		this.Reflect = function (matrix, isHorizontal, isVertical) {
			var m = new CMatrixL();
			m.shx = 0;
			m.sy  = 1;
			m.tx  = 0;
			m.ty  = 0;
			m.sx  = 1;
			m.shy = 0;
			if (isHorizontal && isVertical) {
				m.sx  = -1;
				m.sy = -1;
			}	else if (isHorizontal) {
				m.sx  = -1;
			} else if (isVertical) {
				m.sy = -1;
			} else {
				return;
			}
			this.MultiplyAppend(matrix, m);
		}

		this.CreateDublicateM = function(matrix)
		{
			var m = new CMatrixL();
			m.sx  = matrix.sx;
			m.shx = matrix.shx;
			m.shy = matrix.shy;
			m.sy  = matrix.sy;
			m.tx  = matrix.tx;
			m.ty  = matrix.ty;
			return m;
		}

		this.IsIdentity  = function(m)
		{
			if (m.sx == 1.0 &&
				m.shx == 0.0 &&
				m.shy == 0.0 &&
				m.sy == 1.0 &&
				m.tx == 0.0 &&
				m.ty == 0.0)
			{
				return true;
			}
			return false;
		}
		this.IsIdentity2 = function(m)
		{
			var eps = 0.00001;
			if (Math.abs(m.sx - 1.0) < eps &&
				Math.abs(m.shx) < eps &&
				Math.abs(m.shy) < eps &&
				Math.abs(m.sy - 1.0) < eps)
			{
				return true;
			}
			return false;
		}
	}

	function CClipManager()
	{
		this.clipRects  = [];
		this.curRect    = new _rect();
		this.BaseObject = null;

		this.AddRect    = function(x, y, w, h)
		{
			var _count = this.clipRects.length;
			if (0 == _count)
			{
				this.curRect.x = x;
				this.curRect.y = y;
				this.curRect.w = w;
				this.curRect.h = h;

				var _r                 = new _rect();
				_r.x                   = x;
				_r.y                   = y;
				_r.w                   = w;
				_r.h                   = h;
				this.clipRects[_count] = _r;

				this.BaseObject.SetClip(this.curRect);
			}
			else
			{
				this.BaseObject.RemoveClip();
				var _r = new _rect();
				_r.x   = x;
				_r.y   = y;
				_r.w   = w;
				_r.h   = h;

				this.clipRects[_count] = _r;
				this.curRect           = this.IntersectRect(this.curRect, _r);
				this.BaseObject.SetClip(this.curRect);
			}
		}
		this.RemoveRect = function()
		{
			var _count = this.clipRects.length;
			if (0 != _count)
			{
				this.clipRects.splice(_count - 1, 1);
				--_count;

				this.BaseObject.RemoveClip();

				if (0 != _count)
				{
					this.curRect.x = this.clipRects[0].x;
					this.curRect.y = this.clipRects[0].y;
					this.curRect.w = this.clipRects[0].w;
					this.curRect.h = this.clipRects[0].h;

					for (var i = 1; i < _count; i++)
						this.curRect = this.IntersectRect(this.curRect, this.clipRects[i]);

					this.BaseObject.SetClip(this.curRect);
				}
			}
		}

		this.IntersectRect = function(r1, r2)
		{
			var res = new _rect();
			res.x   = Math.max(r1.x, r2.x);
			res.y   = Math.max(r1.y, r2.y);
			res.w   = Math.min(r1.x + r1.w, r2.x + r2.w) - res.x;
			res.h   = Math.min(r1.y + r1.h, r2.y + r2.h) - res.y;

			if (0 > res.w)
				res.w = 0;
			if (0 > res.h)
				res.h = 0;

			return res;
		}
	}

	function CPen()
	{
		this.Color    = {R : 255, G : 255, B : 255, A : 255};
		this.Style    = 0;
		this.LineCap  = 0;
		this.LineJoin = 0;

		this.LineWidth = 1;
	}

	function CBrush()
	{
		this.Color1 = {R : 255, G : 255, B : 255, A : 255};
		this.Color2 = {R : 255, G : 255, B : 255, A : 255};
		this.Type   = 0;
	}

	function CTableMarkup(Table)
	{
		this.Internal =
		{
			RowIndex  : 0,
			CellIndex : 0,
			PageNum   : 0
		};
		this.Table    = Table;
		this.X        = 0; // Смещение таблицы от начала страницы до первой колонки

		this.Cols    = []; // массив ширин колонок
		this.Margins = []; // массив левых и правых маргинов

		this.Rows = []; // массив позиций, высот строк(для данной страницы)
		// Rows = [ { Y : , H :  }, ... ]

		this.CurCol = 0; // текущая колонка
		this.CurRow = 0; // текущая строка

		this.TransformX = 0;
		this.TransformY = 0;
	}

	CTableMarkup.prototype =
	{
		CreateDublicate : function()
		{
			var obj = new CTableMarkup(this.Table);

			obj.Internal = {RowIndex : this.Internal.RowIndex, CellIndex : this.Internal.CellIndex, PageNum : this.Internal.PageNum};
			obj.X        = this.X;

			var len = this.Cols.length;
			for (var i = 0; i < len; i++)
				obj.Cols[i] = this.Cols[i];

			len = this.Margins.length;
			for (var i = 0; i < len; i++)
				obj.Margins[i] = {Left : this.Margins[i].Left, Right : this.Margins[i].Right};

			len = this.Rows.length;
			for (var i = 0; i < len; i++)
				obj.Rows[i] = {Y : this.Rows[i].Y, H : this.Rows[i].H};

			obj.CurRow = this.CurRow;
			obj.CurCol = this.CurCol;

			return obj;
		},

		CorrectFrom : function()
		{
			this.X += this.TransformX;

			var _len = this.Rows.length;
			for (var i = 0; i < _len; i++)
			{
				this.Rows[i].Y += this.TransformY;
			}
		},

		CorrectTo : function()
		{
			this.X -= this.TransformX;

			var _len = this.Rows.length;
			for (var i = 0; i < _len; i++)
			{
				this.Rows[i].Y -= this.TransformY;
			}
		},

		Get_X : function()
		{
			return this.X;
		},

		Get_Y : function()
		{
			var _Y = 0;
			if (this.Rows.length > 0)
			{
				_Y = this.Rows[0].Y;
			}
			return _Y;
		}
	};

	function CTableOutline(Table, PageNum, X, Y, W, H)
	{
		this.Table   = Table;
		this.PageNum = PageNum;

		this.X = X;
		this.Y = Y;

		this.W = W;
		this.H = H;
	}

	var g_fontManager = new AscFonts.CFontManager();
	g_fontManager.Initialize(true);
	g_fontManager.SetHintsProps(true, true);

	var g_dDpiX = 96.0;
	var g_dDpiY = 96.0;

	var g_dKoef_mm_to_pix = g_dDpiX / 25.4;
	var g_dKoef_pix_to_mm = 25.4 / g_dDpiX;

	function _rect()
	{
		this.x = 0;
		this.y = 0;
		this.w = 0;
		this.h = 0;
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon']                          = window['AscCommon'] || {};
	window['AscCommon'].CGrRFonts                = CGrRFonts;
	window['AscCommon'].CFontSetup               = CFontSetup;
	window['AscCommon'].CGrState                 = CGrState;
	window['AscCommon'].CMemory                  = CMemory;
	window['AscCommon'].CDocumentRenderer        = CDocumentRenderer;
	window['AscCommon'].MATRIX_ORDER_PREPEND     = MATRIX_ORDER_PREPEND;
	window['AscCommon'].MATRIX_ORDER_APPEND      = MATRIX_ORDER_APPEND;
	window['AscCommon'].CMatrix                  = CMatrix;
	window['AscCommon'].CMatrixL                 = CMatrixL;
	window['AscCommon'].CGlobalMatrixTransformer = CGlobalMatrixTransformer;
	window['AscCommon'].CClipManager             = CClipManager;
	window['AscCommon'].CPen                     = CPen;
	window['AscCommon'].CBrush                   = CBrush;
	window['AscCommon'].CTableMarkup             = CTableMarkup;
	window['AscCommon'].CTableOutline            = CTableOutline;
	window['AscCommon']._rect                    = _rect;

	window['AscCommon'].global_MatrixTransformer = new CGlobalMatrixTransformer();
	window['AscCommon'].g_fontManager            = g_fontManager;
	window['AscCommon'].g_dDpiX                  = g_dDpiX;
	window['AscCommon'].g_dKoef_mm_to_pix        = g_dKoef_mm_to_pix;
	window['AscCommon'].g_dKoef_pix_to_mm        = g_dKoef_pix_to_mm;
	window['AscCommon'].GradientGetAngleNoRotate = GradientGetAngleNoRotate;

	window['AscCommon'].DashPatternPresets 		 = DashPatternPresets;
    window['AscCommon'].CommandType 		 	 = CommandType;
})(window);
