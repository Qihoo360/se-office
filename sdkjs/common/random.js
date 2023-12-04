﻿/*
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

(function(window, undefined){

    window["AscCommon"] = window.AscCommon = (window["AscCommon"] || {});

    function RandObj(buf) { this.buf = buf; }
    RandObj.prototype.buffer = function() { return this.buf; };
    RandObj.prototype.base64 = function() { return AscCommon.Base64.encode(this.buf); };
    RandObj.prototype.hex = function() { return AscCommon.Hex.encode(this.buf); };
    RandObj.prototype.base58 = function() { return AscCommon.Base58.encode(this.buf); };

    AscCommon.randomBytes = function(count)
    {
        var array = new Uint8Array(count);
        if (window.crypto && window.crypto.getRandomValues)
            window.crypto.getRandomValues(array);
        else
        {
            for (var i = 0; i < count; i++)
                array[i] = (0x100 * Math.random()) >> 0;
        }
        return new RandObj(array);
    };

})(window);
