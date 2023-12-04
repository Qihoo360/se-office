/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

(function(){

    /**
	 * Class representing a base highlight annotation.
	 * @constructor
    */
    function CAnnotationTextMarkup(sName, nType, nPage, aRect, oDoc)
    {
        AscPDF.CAnnotationBase.call(this, sName, nType, nPage, aRect, oDoc);

        this._quads         = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._width         = undefined;
    }
    CAnnotationTextMarkup.prototype = Object.create(AscPDF.CAnnotationBase.prototype);
	CAnnotationTextMarkup.prototype.constructor = CAnnotationTextMarkup;

    CAnnotationTextMarkup.prototype.SetQuads = function(aQuads) {
        this._quads = aQuads;
    };
    CAnnotationTextMarkup.prototype.GetQuads = function() {
        return this._quads;
    };
    CAnnotationTextMarkup.prototype.SetWidth = function(nWidth) {
        this._width = nWidth;
    };
    CAnnotationTextMarkup.prototype.GetWidth = function() {
        return this._width;
    }; 
    CAnnotationTextMarkup.prototype.IsTextMarkup = function() {
        return true;
    };
    CAnnotationTextMarkup.prototype.AddToRedraw = function() {
        let oViewer = editor.getDocumentRenderer();
        if (oViewer.pagesInfo.pages[this.GetPage()])
            oViewer.pagesInfo.pages[this.GetPage()].needRedrawHighlights = true;
    };
    CAnnotationTextMarkup.prototype.IsInQuads = function(x, y) {
        let oOverlayCtx = editor.getDocumentRenderer().overlay.m_oContext;
        let aQuads      = this.GetQuads();
        
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints = aQuads[i];

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }

            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let X1 = oPoint1.x;
            let Y1 = oPoint1.y;
            let X2 = oPoint2.x;
            let Y2 = oPoint2.y;
            let X3 = oPoint3.x;
            let Y3 = oPoint3.y;
            let X4 = oPoint4.x;
            let Y4 = oPoint4.y;

            oOverlayCtx.beginPath();
            oOverlayCtx.moveTo(X1, Y1);
            oOverlayCtx.lineTo(X2, Y2);
            oOverlayCtx.lineTo(X4, Y4);
            oOverlayCtx.lineTo(X3, Y3);
            oOverlayCtx.closePath();

            if (oOverlayCtx.isPointInPath(x, y))
                return true;
        }

        return false;
    };
    CAnnotationTextMarkup.prototype.DrawSelected = function(overlay) {
        overlay.m_oContext.lineWidth    = 3;
        overlay.m_oContext.globalAlpha  = 1;
        overlay.m_oContext.strokeStyle  = "rgb(33, 117, 200)";
        overlay.m_oContext.beginPath();

        fillRegion(this.GetUnitedRegion(), overlay, this.GetPage());
        overlay.m_oContext.stroke();
    };
    CAnnotationTextMarkup.prototype.GetUnitedRegion = function() {
        if (this.unitedRegion)
            return this.unitedRegion;

        let aQuads  = this.GetQuads();
        let fUniter = AscGeometry.PolyBool.union;
        let resultRegion;

        //let time1 = performance.now();
        let aAllRegions = [];
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints = aQuads[i];

            aAllRegions.push({
                inverted : false,
                regions : [
                    [
                        [aPoints[0], aPoints[1]],
                        [aPoints[2], aPoints[3]],
                        [aPoints[6], aPoints[7]],
                        [aPoints[4], aPoints[5]]
                    ]
                ]
            });
        }

        if (aAllRegions.length > 1) {
            resultRegion = fUniter(aAllRegions[0], aAllRegions[1]);
            for (let i = 2; i < aAllRegions.length; i++) {
                resultRegion = fUniter(resultRegion, aAllRegions[i]);
            }
        }
        else {
            resultRegion = aAllRegions[0];
        }
        //let time2 = performance.now();
        //console.log("union: " + (time2 - time1));

        this.unitedRegion = resultRegion;
        return this.unitedRegion;
    };
    CAnnotationTextMarkup.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // quads
        let aQuads = this.GetQuads();
        let nLen = 0;
        for (let i = 0; i < aQuads.length; i++) {
            nLen += aQuads[i].length;
        }
        memory.WriteLong(nLen);  
        for (let i = 0; i < aQuads.length; i++) {
            for (let j = 0; j < aQuads[i].length; j++) {
                memory.WriteDouble(aQuads[i][j]);
            }
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);

        this._replies.forEach(function(reply) {
            reply.WriteToBinary(memory); 
        });
    };
    /**
	 * Class representing a highlight annotation.
	 * @constructor
    */
    function CAnnotationHighlight(sName, nPage, aRect, oDoc)
    {
        CAnnotationTextMarkup.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Highlight, nPage, aRect, oDoc);
    }
    CAnnotationHighlight.prototype = Object.create(CAnnotationTextMarkup.prototype);
	CAnnotationHighlight.prototype.constructor = CAnnotationHighlight;

    CAnnotationHighlight.prototype.IsHighlight = function() {
        return true;
    };

    CAnnotationHighlight.prototype.Draw = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;

        let oRGBFill = this.GetRGBColor(this.GetStrokeColor());

        let aQuads = this.GetQuads();
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints = aQuads[i];

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }
            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let dx1 = oPoint2.x - oPoint1.x;
            let dy1 = oPoint2.y - oPoint1.y;
            let dx2 = oPoint4.x - oPoint3.x;
            let dy2 = oPoint4.y - oPoint3.y;
            let angle1          = Math.atan2(dy1, dx1);
            let angle2          = Math.atan2(dy2, dx2);
            let rotationAngle   = angle1;

            AscPDF.startMultiplyMode(oGraphicsPDF.context);

            oGraphicsPDF.BeginPath();
            oGraphicsPDF.SetGlobalAlpha(this.GetOpacity());
            oGraphicsPDF.SetFillStyle(oRGBFill.r, oRGBFill.g, oRGBFill.b);

            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                let aMinRect = getMinRect(aPoints);

                oGraphicsPDF.SetIntegerGrid(true);
                oGraphicsPDF.Rect(aMinRect[0], aMinRect[1], aMinRect[2] - aMinRect[0], aMinRect[3] - aMinRect[1], true);
                oGraphicsPDF.SetIntegerGrid(false);
            }
            else {
                oGraphicsPDF.MoveTo(oPoint1.x, oPoint1.y);
                oGraphicsPDF.LineTo(oPoint2.x, oPoint2.y);
                oGraphicsPDF.LineTo(oPoint4.x, oPoint4.y);
                oGraphicsPDF.LineTo(oPoint3.x, oPoint3.y);
                oGraphicsPDF.ClosePath();
            }

            oGraphicsPDF.Fill();
            AscPDF.endMultiplyMode(oGraphicsPDF.context);
        }
    };
        
    /**
	 * Class representing a highlight annotation.
	 * @constructor
    */
    function CAnnotationUnderline(sName, nPage, aRect, oDoc)
    {
        CAnnotationTextMarkup.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Underline, nPage, aRect, oDoc);
    }
    CAnnotationUnderline.prototype = Object.create(CAnnotationTextMarkup.prototype);
	CAnnotationUnderline.prototype.constructor = CAnnotationUnderline;

    CAnnotationUnderline.prototype.Draw = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;

        let aQuads      = this.GetQuads();
        let oRGBFill    = this.GetRGBColor(this.GetStrokeColor());
        
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints     = aQuads[i];
            
            oGraphicsPDF.SetStrokeStyle(oRGBFill.r, oRGBFill.g, oRGBFill.b);
            oGraphicsPDF.BeginPath();

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }
            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let X1 = oPoint3.x
            let Y1 = oPoint3.y;
            let X2 = oPoint4.x;
            let Y2 = oPoint4.y;

            let dx1 = oPoint2.x - oPoint1.x;
            let dy1 = oPoint2.y - oPoint1.y;
            let dx2 = oPoint4.x - oPoint3.x;
            let dy2 = oPoint4.y - oPoint3.y;
            let angle1          = Math.atan2(dy1, dx1);
            let angle2          = Math.atan2(dy2, dx2);
            let rotationAngle   = angle1;

            let nSide;
            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                nSide = Math.abs(oPoint3.y - oPoint1.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
            }
            else {
                nSide = findMaxSideWithRotation(oPoint1.x, oPoint1.y, oPoint2.x, oPoint2.y, oPoint3.x, oPoint3.y, oPoint4.x, oPoint4.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
            }

            let nLineW      = oGraphicsPDF.GetLineWidth();
            let nIndentX    = Math.sin(rotationAngle) * nLineW * 1.5;
            let nIndentY    = Math.cos(rotationAngle) * nLineW * 1.5;

            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                oGraphicsPDF.HorLine(X1, X2, Y2 - nIndentY);
            }
            else {
                oGraphicsPDF.MoveTo(X1 + nIndentX, Y1 - nIndentY);
                oGraphicsPDF.LineTo(X2 + nIndentX, Y2 - nIndentY);
            }
            
            oGraphicsPDF.Stroke();
        }
    };
    
    /**
	 * Class representing a highlight annotation.
	 * @constructor
    */
    function CAnnotationStrikeout(sName, nPage, aRect, oDoc)
    {
        CAnnotationTextMarkup.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Strikeout, nPage, aRect, oDoc);
    }
    CAnnotationStrikeout.prototype = Object.create(CAnnotationTextMarkup.prototype);
	CAnnotationStrikeout.prototype.constructor = CAnnotationStrikeout;

    CAnnotationStrikeout.prototype.Draw = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;

        let aQuads      = this.GetQuads();
        let oRGBFill    = this.GetRGBColor(this.GetStrokeColor());

        for (let i = 0; i < aQuads.length; i++) {
            let aPoints = aQuads[i];

            oGraphicsPDF.BeginPath();

            oGraphicsPDF.SetGlobalAlpha(this.GetOpacity());
            oGraphicsPDF.SetStrokeStyle(oRGBFill.r, oRGBFill.g, oRGBFill.b);

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }
            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let dx1 = oPoint2.x - oPoint1.x;
            let dy1 = oPoint2.y - oPoint1.y;
            let dx2 = oPoint4.x - oPoint3.x;
            let dy2 = oPoint4.y - oPoint3.y;
            let angle1          = Math.atan2(dy1, dx1);
            let angle2          = Math.atan2(dy2, dx2);
            let rotationAngle   = angle1;

            let X1 = oPoint1.x + (oPoint3.x - oPoint1.x) / 2;
            let Y1 = oPoint1.y + (oPoint3.y - oPoint1.y) / 2;
            let X2 = oPoint2.x + (oPoint4.x - oPoint2.x) / 2;
            let Y2 = oPoint2.y + (oPoint4.y - oPoint2.y) / 2;

            let nSide;
            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                nSide = Math.abs(oPoint3.y - oPoint1.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
                oGraphicsPDF.HorLine(X1, X2, Y2);
            }
            else {
                nSide = findMaxSideWithRotation(oPoint1.x, oPoint1.y, oPoint2.x, oPoint2.y, oPoint3.x, oPoint3.y, oPoint4.x, oPoint4.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));

                oGraphicsPDF.MoveTo(X1, Y1);
                oGraphicsPDF.LineTo(X2, Y2);
            }
            
            oGraphicsPDF.Stroke();
        }
    };

    /**
	 * Class representing a squiggly annotation.
	 * @constructor
    */
    function CAnnotationSquiggly(sName, nPage, aRect, oDoc)
    {
        CAnnotationTextMarkup.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Squiggly, nPage, aRect, oDoc);
    }
    CAnnotationSquiggly.prototype = Object.create(CAnnotationTextMarkup.prototype);
	CAnnotationSquiggly.prototype.constructor = CAnnotationSquiggly;

    CAnnotationSquiggly.prototype.Draw = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;

        let aQuads      = this.GetQuads();
        let oRGBFill    = this.GetRGBColor(this.GetStrokeColor());
        
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints     = aQuads[i];
            
            oGraphicsPDF.SetStrokeStyle(oRGBFill.r, oRGBFill.g, oRGBFill.b);
            oGraphicsPDF.BeginPath();

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }
            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let X1 = oPoint3.x
            let Y1 = oPoint3.y;
            let X2 = oPoint4.x;
            let Y2 = oPoint4.y;

            let dx1 = oPoint2.x - oPoint1.x;
            let dy1 = oPoint2.y - oPoint1.y;
            let dx2 = oPoint4.x - oPoint3.x;
            let dy2 = oPoint4.y - oPoint3.y;
            let angle1          = Math.atan2(dy1, dx1);
            let angle2          = Math.atan2(dy2, dx2);
            let rotationAngle   = angle1;

            let nSide;
            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                nSide = Math.abs(oPoint3.y - oPoint1.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
            }
            else {
                nSide = findMaxSideWithRotation(oPoint1.x, oPoint1.y, oPoint2.x, oPoint2.y, oPoint3.x, oPoint3.y, oPoint4.x, oPoint4.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
            }

            let nLineW      = oGraphicsPDF.GetLineWidth();
            let nIndentX    = Math.sin(rotationAngle) * nLineW * 1.5;
            let nIndentY    = Math.cos(rotationAngle) * nLineW * 1.5;

            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                oGraphicsPDF.HorLine(X1, X2, Y2 - nIndentY);
            }
            else {
                oGraphicsPDF.MoveTo(X1 + nIndentX, Y1 - nIndentY);
                oGraphicsPDF.LineTo(X2 + nIndentX, Y2 - nIndentY);
            }
            
            oGraphicsPDF.Stroke();
        }
    };

    let CARET_SYMBOL = {
        None:       0,
        Paragraph:  1,
        Space:      2
    }

    /**
	 * Class representing a caret annotation.
	 * @constructor
    */
    function CAnnotationCaret(sName, nPage, aRect, oDoc)
    {
        CAnnotationTextMarkup.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Caret, nPage, aRect, oDoc);
        this._caretSymbol = CARET_SYMBOL.None;
    }
    CAnnotationCaret.prototype = Object.create(CAnnotationTextMarkup.prototype);
	CAnnotationCaret.prototype.constructor = CAnnotationCaret;

    CAnnotationCaret.prototype.Draw = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;

        let aQuads      = this.GetQuads();
        let oRGBFill    = this.GetRGBColor(this.GetStrokeColor());
        
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints     = aQuads[i];
            
            oGraphicsPDF.SetStrokeStyle(oRGBFill.r, oRGBFill.g, oRGBFill.b);
            oGraphicsPDF.BeginPath();

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }
            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let X1 = oPoint3.x
            let Y1 = oPoint3.y;
            let X2 = oPoint4.x;
            let Y2 = oPoint4.y;

            let dx1 = oPoint2.x - oPoint1.x;
            let dy1 = oPoint2.y - oPoint1.y;
            let dx2 = oPoint4.x - oPoint3.x;
            let dy2 = oPoint4.y - oPoint3.y;
            let angle1          = Math.atan2(dy1, dx1);
            let angle2          = Math.atan2(dy2, dx2);
            let rotationAngle   = angle1;

            let nSide;
            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                nSide = Math.abs(oPoint3.y - oPoint1.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
            }
            else {
                nSide = findMaxSideWithRotation(oPoint1.x, oPoint1.y, oPoint2.x, oPoint2.y, oPoint3.x, oPoint3.y, oPoint4.x, oPoint4.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
            }

            let nLineW      = oGraphicsPDF.GetLineWidth();
            let nIndentX    = Math.sin(rotationAngle) * nLineW * 1.5;
            let nIndentY    = Math.cos(rotationAngle) * nLineW * 1.5;

            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                oGraphicsPDF.HorLine(X1, X2, Y2 - nIndentY);
            }
            else {
                oGraphicsPDF.MoveTo(X1 + nIndentX, Y1 - nIndentY);
                oGraphicsPDF.LineTo(X2 + nIndentX, Y2 - nIndentY);
            }
            
            oGraphicsPDF.Stroke();
        }
    };
    CAnnotationCaret.prototype.SetCaretSymbol = function(nType) {
        this._caretSymbol = nType;
    };
    CAnnotationCaret.prototype.GetCaretSymbol = function() {
        return this._caretSymbol;
    };
    CAnnotationCaret.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // rectange diff
        let aRD = this.GetReqtangleDiff();
        if (aRD) {
            memory.annotFlags |= (1 << 15);
            for (let i = 0; i < 4; i++) {
                memory.WriteDouble(aRD[i]);
            }
        }
        
        // caret symbol
        let nCaretSymbol = this.GetCaretSymbol();
        if (nCaretSymbol != null) {
            memory.annotFlags |= (1 << 16);
            memory.WriteByte(nCaretSymbol);
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);

        this._replies.forEach(function(reply) {
            reply.WriteToBinary(memory); 
        });
    };

    function findMaxSideWithRotation(x1, y1, x2, y2, x3, y3, x4, y4) {
        // Найдите центр поворота
        const x_center = (x1 + x3) / 2;
        const y_center = (y1 + y3) / 2;
      
        // Найдите угол поворота
        const angle = Math.atan2(y3 - y1, x3 - x1);
      
        // Выполните поворот вершин прямоугольника на обратный угол
        const cosAngle = Math.cos(-angle);
        const sinAngle = Math.sin(-angle);
      
        const x1_rotated = cosAngle * (x1 - x_center) - sinAngle * (y1 - y_center) + x_center;
        const y1_rotated = sinAngle * (x1 - x_center) + cosAngle * (y1 - y_center) + y_center;
      
        const x2_rotated = cosAngle * (x2 - x_center) - sinAngle * (y2 - y_center) + x_center;
        const y2_rotated = sinAngle * (x2 - x_center) + cosAngle * (y2 - y_center) + y_center;
      
        const x3_rotated = cosAngle * (x3 - x_center) - sinAngle * (y3 - y_center) + x_center;
        const y3_rotated = sinAngle * (x3 - x_center) + cosAngle * (y3 - y_center) + y_center;
      
        const x4_rotated = cosAngle * (x4 - x_center) - sinAngle * (y4 - y_center) + x_center;
        const y4_rotated = sinAngle * (x4 - x_center) + cosAngle * (y4 - y_center) + y_center;
      
        // Найдите длины сторон
        const sideAB = Math.sqrt(Math.pow(x2_rotated - x1_rotated, 2) + Math.pow(y2_rotated - y1_rotated, 2));
        const sideBC = Math.sqrt(Math.pow(x3_rotated - x2_rotated, 2) + Math.pow(y3_rotated - y2_rotated, 2));
        const sideCD = Math.sqrt(Math.pow(x4_rotated - x3_rotated, 2) + Math.pow(y4_rotated - y3_rotated, 2));
        const sideDA = Math.sqrt(Math.pow(x1_rotated - x4_rotated, 2) + Math.pow(y1_rotated - y4_rotated, 2));
      
        // Найдите максимальную сторону
        const maxSide = Math.max(sideAB, sideBC, sideCD, sideDA);
      
        return maxSide;
    }

    function fillRegion(polygon, overlay, pageIndex)
    {
        let oViewer = editor.getDocumentRenderer();
        let nScale  = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom * (96 / oViewer.file.pages[pageIndex].Dpi);

        let xCenter = oViewer.width >> 1;
        if (oViewer.documentWidth > oViewer.width)
		{
			xCenter = (oViewer.documentWidth >> 1) - (oViewer.scrollX) >> 0;
		}
		let yPos    = oViewer.scrollY >> 0;
        let page    = oViewer.drawingPages[pageIndex];
        let w       = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let h       = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let indLeft = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
        let indTop  = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

        // рисуем всегда в пиксельной сетке. при наклонных линиях - +- 1 пиксел - ничего страшного
        let pointOffset = (overlay.m_oContext.lineWidth & 1) ? 0.5 : 0;
        
        for (let i = 0, countPolygons = polygon.regions.length; i < countPolygons; i++)
        {
            let region = polygon.regions[i];
            let countPoints = region.length;

            if (2 > countPoints)
                continue;

            let X = indLeft + region[0][0] * nScale;
            let Y = indTop + region[0][1] * nScale;

            overlay.m_oContext.moveTo((X >> 0) + pointOffset, (Y >> 0) + pointOffset);

            overlay.CheckPoint1(X, Y);
            overlay.CheckPoint2(X, Y);

            for (let j = 1, countPoints = region.length; j < countPoints; j++)
            {
                X = indLeft + region[j][0] * nScale;
                Y = indTop + region[j][1] * nScale;;

                overlay.m_oContext.lineTo((X >> 0) + pointOffset, (Y >> 0) + pointOffset);
                overlay.CheckPoint1(X, Y);
                overlay.CheckPoint2(X, Y);
            }

            overlay.m_oContext.closePath();
        }
    }
    function getMinRect(aPoints) {
        let xMax = aPoints[0], yMax = aPoints[1], xMin = xMax, yMin = yMax;
        for(let i = 1; i < aPoints.length; i++) {
            if (i % 2 == 0) {
                if(aPoints[i] < xMin)
                {
                    xMin = aPoints[i];
                }
                if(aPoints[i] > xMax)
                {
                    xMax = aPoints[i];
                }
            }
            else {
                if(aPoints[i] < yMin)
                {
                    yMin = aPoints[i];
                }

                if(aPoints[i] > yMax)
                {
                    yMax = aPoints[i];
                }
            }
        }

        return [xMin, yMin, xMax, yMax];
    }

    window["AscPDF"].CAnnotationTextMarkup  = CAnnotationTextMarkup;
    window["AscPDF"].CAnnotationHighlight   = CAnnotationHighlight;
    window["AscPDF"].CAnnotationUnderline   = CAnnotationUnderline;
    window["AscPDF"].CAnnotationStrikeout   = CAnnotationStrikeout;
    window["AscPDF"].CAnnotationSquiggly    = CAnnotationSquiggly;
    window["AscPDF"].CAnnotationCaret       = CAnnotationCaret;
})();

