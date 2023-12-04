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
var moveTo=0,
    lineTo=1,
    arcTo=2,
    bezier3=3,
    bezier4=4,
    close=5;

// Import
var cToRad = AscFormat.cToRad;
var HitToArc = AscFormat.HitToArc;
var ArcToCurvers = AscFormat.ArcToCurvers;
var ArcToOnCanvas = AscFormat.ArcToOnCanvas;
var HitInLine = AscFormat.HitInLine;
var HitInBezier4 = AscFormat.HitInBezier4;
var HitInBezier3 = AscFormat.HitInBezier3;
var MOVE_DELTA = AscFormat.MOVE_DELTA;

    var History = AscCommon.History;

var cToRad2 = (Math.PI/60000)/180;


function CChangesDrawingsAddPathCommand(Class, oCommand, nIndex, bReverse){
    this.Type = AscDFH.historyitem_PathAddPathCommand;
    this.Command = oCommand;
    this.Index = nIndex;
    this.bReverse = bReverse;
	AscDFH.CChangesBase.call(this, Class);
}

	CChangesDrawingsAddPathCommand.prototype = Object.create(AscDFH.CChangesBase.prototype);
	CChangesDrawingsAddPathCommand.prototype.constructor = CChangesDrawingsAddPathCommand;

    CChangesDrawingsAddPathCommand.prototype.Undo = function(){
        if(this.bReverse){
            this.Class.ArrPathCommandInfo.splice(this.Index, 0, this.Command);
        }
        else{
            this.Class.ArrPathCommandInfo.splice(this.Index, 1);
        }
    };
    CChangesDrawingsAddPathCommand.prototype.Redo = function(){
        if(this.bReverse){
            this.Class.ArrPathCommandInfo.splice(this.Index, 1);
        }
        else{
            this.Class.ArrPathCommandInfo.splice(this.Index, 0, this.Command);
        }
    };


    CChangesDrawingsAddPathCommand.prototype.WriteToBinary = function(Writer){
        Writer.WriteLong(this.Index);
        Writer.WriteLong(this.Command.id);
        Writer.WriteBool(!!this.bReverse);
        switch(this.Command.id){
            case moveTo:
            case lineTo:
            {
                Writer.WriteString2(this.Command.X);
                Writer.WriteString2(this.Command.Y);
                break;
            }
            case bezier3:
            {
                Writer.WriteString2(this.Command.X0);
                Writer.WriteString2(this.Command.Y0);
                Writer.WriteString2(this.Command.X1);
                Writer.WriteString2(this.Command.Y1);
                break;
            }
            case bezier4:
            {
                Writer.WriteString2(this.Command.X0);
                Writer.WriteString2(this.Command.Y0);
                Writer.WriteString2(this.Command.X1);
                Writer.WriteString2(this.Command.Y1);
                Writer.WriteString2(this.Command.X2);
                Writer.WriteString2(this.Command.Y2);
                break;
            }
            case arcTo:
            {
                Writer.WriteString2(this.Command.hR);
                Writer.WriteString2(this.Command.wR);
                Writer.WriteString2(this.Command.stAng);
                Writer.WriteString2(this.Command.swAng);
                break;
            }
            case close:
            {
                break;
            }
        }
    };


    CChangesDrawingsAddPathCommand.prototype.ReadFromBinary = function(Reader){
        this.Index = Reader.GetLong();
        this.Command = {};
        this.Command.id = Reader.GetLong();
        this.bReverse = Reader.GetBool();
        switch(this.Command.id){
            case moveTo:
            case lineTo:
            {
                this.Command.X = Reader.GetString2();
                this.Command.Y = Reader.GetString2();
                break;
            }
            case bezier3:
            {
                this.Command.X0 = Reader.GetString2();
                this.Command.Y0 = Reader.GetString2();
                this.Command.X1 = Reader.GetString2();
                this.Command.Y1 = Reader.GetString2();
                break;
            }
            case bezier4:
            {
                this.Command.X0 = Reader.GetString2();
                this.Command.Y0 = Reader.GetString2();
                this.Command.X1 = Reader.GetString2();
                this.Command.Y1 = Reader.GetString2();
                this.Command.X2 = Reader.GetString2();
                this.Command.Y2 = Reader.GetString2();
                break;
            }
            case arcTo:
            {
                this.Command.hR = Reader.GetString2();
                this.Command.wR = Reader.GetString2();
                this.Command.stAng = Reader.GetString2();
                this.Command.swAng = Reader.GetString2();
                break;
            }
            case close:
            {
                break;
            }
        }
    };



    AscDFH.changesFactory[AscDFH.historyitem_PathSetStroke] = AscDFH.CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_PathSetExtrusionOk] = AscDFH.CChangesDrawingsBool;
    AscDFH.changesFactory[AscDFH.historyitem_PathSetFill] = AscDFH.CChangesDrawingsString;
    AscDFH.changesFactory[AscDFH.historyitem_PathSetPathH] = AscDFH.CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PathSetPathW] = AscDFH.CChangesDrawingsLong;
    AscDFH.changesFactory[AscDFH.historyitem_PathAddPathCommand] = CChangesDrawingsAddPathCommand;

    AscDFH.drawingsChangesMap[AscDFH.historyitem_PathSetStroke] = function(oClass, value){oClass.stroke = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_PathSetExtrusionOk] = function(oClass, value){oClass.extrusionOk = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_PathSetFill] = function(oClass, value){oClass.fill = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_PathSetPathH] = function(oClass, value){oClass.pathH = value;};
    AscDFH.drawingsChangesMap[AscDFH.historyitem_PathSetPathW] = function(oClass, value){oClass.pathW = value;};


function Path()
{
    AscFormat.CBaseFormatObject.call(this);
    this.stroke      = null;
    this.extrusionOk = null;
    this.fill        = null;
    this.pathH       = null;
    this.pathW       = null;

    this.ArrPathCommandInfo = [];
    this.ArrPathCommand = [];
}
AscFormat.InitClass(Path, AscFormat.CBaseFormatObject, AscDFH.historyitem_type_Path);
    Path.prototypeRefresh_RecalcData = function()
    {};
    Path.prototype.createDuplicate = function()
    {
        var p = new Path();
        p.setStroke(this.stroke);
        p.setExtrusionOk(this.extrusionOk);
        p.setFill(this.fill);
        p.setPathH(this.pathH);
        p.setPathW(this.pathW);
        for(var i = 0; i < this.ArrPathCommandInfo.length; ++i)
        {
            var command = this.ArrPathCommandInfo[i];
            switch (command.id)
            {
                case moveTo:
                case lineTo:
                {
                    var x = command.X;
                    var y = command.Y;
                    p.addPathCommand({id: command.id, X: x, Y: y});
                    break;
                }
                case bezier3:
                {
                    var X0 = command.X0;
                    var Y0 = command.Y0;
                    var X1 = command.X1;
                    var Y1 = command.Y1;
                    p.addPathCommand({id: bezier3, X0: X0, Y0: Y0, X1: X1, Y1: Y1});
                    break;
                }
                case bezier4:
                {
                    var X0 = command.X0;
                    var Y0 = command.Y0;
                    var X1 = command.X1;
                    var Y1 = command.Y1;
                    var X2 = command.X2;
                    var Y2 = command.Y2;
                    p.addPathCommand({id: bezier4, X0: X0, Y0: Y0, X1: X1, Y1: Y1, X2: X2, Y2: Y2});
                    break;
                }
                case arcTo:
                {
                    var hR    = command.hR;
                    var wR    = command.wR;
                    var stAng = command.stAng;
                    var swAng = command.swAng;
                    p.addPathCommand({id: arcTo, hR: hR, wR: wR, stAng: stAng, swAng: swAng});
                    break;
                }
                case close:
                {
                    p.addPathCommand({id:close});
                    break;
                }
            }
        }
        return p;
    };
    Path.prototype.setStroke = function(pr)
    {
        History.CanAddChanges() && History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_PathSetStroke, this.stroke, pr));
        this.stroke = pr;
    };
    Path.prototype.setExtrusionOk = function(pr)
    {
        History.CanAddChanges() && History.Add(new AscDFH.CChangesDrawingsBool(this, AscDFH.historyitem_PathSetExtrusionOk, this.extrusionOk, pr));
        this.extrusionOk = pr;
    };
    Path.prototype.setFill = function(pr)
    {
        History.CanAddChanges() && History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_PathSetFill, this.fill, pr));
        this.fill = pr;
    };
    Path.prototype.setPathH = function(pr)
    {
        History.CanAddChanges() && History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_PathSetPathH, this.pathH, pr));
        this.pathH = pr;
    };
    Path.prototype.setPathW = function(pr)
    {
        History.CanAddChanges() && History.Add(new AscDFH.CChangesDrawingsLong(this, AscDFH.historyitem_PathSetPathW, this.pathW, pr));
        this.pathW = pr;
    };
    Path.prototype.addPathCommand = function(cmd)
    {
        History.CanAddChanges() && History.Add(new CChangesDrawingsAddPathCommand(this, cmd, this.ArrPathCommandInfo.length));
        this.ArrPathCommandInfo.push(cmd);
    };
    Path.prototype.moveTo = function(x, y)
    {
        this.addPathCommand({id:moveTo, X:x, Y:y});
    };
    Path.prototype.lnTo = function(x, y)
    {
        this.addPathCommand({id:lineTo, X:x, Y:y});
    };
    Path.prototype.arcTo = function(wR, hR, stAng, swAng)
    {
        this.addPathCommand({id: arcTo, wR: wR, hR: hR, stAng: stAng, swAng: swAng});
    };
    Path.prototype.quadBezTo = function(x0, y0, x1, y1)
    {
        this.addPathCommand({id:bezier3, X0:x0, Y0:y0, X1:x1, Y1:y1});
    };
    Path.prototype.cubicBezTo = function(x0, y0, x1, y1, x2, y2)
    {
        this.addPathCommand({id:bezier4, X0:x0, Y0:y0, X1:x1, Y1:y1, X2:x2, Y2:y2});
    };
    Path.prototype.close = function()
    {
        if(this.ArrPathCommandInfo.length > 0) {
            this.addPathCommand({id:close});
        }
    };
	Path.prototype.calculateCommandCoord = function(oGdLst, sFormula, dFormulaCoeff, dNumberCoeff)
    {
        let dVal;
        dVal = oGdLst[sFormula];
        if(dVal !== undefined)
        {
            return dVal*dFormulaCoeff;
        }
        return parseInt(sFormula, 10)*dNumberCoeff;
    };
    Path.prototype.recalculate = function(gdLst, bResetPathsInfo)
    {
        var ch, cw;
        var dCustomPathCoeffW, dCustomPathCoeffH;
        if(this.pathW!=undefined)
        {
            if(this.pathW > MOVE_DELTA)
            {
                cw = (gdLst["w"]/this.pathW);
            }
            else
            {
                cw = 0;
            }
            dCustomPathCoeffW = cw;
        }
        else
        {
            cw = 1;
            dCustomPathCoeffW = 1/36000;
        }
        if(this.pathH!=undefined)
        {
            if(this.pathH > MOVE_DELTA)
            {
                ch = (gdLst["h"]/this.pathH);
            }
            else
            {
                ch = 0;
            }
            dCustomPathCoeffH = ch;
        }
        else
        {
            ch = 1;
            dCustomPathCoeffH = 1/36000;
        }
        var APCI=this.ArrPathCommandInfo, n = APCI.length, cmd;
        var x0, y0, x1, y1, x2, y2, wR, hR, stAng, swAng, lastX, lastY;
        for(var i=0; i<n; ++i)
        {
            cmd=APCI[i];
            switch(cmd.id)
            {
                case moveTo:
                case lineTo:
                {
                    x0 = this.calculateCommandCoord(gdLst, cmd.X, cw, dCustomPathCoeffW);
                    y0 = this.calculateCommandCoord(gdLst, cmd.Y, ch, dCustomPathCoeffH);
                    this.ArrPathCommand[i] ={id:cmd.id, X:x0, Y:y0};
                    lastX = x0;
                    lastY = y0;
                    break;
                }
                case bezier3:
                {
                    x0 = this.calculateCommandCoord(gdLst, cmd.X0, cw, dCustomPathCoeffW);
                    y0 = this.calculateCommandCoord(gdLst, cmd.Y0, ch, dCustomPathCoeffH);
                    x1 = this.calculateCommandCoord(gdLst, cmd.X1, cw, dCustomPathCoeffW);
                    y1 = this.calculateCommandCoord(gdLst, cmd.Y1, ch, dCustomPathCoeffH);
                    this.ArrPathCommand[i] = {id:bezier3, X0: x0, Y0: y0, X1: x1, Y1: y1};
                    lastX = x1;
                    lastY = y1;
                    break;
                }
                case bezier4:
                {
                    x0 = this.calculateCommandCoord(gdLst, cmd.X0, cw, dCustomPathCoeffW);
                    y0 = this.calculateCommandCoord(gdLst, cmd.Y0, ch, dCustomPathCoeffH);
                    x1 = this.calculateCommandCoord(gdLst, cmd.X1, cw, dCustomPathCoeffW);
                    y1 = this.calculateCommandCoord(gdLst, cmd.Y1, ch, dCustomPathCoeffH);
                    x2 = this.calculateCommandCoord(gdLst, cmd.X2, cw, dCustomPathCoeffW);
                    y2 = this.calculateCommandCoord(gdLst, cmd.Y2, ch, dCustomPathCoeffH);
                    this.ArrPathCommand[i] = {id:bezier4, X0:x0, Y0: y0, X1:x1, Y1:y1, X2:x2, Y2:y2};
                    lastX = x2;
                    lastY = y2;
                    break;
                }
                case arcTo:
                {
                    wR = this.calculateCommandCoord(gdLst, cmd.wR, cw, dCustomPathCoeffW);
                    hR = this.calculateCommandCoord(gdLst, cmd.hR, ch, dCustomPathCoeffH);

                    stAng = gdLst[cmd.stAng];
                    if(stAng===undefined)
                    {
                        stAng=parseInt(cmd.stAng, 10);
                    }

                    swAng = gdLst[cmd.swAng];
                    if(swAng===undefined)
                    {
                        swAng=parseInt(cmd.swAng, 10);
                    }


                    var a1 = stAng;
                    var a2 = stAng + swAng;
                    var a3 = swAng;

                    stAng = Math.atan2(ch * Math.sin(a1 * cToRad), cw * Math.cos(a1 * cToRad)) / cToRad;
                    swAng = Math.atan2(ch * Math.sin(a2 * cToRad), cw * Math.cos(a2 * cToRad)) / cToRad - stAng;

                    if((swAng > 0) && (a3 < 0)) swAng -= 21600000;
                    if((swAng < 0) && (a3 > 0)) swAng += 21600000;
                    if(swAng == 0 && a3 != 0) swAng = 21600000;

                    var a = wR;
                    var b = hR;
                    var sin2 = Math.sin(stAng*cToRad);
                    var cos2 = Math.cos(stAng*cToRad);
                    var _xrad = cos2 / a;
                    var _yrad = sin2 / b;
                    var l = 1 / Math.sqrt(_xrad * _xrad + _yrad * _yrad);
                    var xc = lastX - l * cos2;
                    var yc = lastY - l * sin2;

                    var sin1 = Math.sin((stAng+swAng)*cToRad);
                    var cos1 = Math.cos((stAng+swAng)*cToRad);
                    var _xrad1 = cos1 / a;
                    var _yrad1 = sin1 / b;
                    var l1 = 1 / Math.sqrt(_xrad1 * _xrad1 + _yrad1 * _yrad1);

                    this.ArrPathCommand[i]={id: arcTo,
                        stX: lastX,
                        stY: lastY,
                        wR: wR,
                        hR: hR,
                        stAng: stAng*cToRad,
                        swAng: swAng*cToRad};

                    lastX = xc + l1 * cos1;
                    lastY = yc + l1 * sin1;


                    break;
                }
                case close:
                {
                    this.ArrPathCommand[i]={id: close};
                    break;
                }
                default:
                {
                    break;
                }
            }
        }
        if(bResetPathsInfo){
            delete this.ArrPathCommandInfo;
        }
    };
    Path.prototype.recalculate2 = function(gdLst, bResetPathsInfo)
    {
        var k = 10e-10;
        var APCI=this.ArrPathCommandInfo, n = APCI.length, cmd;
        var stAng, swAng, lastX, lastY;
        for(var i=0; i<n; ++i)
        {
            cmd=APCI[i];
            switch(cmd.id)
            {
                case moveTo:
                case lineTo:
                {

                    lastX=cmd.X*k;
                    lastY=cmd.Y*k;
                    this.ArrPathCommand[i] ={id:cmd.id, X:lastX, Y:lastY};
                    break;
                }
                case bezier3:
                {

                    lastX=cmd.X1;
                    lastY=cmd.Y1;

                    this.ArrPathCommand[i]={id:bezier3, X0: cmd.X0*k, Y0: cmd.Y0*k, X1:lastX, Y1:lastY};

                    break;
                }
                case bezier4:
                {
                    lastX=cmd.X2;
                    lastY=cmd.Y2;

                    this.ArrPathCommand[i]={id:bezier4, X0: cmd.X0*k, Y0: cmd.Y0*k, X1: cmd.X1*k, Y1: cmd.Y1*k, X2:lastX, Y2:lastY};
                    break;
                }
                case arcTo:
                {


                    var a1 = cmd.stAng;
                    var a2 = cmd.stAng + cmd.swAng;
                    var a3 = cmd.swAng;

                    stAng = Math.atan2(k * Math.sin(a1 * cToRad), k * Math.cos(a1 * cToRad)) / cToRad;
                    swAng = Math.atan2(k * Math.sin(a2 * cToRad), k * Math.cos(a2 * cToRad)) / cToRad - cmd.stAng;

                    if((swAng > 0) && (a3 < 0)) swAng -= 21600000;
                    if((swAng < 0) && (a3 > 0)) swAng += 21600000;
                    if(swAng == 0 && a3 != 0) swAng = 21600000;

                    var a = cmd.wR*k;
                    var b = cmd.hR*k;
                    var sin2 = Math.sin(stAng*cToRad);
                    var cos2 = Math.cos(stAng*cToRad);
                    var _xrad = cos2 / a;
                    var _yrad = sin2 / b;
                    var l = 1 / Math.sqrt(_xrad * _xrad + _yrad * _yrad);
                    var xc = lastX - l * cos2;
                    var yc = lastY - l * sin2;

                    var sin1 = Math.sin((stAng+swAng)*cToRad);
                    var cos1 = Math.cos((stAng+swAng)*cToRad);
                    var _xrad1 = cos1 / a;
                    var _yrad1 = sin1 / b;
                    var l1 = 1 / Math.sqrt(_xrad1 * _xrad1 + _yrad1 * _yrad1);

                    this.ArrPathCommand[i]={id: arcTo,
                        stX: lastX,
                        stY: lastY,
                        wR: cmd.wR*k,
                        hR: cmd.hR*k,
                        stAng: stAng*cToRad,
                        swAng: swAng*cToRad};

                    lastX = xc + l1 * cos1;
                    lastY = yc + l1 * sin1;


                    break;
                }
                case close:
                {
                    this.ArrPathCommand[i]={id: close};
                    break;
                }
                default:
                {
                    break;
                }
            }
        }
       // if(bResetPathsInfo)
        {
            delete this.ArrPathCommandInfo;
        }
    };
    Path.prototype.draw = function(shape_drawer)
    {
        if (shape_drawer.bIsCheckBounds === true && this.fill == "none")
        {
            // это для текстур
            return;
        }
        var bIsDrawLast = false;
        var path = this.ArrPathCommand;
        shape_drawer._s();
        for(var j = 0, l = path.length; j < l; ++j)
        {
            var cmd=path[j];
            switch(cmd.id)
            {
                case moveTo:
                {
                    bIsDrawLast = true;
                    shape_drawer._m(cmd.X, cmd.Y);
                    break;
                }
                case lineTo:
                {
                    bIsDrawLast = true;
                    shape_drawer._l(cmd.X, cmd.Y);
                    break;
                }
                case bezier3:
                {
                    bIsDrawLast = true;
                    shape_drawer._c2(cmd.X0, cmd.Y0, cmd.X1, cmd.Y1);
                    break;
                }
                case bezier4:
                {
                    bIsDrawLast = true;
                    shape_drawer._c(cmd.X0, cmd.Y0, cmd.X1, cmd.Y1, cmd.X2, cmd.Y2);
                    break;
                }
                case arcTo:
                {
                    bIsDrawLast = true;
                    ArcToCurvers(shape_drawer, cmd.stX, cmd.stY, cmd.wR, cmd.hR, cmd.stAng, cmd.swAng, shape_drawer.Shape.calcGeometry);
                    break;
                }
                case close:
                {
                    shape_drawer._z();
                    break;
                }
            }
        }

        if (bIsDrawLast)
        {
            shape_drawer.drawFillStroke(true, this.fill, this.stroke && !shape_drawer.bIsNoStrokeAttack);
        }

        shape_drawer._e();
    };
    Path.prototype.check_bounds = function(checker, geom)
    {
        var path=this.ArrPathCommand;
        for(var j=0, l=path.length; j<l; ++j)
        {
            var cmd=path[j];
            switch(cmd.id)
            {
                case moveTo:
                {
                    checker._m(cmd.X, cmd.Y);
                    break;
                }
                case lineTo:
                {
                    checker._l(cmd.X, cmd.Y);
                    break;
                }
                case bezier3:
                {
                    checker._c2(cmd.X0, cmd.Y0, cmd.X1, cmd.Y1);
                    break;
                }
                case bezier4:
                {
                    checker._c(cmd.X0, cmd.Y0, cmd.X1, cmd.Y1, cmd.X2, cmd.Y2);
                    break;
                }
                case arcTo:
                {
                    ArcToCurvers(checker, cmd.stX, cmd.stY, cmd.wR, cmd.hR, cmd.stAng, cmd.swAng);
                    break;
                }
                case close:
                {
                    checker._z();
                    break;
                }
            }
        }
    };
    Path.prototype.hitInInnerArea = function(canvasContext, x, y)
    {
        if(this.fill === "none")
            return false;

        var _arr_commands = this.ArrPathCommand;
        var _commands_count = _arr_commands.length;
        var _command_index;
        var _command;
        canvasContext.beginPath();
        for(_command_index = 0; _command_index < _commands_count; ++_command_index)
        {
            _command = _arr_commands[_command_index];
            switch(_command.id)
            {
                case moveTo:
                {
                    canvasContext.moveTo(_command.X, _command.Y);
                    break;
                }
                case lineTo:
                {
                    canvasContext.lineTo(_command.X, _command.Y);
                    break;
                }
                case arcTo:
                {
                    ArcToOnCanvas(canvasContext, _command.stX, _command.stY, _command.wR, _command.hR, _command.stAng, _command.swAng);
                    break;
                }
                case bezier3:
                {
                    canvasContext.quadraticCurveTo(_command.X0, _command.Y0, _command.X1, _command.Y1);
                    break;
                }
                case bezier4:
                {
                    canvasContext.bezierCurveTo(_command.X0, _command.Y0, _command.X1, _command.Y1, _command.X2, _command.Y2);
                    break;
                }
                case close:
                {
                    canvasContext.closePath();
                }
            }
        }
	    return !!canvasContext.isPointInPath(x, y);
    };
    Path.prototype.hitInPath = function(canvasContext, x, y, oAddingPoint, _path_index)
    {
        var _arr_commands = this.ArrPathCommand;
        var _commands_count = _arr_commands.length;
        var _command_index;
        var _command;
        var _last_x, _last_y;
        var _begin_x, _begin_y;
        for(_command_index = 0; _command_index< _commands_count; ++_command_index)
        {
            _command = _arr_commands[_command_index];
            switch(_command.id)
            {
                case moveTo:
                {
                    _last_x = _command.X;
                    _last_y = _command.Y;
                    _begin_x = _command.X;
                    _begin_y = _command.Y;
                    break;
                }
                case lineTo:
                {
                    if(HitInLine(canvasContext, x, y, _last_x, _last_y, _command.X, _command.Y))
                        return true;
                    _last_x = _command.X;
                    _last_y = _command.Y;
                    break;
                }
                case arcTo:
                {
                    if(HitToArc(canvasContext, x, y,  _command.stX, _command.stY, _command.wR, _command.hR, _command.stAng, _command.swAng))
                        return true;
                    _last_x=(_command.stX-_command.wR*Math.cos(_command.stAng)+_command.wR*Math.cos(_command.swAng));
                    _last_y=(_command.stY-_command.hR*Math.sin(_command.stAng)+_command.hR*Math.sin(_command.swAng));
                    break;
                }
                case bezier3:
                {
                    if(HitInBezier3(canvasContext, x, y, _last_x, _last_y, _command.X0, _command.Y0, _command.X1, _command.Y1))
                        return true;
                    _last_x=_command.X1;
                    _last_y=_command.Y1;
                    break;
                }
                case bezier4:
                {
                    if(HitInBezier4(canvasContext, x, y, _last_x, _last_y, _command.X0, _command.Y0, _command.X1, _command.Y1, _command.X2, _command.Y2)) {
                        if(oAddingPoint) {
                            oAddingPoint.pathIndex = _path_index;
                            oAddingPoint.commandIndex = _command_index;
                        }
                        return true;
                    }
                    _last_x=_command.X2;
                    _last_y=_command.Y2;
                    break;
                }
                case close:
                {
                    if(HitInLine(canvasContext, x, y, _last_x, _last_y, _begin_x, _begin_y))
                        return true;
                }
            }
        }
        return false;
    };
    Path.prototype.isSmartLine  = function()
    {
        if (this.ArrPathCommand.length != 2)
            return false;

        if (this.ArrPathCommand[0].id == moveTo && this.ArrPathCommand[1].id == lineTo)
        {
            if (Math.abs(this.ArrPathCommand[0].X - this.ArrPathCommand[1].X) < 0.0001)
                return true;

            if (Math.abs(this.ArrPathCommand[0].Y - this.ArrPathCommand[1].Y) < 0.0001)
                return true;
        }

        return false;
    };
    Path.prototype.isSmartRect  = function()
    {
        if (this.ArrPathCommand.length != 5)
            return false;

        if (this.ArrPathCommand[0].id != moveTo ||
            this.ArrPathCommand[1].id != lineTo ||
            this.ArrPathCommand[2].id != lineTo ||
            this.ArrPathCommand[3].id != lineTo ||
            (this.ArrPathCommand[4].id != lineTo && this.ArrPathCommand[4].id != close))
            return false;

        var _float_eps = 0.0001;
        if (Math.abs(this.ArrPathCommand[0].X - this.ArrPathCommand[1].X) < _float_eps)
        {
            if (Math.abs(this.ArrPathCommand[1].Y - this.ArrPathCommand[2].Y) < _float_eps)
            {
                if (Math.abs(this.ArrPathCommand[2].X - this.ArrPathCommand[3].X) < _float_eps &&
                    Math.abs(this.ArrPathCommand[3].Y - this.ArrPathCommand[0].Y) < _float_eps)
                {
                    if (this.ArrPathCommand[4].id == close)
                        return true;

                    if (Math.abs(this.ArrPathCommand[0].X - this.ArrPathCommand[4].X) < _float_eps &&
                        Math.abs(this.ArrPathCommand[0].Y - this.ArrPathCommand[4].Y) < _float_eps)
                    {
                        return true;
                    }
                }
            }
        }
        else if (Math.abs(this.ArrPathCommand[0].Y - this.ArrPathCommand[1].Y) < _float_eps)
        {
            if (Math.abs(this.ArrPathCommand[1].X - this.ArrPathCommand[2].X) < _float_eps)
            {
                if (Math.abs(this.ArrPathCommand[2].Y - this.ArrPathCommand[3].Y) < _float_eps &&
                    Math.abs(this.ArrPathCommand[3].X - this.ArrPathCommand[0].X) < _float_eps)
                {
                    if (this.ArrPathCommand[4].id == close)
                        return true;

                    if (Math.abs(this.ArrPathCommand[0].X - this.ArrPathCommand[4].X) < _float_eps &&
                        Math.abs(this.ArrPathCommand[0].Y - this.ArrPathCommand[4].Y) < _float_eps)
                    {
                        return true;
                    }
                }
            }
        }

        return false;
    };
    Path.prototype.drawSmart  = function(shape_drawer)
    {
        var _graphics   = shape_drawer.Graphics;
        var _full_trans = _graphics.m_oFullTransform;

        if (!_graphics || !_full_trans || undefined == _graphics.m_bIntegerGrid || true === shape_drawer.bIsNoSmartAttack)
            return this.draw(shape_drawer);

        var bIsTransformed = (_full_trans.shx == 0 && _full_trans.shy == 0) ? false : true;

        if (bIsTransformed)
            return this.draw(shape_drawer);

        var isLine = this.isSmartLine();
        var isRect = false;
        if (!isLine)
            isRect = this.isSmartRect();

        if (window["NATIVE_EDITOR_ENJINE"] || ( !isLine && !isRect))
            return this.draw(shape_drawer);

        var _old_int = _graphics.m_bIntegerGrid;

        if (false == _old_int)
            _graphics.SetIntegerGrid(true);

        var dKoefMMToPx = Math.max(_graphics.m_oCoordTransform.sx, 0.001);

        var _ctx = _graphics.m_oContext;
        var bIsStroke = (shape_drawer.bIsNoStrokeAttack || (this.stroke !== true)) ? false : true;
        var bIsEven = false;
        if (bIsStroke)
        {
            var _lineWidth = Math.max((shape_drawer.StrokeWidth * dKoefMMToPx + 0.5) >> 0, 1);
            _ctx.lineWidth = _lineWidth;

            if ((_lineWidth & 0x01) == 0x01)
                bIsEven = true;

            if (_graphics.dash_no_smart)
            {
                for (var index = 0; index < _graphics.dash_no_smart.length; index++)
                    _graphics.dash_no_smart[index] = (_graphics.m_oCoordTransform.sx * _graphics.dash_no_smart[index] + 0.5) >> 0;

                _graphics.m_oContext.setLineDash(_graphics.dash_no_smart);
                _graphics.dash_no_smart = null;
            }
        }

        var bIsDrawLast = false;
        var path = this.ArrPathCommand;
        shape_drawer._s();

        if (!isRect)
        {
            for(var j = 0, l = path.length; j < l; ++j)
            {
                var cmd=path[j];
                switch(cmd.id)
                {
                    case moveTo:
                    {
                        bIsDrawLast = true;

                        var _x = (_full_trans.TransformPointX(cmd.X, cmd.Y)) >> 0;
                        var _y = (_full_trans.TransformPointY(cmd.X, cmd.Y)) >> 0;
                        if (bIsEven)
                        {
                            _x -= 0.5;
                            _y -= 0.5;
                        }
                        _ctx.moveTo(_x, _y);
                        break;
                    }
                    case lineTo:
                    {
                        bIsDrawLast = true;

                        var _x = (_full_trans.TransformPointX(cmd.X, cmd.Y)) >> 0;
                        var _y = (_full_trans.TransformPointY(cmd.X, cmd.Y)) >> 0;
                        if (bIsEven)
                        {
                            _x -= 0.5;
                            _y -= 0.5;
                        }
                        _ctx.lineTo(_x, _y);
                        break;
                    }
                    case close:
                    {
                        _ctx.closePath();
                        break;
                    }
                }
            }
        }
        else
        {
            var minX = 100000;
            var minY = 100000;
            var maxX = -100000;
            var maxY = -100000;
            bIsDrawLast = true;
            for(var j = 0, l = path.length; j < l; ++j)
            {
                var cmd=path[j];
                switch(cmd.id)
                {
                    case moveTo:
                    case lineTo:
                    {
                        if (minX > cmd.X)
                            minX = cmd.X;
                        if (minY > cmd.Y)
                            minY = cmd.Y;

                        if (maxX < cmd.X)
                            maxX = cmd.X;
                        if (maxY < cmd.Y)
                            maxY = cmd.Y;

                        break;
                    }
                    default:
                        break;
                }
            }

            var _x1 = (_full_trans.TransformPointX(minX, minY)) >> 0;
            var _y1 = (_full_trans.TransformPointY(minX, minY)) >> 0;
            var _x2 = (_full_trans.TransformPointX(maxX, maxY)) >> 0;
            var _y2 = (_full_trans.TransformPointY(maxX, maxY)) >> 0;

            if (bIsEven)
                _ctx.rect(_x1 + 0.5, _y1 + 0.5, _x2 - _x1, _y2 - _y1);
            else
                _ctx.rect(_x1, _y1, _x2 - _x1, _y2 - _y1);
        }

        if (bIsDrawLast)
        {
            shape_drawer.drawFillStroke(true, this.fill, bIsStroke);
        }

        shape_drawer._e();

        if (false == _old_int)
            _graphics.SetIntegerGrid(false);
    };
    Path.prototype.isEmpty  = function () {
        return this.ArrPathCommandInfo.length <= 0;
    };
    Path.prototype.checkBetweenPolygons = function (oBoundsController, oPolygonWrapper1, oPolygonWrapper2) {
    };
    Path.prototype.checkByPolygon = function (oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds) {
    };
    Path.prototype.transform = function (oTransform, dKoeff) {
    };
    Path.prototype.getSVGPath = function(oTransform, dStartX, dStartY) {
        var aCmds = this.ArrPathCommand;
        var sSVG = "";
        var oPresentation = editor.WordControl.m_oLogicDocument;
        var dSlideWidth = oPresentation.GetWidthMM();
        var dSlideHeight = oPresentation.GetHeightMM();
        var calcX = function(dX, dY) {
            var dX_ = oTransform.TransformPointX(dX, dY);
            return ((((dX_ - dStartX) / dSlideWidth) * 1000 + 0.5 >> 0) / 1000) + "";
        }
        var calcY = function(dX, dY) {
            var dY_ = oTransform.TransformPointY(dX, dY);
            return ((((dY_ - dStartY) / dSlideHeight) * 1000 + 0.5 >> 0) / 1000) + "";
        }
        let nLastCmd = null, nLastX = null, nLastY = null;
        for(var nCmd = 0; nCmd < aCmds.length; ++nCmd) {
            var oCmd = aCmds[nCmd];
            if(sSVG.length > 0) {
                sSVG += " ";
            }
            switch(oCmd.id) {
                case moveTo: {
                    if(nLastX !== null && nLastY !== null && AscFormat.fApproxEqual(nLastX, oCmd.X) && AscFormat.fApproxEqual(nLastY, oCmd.Y)) {
                        break;
                    }
                    sSVG += "M ";
                    sSVG += calcX(oCmd.X, oCmd.Y);
                    sSVG += " ";
                    sSVG += calcY(oCmd.X, oCmd.Y);
                    nLastX = oCmd.X;
                    nLastY = oCmd.Y;
                    break;
                }
                case lineTo: {
                    if(nLastX !== null && nLastY !== null && AscFormat.fApproxEqual(nLastX, oCmd.X) && AscFormat.fApproxEqual(nLastY, oCmd.Y)) {
                        break;
                    }
                    sSVG += "L ";
                    sSVG += calcX(oCmd.X, oCmd.Y);
                    sSVG += " ";
                    sSVG += calcY(oCmd.X, oCmd.Y);
                    nLastX = oCmd.X;
                    nLastY = oCmd.Y;
                    break;
                }
                case bezier4: {
                    sSVG += "C ";
                    sSVG += calcX(oCmd.X0, oCmd.Y0);
                    sSVG += " ";
                    sSVG += calcY(oCmd.X0, oCmd.Y0);
                    sSVG += " ";
                    sSVG += calcX(oCmd.X1, oCmd.Y1);
                    sSVG += " ";
                    sSVG += calcY(oCmd.X1, oCmd.Y1);
                    sSVG += " ";
                    sSVG += calcX(oCmd.X2, oCmd.Y2);
                    sSVG += " ";
                    sSVG += calcY(oCmd.X2, oCmd.Y2);
                    nLastX = oCmd.X2;
                    nLastY = oCmd.Y2;
                    break;
                }
                case close: {
                    sSVG += "Z";
                    nLastCmd = null;
                    nLastX = null;
                    nLastY = null;
                    break;
                }
                default: {
                    break;
                }
            }

            nLastCmd = oCmd.id;
        }
        return sSVG;
    };
    Path.prototype.isInk = function() {
        const nCmdCount = this.ArrPathCommandInfo.length;
        for(let nCmd = 0; nCmd < nCmdCount; ++nCmd) {
            let oCmd = this.ArrPathCommandInfo[nCmd];
            if(oCmd.id === close) {
                return false;
            }
            if(oCmd.id === arcTo) {
                return false;
            }
        }
        return true;
    };
    Path.prototype.convertToBezierCurves = function (oPath, oTransform) {
        const nCmdCount = this.ArrPathCommandInfo.length;
        let dX0, dY0, dX1, dY1, dX2, dY2;
        let oLastMoveTo = null;
        let dLastX, dLastY;
        let oFirstCmd = this.ArrPathCommand[0];
        if(!oFirstCmd) {
            return null;
        }
        if(oFirstCmd.id !== moveTo) {
            return null;
        }
        for(let nCmd = 0; nCmd < nCmdCount; ++nCmd) {
            let oCmd = this.ArrPathCommand[nCmd];
            switch (oCmd.id) {
                case moveTo: {
                    dX0 = oTransform.TransformPointX(oCmd.X, oCmd.Y)*36000 >> 0;
                    dY0 = oTransform.TransformPointY(oCmd.X, oCmd.Y)*36000 >> 0;
                    oPath.moveTo(dX0, dY0);
                    oLastMoveTo = oCmd;
                    dLastX = oCmd.X;
                    dLastY = oCmd.Y;
                    break;
                }
                case lineTo: {
                    dX0 = oTransform.TransformPointX(oCmd.X, oCmd.Y)*36000 >> 0;
                    dY0 = oTransform.TransformPointY(oCmd.X, oCmd.Y)*36000 >> 0;
                    dX0 = oTransform.TransformPointX(dLastX + (oCmd.X - dLastX)*(1/3), dLastY + (oCmd.Y - dLastY)*(1/3))*36000 >> 0;
                    dY0 = oTransform.TransformPointY(dLastX + (oCmd.X - dLastX)*(1/3), dLastY + (oCmd.Y - dLastY)*(1/3))*36000 >> 0;
                    dX1 = oTransform.TransformPointX(dLastX + (oCmd.X - dLastX)*(2/3), dLastY + (oCmd.Y - dLastY)*(2/3))*36000 >> 0;
                    dY1 = oTransform.TransformPointY(dLastX + (oCmd.X - dLastX)*(2/3), dLastY + (oCmd.Y - dLastY)*(2/3))*36000 >> 0;
                    dX2 = oTransform.TransformPointX(oCmd.X, oCmd.Y)*36000 >> 0;
                    dY2 = oTransform.TransformPointY(oCmd.X, oCmd.Y)*36000 >> 0;
                    oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
                    dLastX = oCmd.X;
                    dLastY = oCmd.Y;
                    break;
                }
                case bezier3: {
                    dX0 = oTransform.TransformPointX(oCmd.X0, oCmd.Y0)*36000 >> 0;
                    dY0 = oTransform.TransformPointY(oCmd.X0, oCmd.Y0)*36000 >> 0;
                    dX1 = oTransform.TransformPointX(oCmd.X1, oCmd.Y1)*36000 >> 0;
                    dY1 = oTransform.TransformPointY(oCmd.X1, oCmd.Y1)*36000 >> 0;
                    oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX1, dY1);
                    dLastX = oCmd.X1;
                    dLastY = oCmd.Y1;
                    break;
                }
                case bezier4: {
                    dX0 = oTransform.TransformPointX(oCmd.X0, oCmd.Y0)*36000 >> 0;
                    dY0 = oTransform.TransformPointY(oCmd.X0, oCmd.Y0)*36000 >> 0;
                    dX1 = oTransform.TransformPointX(oCmd.X1, oCmd.Y1)*36000 >> 0;
                    dY1 = oTransform.TransformPointY(oCmd.X1, oCmd.Y1)*36000 >> 0;
                    dX2 = oTransform.TransformPointX(oCmd.X2, oCmd.Y2)*36000 >> 0;
                    dY2 = oTransform.TransformPointY(oCmd.X2, oCmd.Y2)*36000 >> 0;
                    oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
                    dLastX = oCmd.X2;
                    dLastY = oCmd.Y2;
                    break;
                }
                case arcTo: {
                    let oPathAccumulator = new AscFormat.PathAccumulator();
                    ArcToCurvers(oPathAccumulator, oCmd.stX, oCmd.stY, oCmd.wR, oCmd.hR, oCmd.stAng, oCmd.swAng);
                    let aArcToCommands = oPathAccumulator.pathCommand;
                    for(let nArcCmd = 0; nArcCmd < aArcToCommands.length; ++nArcCmd)  {
                        let oArcToCmd = aArcToCommands[nArcCmd];
                        switch (oArcToCmd.id) {
                            case AscFormat.moveTo: {
                                break;
                            }
                            case AscFormat.bezier4: {
                                dX0 = oTransform.TransformPointX(oArcToCmd.X0, oArcToCmd.Y0)*36000 >> 0;
                                dY0 = oTransform.TransformPointY(oArcToCmd.X0, oArcToCmd.Y0)*36000 >> 0;
                                dX1 = oTransform.TransformPointX(oArcToCmd.X1, oArcToCmd.Y1)*36000 >> 0;
                                dY1 = oTransform.TransformPointY(oArcToCmd.X1, oArcToCmd.Y1)*36000 >> 0;
                                dX2 = oTransform.TransformPointX(oArcToCmd.X2, oArcToCmd.Y2)*36000 >> 0;
                                dY2 = oTransform.TransformPointY(oArcToCmd.X2, oArcToCmd.Y2)*36000 >> 0;
                                oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);

                                dLastX = oArcToCmd.X2;
                                dLastY = oArcToCmd.Y2;
                                break;
                            }
                        }
                    }
                    break;
                }
                case close: {
                    if(oLastMoveTo) {
                        let dXM = oTransform.TransformPointX(oLastMoveTo.X, oLastMoveTo.Y);
                        let dYM = oTransform.TransformPointY(oLastMoveTo.X, oLastMoveTo.Y);
                        let dLastXM = oTransform.TransformPointX(dLastX, dLastY);
                        let dLastYM = oTransform.TransformPointY(dLastX, dLastY);
                        dX0 = (dLastXM + (dXM - dLastXM) / 4) * 36000 >> 0;
                        dY0 = (dLastYM + (dYM - dLastYM) / 4) * 36000 >> 0;
                        dX1 = (dLastXM + 3 * (dXM - dLastXM) / 4) * 36000 >> 0;
                        dY1 = (dLastYM + 3 * (dYM - dLastYM) / 4) * 36000 >> 0;
                        dX2 = (dXM) * 36000 >> 0;
                        dY2 = (dYM) * 36000 >> 0;
                        oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
                    }
                    oPath.close();
                    break;
                }
            }
        }
        oPath.recalculate({}, true);
    };
    Path.prototype.isEqual = function(oPath) {
        if(!oPath) {
            return false;
        }
        if(this.stroke !== oPath.stroke) {
            return false;
        }
        if(this.extrusionOk !== oPath.extrusionOk) {
            return false;
        }
        if(this.fill !== oPath.fill) {
            return false;
        }
        if(this.pathH !== oPath.pathH) {
            return false;
        }
        if(this.pathW !== oPath.pathW) {
            return false;
        }
        if(this.ArrPathCommandInfo.length !== oPath.ArrPathCommandInfo.length) {
            return false;
        }
        for(let nCmd = 0; nCmd < this.ArrPathCommandInfo.length; ++nCmd) {
            let oCmd1 = this.ArrPathCommandInfo[nCmd];
            let oCmd2 = oPath.ArrPathCommandInfo[nCmd];
            if(oCmd1.id !== oCmd2.id) {
                return false;
            }
            switch(oCmd1.id)
            {
                case moveTo:
                case lineTo:
                {
                    if(oCmd1.X !== oCmd2.X) {
                        return false;
                    }
                    if(oCmd1.Y !== oCmd2.Y) {
                        return false;
                    }
                    break;
                }
                case bezier3:
                {
                    if(oCmd1.X0 !== oCmd2.X0) {
                        return false;
                    }
                    if(oCmd1.X1 !== oCmd2.X1) {
                        return false;
                    }
                    if(oCmd1.Y0 !== oCmd2.Y0) {
                        return false;
                    }
                    if(oCmd1.Y1 !== oCmd2.Y1) {
                        return false;
                    }
                    break;
                }
                case bezier4:
                {

                    if(oCmd1.X0 !== oCmd2.X0) {
                        return false;
                    }
                    if(oCmd1.X1 !== oCmd2.X1) {
                        return false;
                    }
                    if(oCmd1.X2 !== oCmd2.X2) {
                        return false;
                    }
                    if(oCmd1.Y0 !== oCmd2.Y0) {
                        return false;
                    }
                    if(oCmd1.Y1 !== oCmd2.Y1) {
                        return false;
                    }
                    if(oCmd1.Y2 !== oCmd2.Y2) {
                        return false;
                    }
                    break;
                }
                case arcTo:
                {

                    if(oCmd1.wR !== oCmd2.wR) {
                        return false;
                    }
                    if(oCmd1.hR !== oCmd2.hR) {
                        return false;
                    }
                    if(oCmd1.stAng !== oCmd2.stAng) {
                        return false;
                    }
                    if(oCmd1.swAng !== oCmd2.swAng) {
                        return false;
                    }
                    break;
                }
                case close:
                {
                    break;
                }
                default:
                {
                    break;
                }
            }
        }
        return true;
    };
    function CPathCmd() {
        AscFormat.CBaseNoIdObject.call(this);
        this.pts = [];
    }
    AscFormat.InitClass(CPathCmd, AscFormat.CBaseNoIdObject, 0);

    function CheckPointByPaths(dX, dY, dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2)
    {
        var cX, cY, point, topX, topY, bottomX, bottomY;
        cX = (dX - dMinX)/dWidth;
        cY = (dY - dMinY)/dHeight;
        if(cX > 1)
        {
            cX = 1;
        }
        if(cX < 0)
        {
            cX = 0;
        }
        point = oPolygonWrapper1.getPointOnPolygon(cX);
        topX = point.x;
        topY = point.y;
        point = oPolygonWrapper2.getPointOnPolygon(cX);
        bottomX = point.x;
        bottomY = point.y;
        return {x: topX + cY*(bottomX - topX), y: topY + cY*(bottomY - topY)};
    }

    function Path2(oPathMemory) {
        this.stroke = null;
        this.extrusionOk = null;
        this.fill = null;
        this.pathH = null;
        this.pathW = null;


        this.startPos = 0;
        this.size = 25;

        this.PathMemory = oPathMemory;
        this.ArrPathCommand = oPathMemory.ArrPathCommand;
        this.curLen = 0;

        this.lastX = null;
        this.lastY = null;
        this.bEmpty = true;
    }

    Path2.prototype.isEmpty = function () {
        return this.bEmpty;
    };
    Path2.prototype.isInk = function () {
        return false;
    };
    Path2.prototype.checkArray = function (nSize) {
        this.bEmpty = false;
        this.ArrPathCommand[this.startPos] += nSize;
        this.PathMemory.curPos = this.startPos + this.ArrPathCommand[this.startPos];
        if (this.PathMemory.curPos + 1 > this.ArrPathCommand.length) {

            var aNewArray = new Float64Array((((3 / 2) * (this.PathMemory.curPos + 1)) >> 0) + 1);
            for (var i = 0; i < this.ArrPathCommand.length; ++i) {
                aNewArray[i] = this.ArrPathCommand[i];
            }
            this.PathMemory.ArrPathCommand = aNewArray;
            this.ArrPathCommand = aNewArray;
        }
    };
    Path2.prototype.setStroke = function (pr) {

        this.stroke = pr;
    };
    Path2.prototype.setExtrusionOk = function (pr) {

        this.extrusionOk = pr;
    };
    Path2.prototype.setFill = function (pr) {

        this.fill = pr;
    };
    Path2.prototype.setPathH = function (pr) {

        this.pathH = pr;
    };
    Path2.prototype.setPathW = function (pr) {

        this.pathW = pr;
    };
    Path2.prototype.addPathCommand = function (cmd) {

        this.ArrPathCommand.push(cmd);
    };
    Path2.prototype.moveTo = function (x, y) {
        this.lastX = x * 10e-10;
        this.lastY = y * 10e-10;
        this.checkArray(3);
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = moveTo;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = this.lastX;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = this.lastY;
    };
    Path2.prototype.lnTo = function (x, y) {

        this.lastX = x * 10e-10;
        this.lastY = y * 10e-10;
        this.checkArray(3);
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = lineTo;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = this.lastX;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = this.lastY;
    };
    Path2.prototype.arcTo = function (wR, hR, stAng, swAng) {
        var nSign = swAng < 0 ? -1 : 1;
        if (AscFormat.fApproxEqual(Math.abs(swAng), 21600000)) {
            swAng = nSign * (21600000 - 1);
        }
        var a1 = stAng;
        var a2 = stAng + swAng;
        var a3 = swAng;


        stAng = Math.atan2(10e-10 * Math.sin(a1 * cToRad), 10e-10 * Math.cos(a1 * cToRad)) / cToRad;
        swAng = Math.atan2(10e-10 * Math.sin(a2 * cToRad), 10e-10 * Math.cos(a2 * cToRad)) / cToRad - stAng;

        if ((swAng > 0) && (a3 < 0)) swAng -= 21600000;
        if ((swAng < 0) && (a3 > 0)) swAng += 21600000;
        if (swAng == 0 && a3 != 0) swAng = 21600000;

        var a = wR * 10e-10;
        var b = hR * 10e-10;
        var sin2 = Math.sin(stAng * cToRad);
        var cos2 = Math.cos(stAng * cToRad);
        var _xrad = cos2 / a;
        var _yrad = sin2 / b;
        var l = 1 / Math.sqrt(_xrad * _xrad + _yrad * _yrad);
        var xc = this.lastX - l * cos2;
        var yc = this.lastY - l * sin2;

        var sin1 = Math.sin((stAng + swAng) * cToRad);
        var cos1 = Math.cos((stAng + swAng) * cToRad);
        var _xrad1 = cos1 / a;
        var _yrad1 = sin1 / b;
        var l1 = 1 / Math.sqrt(_xrad1 * _xrad1 + _yrad1 * _yrad1);


        this.checkArray(7);
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = arcTo;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = this.lastX;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = this.lastY;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = wR * 10e-10;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = hR * 10e-10;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = stAng * cToRad;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = swAng * cToRad;
    };
    Path2.prototype.quadBezTo = function (x0, y0, x1, y1) {


        this.lastX = x1 * 10e-10;
        this.lastY = y1 * 10e-10;


        this.checkArray(5);
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = bezier3;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = x0 * 10e-10;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = y0 * 10e-10;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = this.lastX;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = this.lastY;
    };
    Path2.prototype.cubicBezTo = function (x0, y0, x1, y1, x2, y2) {


        this.lastX = x2 * 10e-10;
        this.lastY = y2 * 10e-10;

        this.checkArray(7);
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = bezier4;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = x0 * 10e-10;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = y0 * 10e-10;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = x1 * 10e-10;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = y1 * 10e-10;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = this.lastX;
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = this.lastY;
    };
    Path2.prototype.close = function () {
        this.checkArray(1);
        this.ArrPathCommand[this.startPos + (this.curLen++) + 1] = close;
    };
    Path2.prototype.draw = function (shape_drawer) {
        if (this.isEmpty()) {
            return;
        }
        if (shape_drawer.bIsCheckBounds === true && this.fill == "none") {
            // это для текстур
            return;
        }
        var bIsDrawLast = false;
        var path = this.ArrPathCommand;
        shape_drawer._s();
        var i = 0;
        var len = this.PathMemory.ArrPathCommand[this.startPos];
        while (i < len) {
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo: {
                    bIsDrawLast = true;
                    shape_drawer._m(path[this.startPos + i + 2], path[this.startPos + i + 3]);
                    i += 3;
                    break;
                }
                case lineTo: {
                    bIsDrawLast = true;
                    shape_drawer._l(path[this.startPos + i + 2], path[this.startPos + i + 3]);
                    i += 3;
                    break;
                }
                case bezier3: {
                    bIsDrawLast = true;
                    shape_drawer._c2(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5]);
                    i += 5;
                    break;
                }
                case bezier4: {
                    bIsDrawLast = true;
                    shape_drawer._c(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
                    i += 7;
                    break;
                }
                case arcTo: {
                    bIsDrawLast = true;
                    ArcToCurvers(shape_drawer, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
                    i += 7;
                    break;
                }
                case close: {
                    shape_drawer._z();
                    i += 1;
                    break;
                }
            }
        }

        if (bIsDrawLast) {
            shape_drawer.drawFillStroke(true, "normal", this.stroke && !shape_drawer.bIsNoStrokeAttack);
        }

        shape_drawer._e();
    };
    Path2.prototype.hitInInnerArea = function (canvasContext, x, y) {
        var path = this.ArrPathCommand;
        canvasContext.beginPath();
        var i = 0;
        var len = this.PathMemory.ArrPathCommand[this.startPos];
        while (i < len) {
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo: {
                    canvasContext.moveTo(path[this.startPos + i + 2], path[this.startPos + i + 3]);
                    i += 3;
                    break;
                }
                case lineTo: {
                    canvasContext.lineTo(path[this.startPos + i + 2], path[this.startPos + i + 3]);
                    i += 3;
                    break;
                }
                case bezier3: {
                    canvasContext.quadraticCurveTo(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5]);
                    i += 5;
                    break;
                }
                case bezier4: {
                    canvasContext.bezierCurveTo(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
                    i += 7;
                    break;
                }
                case arcTo: {
                    ArcToOnCanvas(canvasContext, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
                    i += 7;
                    break;
                }
                case close: {
                    canvasContext.closePath();
                    i += 1;
                    break;
                }
            }
        }
        canvasContext.closePath();
        return !!canvasContext.isPointInPath(x, y);
    };
    Path2.prototype.hitInPath = function (canvasContext, x, y) {
        var _arr_commands = this.ArrPathCommand;
        var _commands_count = _arr_commands.length;
        var _command_index;
        var _command;
        var _last_x, _last_y;
        var _begin_x, _begin_y;

        var path = this.ArrPathCommand;
        var i = 0;
        var len = this.PathMemory.ArrPathCommand[this.startPos];
        while (i < len) {
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo: {
                    canvasContext.moveTo(path[this.startPos + i + 2], path[this.startPos + i + 3]);
                    _last_x = path[this.startPos + i + 2];
                    _last_y = path[this.startPos + i + 3];
                    _begin_x = path[this.startPos + i + 2];
                    _begin_y = path[this.startPos + i + 3];
                    i += 3;
                    break;
                }
                case lineTo: {
                    if (HitInLine(canvasContext, x, y, _last_x, _last_y, path[this.startPos + i + 2], path[this.startPos + i + 3]))
                        return true;
                    _last_x = path[this.startPos + i + 2];
                    _last_y = path[this.startPos + i + 3];
                    i += 3;
                    break;
                }
                case bezier3: {
                    canvasContext.quadraticCurveTo(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5]);
                    if (HitInBezier3(canvasContext, x, y, _last_x, _last_y, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5]))
                        return true;
                    _last_x = path[this.startPos + i + 4];
                    _last_y = path[this.startPos + i + 5];
                    i += 5;
                    break;
                }
                case bezier4: {
                    if (HitInBezier4(canvasContext, x, y, _last_x, _last_y, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]))
                        return true;
                    _last_x = path[this.startPos + i + 6];
                    _last_y = path[this.startPos + i + 7];
                    i += 7;
                    break;
                }
                case arcTo: {
                    if (HitToArc(canvasContext, x, y, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]))
                        return true;
                    _last_x = (path[this.startPos + i + 2] - path[this.startPos + i + 4] * Math.cos(path[this.startPos + i + 6]) + path[this.startPos + i + 4] * Math.cos(path[this.startPos + i + 7]));
                    _last_y = (path[this.startPos + i + 3] - path[this.startPos + i + 5] * Math.sin(path[this.startPos + i + 6]) + path[this.startPos + i + 5] * Math.sin(path[this.startPos + i + 7]));
                    i += 7;
                    break;
                }
                case close: {
                    if (HitInLine(canvasContext, x, y, _last_x, _last_y, _begin_x, _begin_y))
                        return true;
                    i += 1;
                    break;
                }
            }
        }
        return false;
    };
    Path2.prototype.drawTracks = function (drawingDocument, transform) {

        var oApi = Asc.editor || editor;
        var isDrawHandles = oApi ? oApi.isShowShapeAdjustments() : true;
        var i = 0;
        var len = this.PathMemory.ArrPathCommand[this.startPos];

        var path = this.ArrPathCommand;
        var dDist = 0;
        while (i < len) {
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo: {
                    drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i + 2] - dDist, path[this.startPos + i + 3] - dDist, 2 * dDist, 2 * dDist, false, false, undefined, isDrawHandles);
                    i += 3;
                    break;
                }
                case lineTo: {
                    drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i + 2] - dDist, path[this.startPos + i + 3] - dDist, 2 * dDist, 2 * dDist, false, false, undefined, isDrawHandles);
                    i += 3;
                    break;
                }
                case bezier3: {
                    //  drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i+2] - dDist, path[this.startPos + i + 3] - dDist, 2*dDist, 2*dDist, false, false);
                    drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i + 4] - dDist, path[this.startPos + i + 5] - dDist, 2 * dDist, 2 * dDist, false, false, undefined, isDrawHandles);
                    i += 5;
                    break;
                }
                case bezier4: {
                    //  drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i+2] - dDist, path[this.startPos + i + 3] - dDist, 2*dDist,2*dDist, false, false);
                    // drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i+4] - dDist, path[this.startPos + i + 5] - dDist, 2*dDist,2*dDist, false, false); i+=7;
                    drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, path[this.startPos + i + 6] - dDist, path[this.startPos + i + 7] - dDist, 2 * dDist, 2 * dDist, false, false, undefined, isDrawHandles);
                    i += 7;
                    break;
                }
                case arcTo: {
                    var path_accumulator = new AscFormat.PathAccumulator();
                    ArcToCurvers(path_accumulator, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
                    var arc_to_path_commands = path_accumulator.pathCommand;
                    var lastX, lastY;
                    for (var arc_to_path_index = 0; arc_to_path_index < arc_to_path_commands.length; ++arc_to_path_index) {
                        var cur_arc_to_command = arc_to_path_commands[arc_to_path_index];
                        switch (cur_arc_to_command.id) {
                            case AscFormat.moveTo: {
                                lastX = cur_arc_to_command.X;
                                lastY = cur_arc_to_command.Y;
                                //  drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, cur_arc_to_command.X - dDist, cur_arc_to_command.Y - dDist,  2*dDist, 2*dDist, false, false);
                                break;
                            }
                            case AscFormat.bezier4: {

                                lastX = cur_arc_to_command.X2;
                                lastY = cur_arc_to_command.Y2;
                                //drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, cur_arc_to_command.X0 - dDist, cur_arc_to_command.Y0 - dDist,  2*dDist, 2*dDist, false, false);
                                //drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, cur_arc_to_command.X2 - dDist, cur_arc_to_command.Y2 - dDist,  2*dDist, 2*dDist, false, false);
                                break;
                            }
                        }
                    }
                    drawingDocument.DrawTrack(AscFormat.TYPE_TRACK.CHART_TEXT, transform, lastX - dDist, lastY - dDist, 2 * dDist, 2 * dDist, false, false, undefined, isDrawHandles);
                    i += 7;
                    break;
                }
                case close: {
                    i += 1;
                    break;
                }
            }
        }

    };
    Path2.prototype.getCommandByIndex = function (idx) {
        var i = 0;
        var path = this.PathMemory.ArrPathCommand;
        var len = path[this.startPos];
        var commandIndex = 0;
        while (i < len) {
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo: {
                    if (idx === commandIndex) {
                        return {id: moveTo, X: path[this.startPos + i + 2], Y: path[this.startPos + i + 3]};
                    }
                    i += 3;
                    break;
                }
                case lineTo: {
                    if (idx === commandIndex) {
                        return {id: moveTo, X: path[this.startPos + i + 2], Y: path[this.startPos + i + 3]};
                    }
                    i += 3;
                    break;
                }
                case bezier3: {
                    if (idx === commandIndex) {
                        return {
                            id: bezier3,
                            X0: path[this.startPos + i + 2],
                            Y0: path[this.startPos + i + 3],
                            X1: path[this.startPos + i + 4],
                            Y1: path[this.startPos + i + 5]
                        };
                    }
                    i += 5;
                    break;
                }
                case bezier4: {
                    if (idx === commandIndex) {
                        return {
                            id: bezier4,
                            X0: path[this.startPos + i + 2],
                            Y0: path[this.startPos + i + 3],
                            X1: path[this.startPos + i + 4],
                            Y1: path[this.startPos + i + 5],
                            X2: path[this.startPos + i + 6],
                            Y2: path[this.startPos + i + 7]
                        };
                    }
                    i += 7;
                    break;
                }
                case arcTo: {
                    if (idx === commandIndex) {
                        return {
                            id: arcTo,
                            stX: path[this.startPos + i + 2],
                            stY: path[this.startPos + i + 3],
                            wR: path[this.startPos + i + 4],
                            hR: path[this.startPos + i + 5],
                            stAng: path[this.startPos + i + 6],
                            swAng: path[this.startPos + i + 7]
                        };
                    }

                    i += 7;
                    break;
                }
                case close: {

                    if (idx === commandIndex) {
                        return {id: close};
                    }
                    i += 1;
                    break;
                }
            }
            ++commandIndex;
        }
        return null;
    };
    Path2.prototype.check_bounds = function (shape_drawer) {
        var path = this.ArrPathCommand;

        var i = 0;
        var len = this.PathMemory.ArrPathCommand[this.startPos];
        while (i < len) {
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo: {

                    shape_drawer._m(path[this.startPos + i + 2], path[this.startPos + i + 3]);
                    i += 3;
                    break;
                }
                case lineTo: {

                    shape_drawer._l(path[this.startPos + i + 2], path[this.startPos + i + 3]);
                    i += 3;
                    break;
                }
                case bezier3: {

                    shape_drawer._c2(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5]);
                    i += 5;
                    break;
                }
                case bezier4: {

                    shape_drawer._c(path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
                    i += 7;
                    break;
                }
                case arcTo: {

                    ArcToCurvers(shape_drawer, path[this.startPos + i + 2], path[this.startPos + i + 3], path[this.startPos + i + 4], path[this.startPos + i + 5], path[this.startPos + i + 6], path[this.startPos + i + 7]);
                    i += 7;
                    break;
                }
                case close: {
                    shape_drawer._z();
                    i += 1;
                    break;
                }
            }
        }

    };
    Path2.prototype.isSmartLine = function () {
        var i = 0;
        var path = this.PathMemory.ArrPathCommand;
        var len = path[this.startPos];
        var commandIndex = 0;
        while (i < len) {
            if (commandIndex > 1) {
                return false;
            }
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo: {
                    if (0 !== commandIndex) {
                        return false;
                    }
                    i += 3;
                    break;
                }
                case lineTo: {
                    if (1 !== commandIndex) {
                        return false;
                    }
                    i += 3;
                    break;
                }
                default: {
                    return false;
                }
            }
            ++commandIndex;
        }

        return true;
    };
    Path2.prototype.isSmartRect = function () {
        var i = 0;
        var path = this.PathMemory.ArrPathCommand;
        var len = path[this.startPos];
        var commandIndex = 0;
        var x0, y0, x1, y1, x2, y2, x3, y3, x4, y4, isCommand4Close = false;
        while (i < len) {
            if (commandIndex > 4) {
                return false;
            }
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo: {
                    if (0 !== commandIndex) {
                        return false;
                    }
                    x0 = path[this.startPos + i + 2];
                    y0 = path[this.startPos + i + 3];
                    i += 3;
                    break;
                }
                case lineTo: {
                    if (commandIndex === 1) {
                        x1 = path[this.startPos + i + 2];
                        y1 = path[this.startPos + i + 3];
                    } else if (commandIndex === 2) {
                        x2 = path[this.startPos + i + 2];
                        y2 = path[this.startPos + i + 3];
                    } else if (commandIndex === 3) {
                        x3 = path[this.startPos + i + 2];
                        y3 = path[this.startPos + i + 3];
                    } else if (commandIndex === 4) {
                        x4 = path[this.startPos + i + 2];
                        y4 = path[this.startPos + i + 3];
                    }
                    i += 3;
                    break;
                }
                case close: {
                    if (4 !== commandIndex) {
                        return false;
                    }
                    isCommand4Close = true;
                    break;
                }
                default: {
                    return false;
                }
            }
            ++commandIndex;
        }

        if (AscFormat.fApproxEqual(x0, x1)) {
            if (AscFormat.fApproxEqual(y1, y2)) {
                if (AscFormat.fApproxEqual(x2, x3) &&
                    AscFormat.fApproxEqual(y3, y0)) {
                    if (isCommand4Close)
                        return true;

                    if (AscFormat.fApproxEqual(x0, x4) &&
                        AscFormat.fApproxEqual(y0, y4)) {
                        return true;
                    }
                }
            }
        } else if (AscFormat.fApproxEqual(y0, y1)) {
            if (AscFormat.fApproxEqual(x1, x2)) {
                if (AscFormat.fApproxEqual(y2, y3) &&
                    AscFormat.fApproxEqual(x3, x0)) {
                    if (isCommand4Close)
                        return true;

                    if (AscFormat.fApproxEqual(x0, x4) &&
                        AscFormat.fApproxEqual(y0, y4)) {
                        return true;
                    }
                }
            }
        }

        return false;
    };
    Path2.prototype.drawSmart = function (shape_drawer) {
        var _graphics = shape_drawer.Graphics;
        var _full_trans = _graphics.m_oFullTransform;

        if (!_graphics || !_full_trans || undefined == _graphics.m_bIntegerGrid || true === shape_drawer.bIsNoSmartAttack)
            return this.draw(shape_drawer);

        var bIsTransformed = (_full_trans.shx == 0 && _full_trans.shy == 0) ? false : true;

        if (bIsTransformed)
            return this.draw(shape_drawer);

        var isLine = this.isSmartLine();
        var isRect = false;
        if (!isLine)
            isRect = this.isSmartRect();

        if (window["NATIVE_EDITOR_ENJINE"] || (!isLine && !isRect && !shape_drawer.bDrawSmartAttack))
            return this.draw(shape_drawer);

        var _old_int = _graphics.m_bIntegerGrid;

        if (false == _old_int)
            _graphics.SetIntegerGrid(true);

        var dKoefMMToPx = Math.max(_graphics.m_oCoordTransform.sx, 0.001);

        var _ctx = _graphics.m_oContext;
        var bIsStroke = (shape_drawer.bIsNoStrokeAttack || (this.stroke !== true)) ? false : true;
        var bIsEven = false;
        if (bIsStroke) {
            var _lineWidth = Math.max((shape_drawer.StrokeWidth * dKoefMMToPx + 0.5) >> 0, 1);
            _ctx.lineWidth = _lineWidth;

            if ((_lineWidth & 0x01) == 0x01)
                bIsEven = true;
        }

        var bIsDrawLast = false;
        var path = this.ArrPathCommand;
        shape_drawer._s();

        if (!isRect) {
            var i = 0;
            var len = this.PathMemory.ArrPathCommand[this.startPos];
            var X, Y;
            while (i < len) {
                var cmd = path[this.startPos + i + 1];
                switch (cmd) {
                    case moveTo: {
                        bIsDrawLast = true;
                        X = path[this.startPos + i + 2];
                        Y = path[this.startPos + i + 3];
                        var _x = (_full_trans.TransformPointX(X, Y)) >> 0;
                        var _y = (_full_trans.TransformPointY(X, Y)) >> 0;
                        if (bIsEven) {
                            _x -= 0.5;
                            _y -= 0.5;
                        }
                        _ctx.moveTo(_x, _y);

                        if (_graphics.ArrayPoints != null) {
                            _graphics.ArrayPoints.push({x: X, y: Y});
                        }

                        i += 3;
                        break;
                    }
                    case lineTo: {
                        bIsDrawLast = true;
                        X = path[this.startPos + i + 2];
                        Y = path[this.startPos + i + 3];
                        var _x = (_full_trans.TransformPointX(X, Y)) >> 0;
                        var _y = (_full_trans.TransformPointY(X, Y)) >> 0;
                        if (bIsEven) {
                            _x -= 0.5;
                            _y -= 0.5;
                        }
                        _ctx.lineTo(_x, _y);

                        if (_graphics.ArrayPoints != null) {
                            _graphics.ArrayPoints.push({x: X, y: Y});
                        }

                        i += 3;
                        break;
                    }
                    case bezier3: {
                        bIsDrawLast = true;

                        i += 5;
                        break;
                    }
                    case bezier4: {
                        bIsDrawLast = true;
                        i += 7;
                        break;
                    }
                    case arcTo: {
                        bIsDrawLast = true;
                        i += 7;
                        break;
                    }
                    case close: {
                        _ctx.closePath();
                        i += 1;
                        break;
                    }
                }
            }
        } else {
            var minX = 100000;
            var minY = 100000;
            var maxX = -100000;
            var maxY = -100000;
            bIsDrawLast = true;
            var i = 0;
            var len = this.PathMemory.ArrPathCommand[this.startPos], X, Y;
            while (i < len) {
                var cmd = path[this.startPos + i + 1];
                switch (cmd) {
                    case moveTo:
                    case lineTo: {
                        bIsDrawLast = true;
                        X = path[this.startPos + i + 2];
                        Y = path[this.startPos + i + 3];
                        if (minX > X)
                            minX = X;
                        if (minY > Y)
                            minY = Y;

                        if (maxX < X)
                            maxX = X;
                        if (maxY < Y)
                            maxY = Y;

                        i += 3;
                        break;
                    }
                    case bezier3: {
                        bIsDrawLast = true;
                        i += 5;
                        break;
                    }
                    case bezier4: {
                        bIsDrawLast = true;
                        i += 7;
                        break;
                    }
                    case arcTo: {
                        bIsDrawLast = true;
                        i += 7;
                        break;
                    }
                    case close: {
                        i += 1;
                        break;
                    }
                }
            }

            var _x1 = (_full_trans.TransformPointX(minX, minY)) >> 0;
            var _y1 = (_full_trans.TransformPointY(minX, minY)) >> 0;
            var _x2 = (_full_trans.TransformPointX(maxX, maxY)) >> 0;
            var _y2 = (_full_trans.TransformPointY(maxX, maxY)) >> 0;

            if (bIsEven)
                _ctx.rect(_x1 + 0.5, _y1 + 0.5, _x2 - _x1, _y2 - _y1);
            else
                _ctx.rect(_x1, _y1, _x2 - _x1, _y2 - _y1);
        }

        if (bIsDrawLast) {
            shape_drawer.isArrPix = true;
            shape_drawer.drawFillStroke(true, this.fill, bIsStroke);
            shape_drawer.isArrPix = false;
        }

        shape_drawer._e();

        if (false == _old_int)
            _graphics.SetIntegerGrid(false);
    };
    Path2.prototype.recalculate = function (gdLst, bResetPathsInfo) {
    };
    Path2.prototype.recalculate2 = function (gdLst, bResetPathsInfo) {
    };
    Path2.prototype.transformPointPolygon = function (x, y, oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds) {
        var oRet = {x: 0, y: 0}, y0, y1, cX, oPointOnPolygon, x1t, y1t, dX, dY, x0t, y0t;
        y0 = y;//dKoeff;
        if (bFlag) {
            y1 = 0;
            if (oBounds) {
                y0 -= oBounds.min_y;
            }
        } else {
            y1 = ContentHeight * dKoeff;
            if (oBounds) {
                y1 = (oBounds.max_y - oBounds.min_y);
                y0 -= oBounds.min_y;
            }
        }
        cX = x / XLimit;
        oPointOnPolygon = oPolygon.getPointOnPolygon(cX, true);
        x1t = oPointOnPolygon.x;
        y1t = oPointOnPolygon.y;
        dX = oPointOnPolygon.oP2.x - oPointOnPolygon.oP1.x;
        dY = oPointOnPolygon.oP2.y - oPointOnPolygon.oP1.y;

        if (bFlag) {
            dX = -dX;
            dY = -dY;
        }
        var dNorm = Math.sqrt(dX * dX + dY * dY);

        if (bFlag) {
            x0t = x1t + (dY / dNorm) * (y0);
            y0t = y1t - (dX / dNorm) * (y0);
        } else {

            x0t = x1t + (dY / dNorm) * (y1 - y0);
            y0t = y1t - (dX / dNorm) * (y1 - y0);
        }
        oRet.x = x0t;
        oRet.y = y0t;
        return oRet;
    };
    Path2.prototype.checkBetweenPolygons = function (oBoundsController, oPolygonWrapper1, oPolygonWrapper2) {

        var path = this.ArrPathCommand;
        var i = 0;
        var len = this.PathMemory.ArrPathCommand[this.startPos];
        var p;
        var dMinX = oBoundsController.min_x, dMinY = oBoundsController.min_y,
            dWidth = oBoundsController.max_x - oBoundsController.min_x,
            dHeight = oBoundsController.max_y - oBoundsController.min_y;
        while (i < len) {
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo:
                case lineTo: {
                    p = CheckPointByPaths(path[this.startPos + i + 2], path[this.startPos + i + 3], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
                    path[this.startPos + i + 2] = p.x;
                    path[this.startPos + i + 3] = p.y;
                    i += 3;
                    break;
                }
                case bezier3: {

                    p = CheckPointByPaths(path[this.startPos + i + 2], path[this.startPos + i + 3], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
                    path[this.startPos + i + 2] = p.x;
                    path[this.startPos + i + 3] = p.y;
                    p = CheckPointByPaths(path[this.startPos + i + 4], path[this.startPos + i + 5], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
                    path[this.startPos + i + 4] = p.x;
                    path[this.startPos + i + 5] = p.y;
                    i += 5;
                    break;
                }
                case bezier4: {
                    p = CheckPointByPaths(path[this.startPos + i + 2], path[this.startPos + i + 3], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
                    path[this.startPos + i + 2] = p.x;
                    path[this.startPos + i + 3] = p.y;
                    p = CheckPointByPaths(path[this.startPos + i + 4], path[this.startPos + i + 5], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
                    path[this.startPos + i + 4] = p.x;
                    path[this.startPos + i + 5] = p.y;
                    p = CheckPointByPaths(path[this.startPos + i + 6], path[this.startPos + i + 7], dWidth, dHeight, dMinX, dMinY, oPolygonWrapper1, oPolygonWrapper2);
                    path[this.startPos + i + 6] = p.x;
                    path[this.startPos + i + 7] = p.y;
                    i += 7;
                    break;
                }
                case arcTo: {
                    i += 7;
                    break;
                }
                case close: {
                    i += 1;
                    break;
                }
            }
        }
    };
    Path2.prototype.checkByPolygon = function (oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds) {

        var path = this.ArrPathCommand;
        var i = 0;
        var len = this.PathMemory.ArrPathCommand[this.startPos];
        var p;
        while (i < len) {
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo:
                case lineTo: {
                    p = this.transformPointPolygon(path[this.startPos + i + 2], path[this.startPos + i + 3], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
                    path[this.startPos + i + 2] = p.x;
                    path[this.startPos + i + 3] = p.y;
                    i += 3;
                    break;
                }
                case bezier3: {

                    p = this.transformPointPolygon(path[this.startPos + i + 2], path[this.startPos + i + 3], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
                    path[this.startPos + i + 2] = p.x;
                    path[this.startPos + i + 3] = p.y;
                    p = this.transformPointPolygon(path[this.startPos + i + 4], path[this.startPos + i + 5], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
                    path[this.startPos + i + 4] = p.x;
                    path[this.startPos + i + 5] = p.y;
                    i += 5;
                    break;
                }
                case bezier4: {
                    p = this.transformPointPolygon(path[this.startPos + i + 2], path[this.startPos + i + 3], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
                    path[this.startPos + i + 2] = p.x;
                    path[this.startPos + i + 3] = p.y;
                    p = this.transformPointPolygon(path[this.startPos + i + 4], path[this.startPos + i + 5], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
                    path[this.startPos + i + 4] = p.x;
                    path[this.startPos + i + 5] = p.y;
                    p = this.transformPointPolygon(path[this.startPos + i + 6], path[this.startPos + i + 7], oPolygon, bFlag, XLimit, ContentHeight, dKoeff, oBounds);
                    path[this.startPos + i + 6] = p.x;
                    path[this.startPos + i + 7] = p.y;
                    i += 7;
                    break;
                }
                case arcTo: {
                    i += 7;
                    break;
                }
                case close: {
                    i += 1;
                    break;
                }
            }
        }
    };
    Path2.prototype.transform = function (oTransform, dKoeff) {
        var path = this.ArrPathCommand;
        var i = 0;
        var len = this.PathMemory.ArrPathCommand[this.startPos];
        var p, x, y;
        while (i < len) {
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo:
                case lineTo: {
                    x = oTransform.TransformPointX(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
                    y = oTransform.TransformPointY(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
                    path[this.startPos + i + 2] = x;
                    path[this.startPos + i + 3] = y;
                    i += 3;
                    break;
                }
                case bezier3: {
                    x = oTransform.TransformPointX(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
                    y = oTransform.TransformPointY(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
                    path[this.startPos + i + 2] = x;
                    path[this.startPos + i + 3] = y;
                    x = oTransform.TransformPointX(path[this.startPos + i + 4] * dKoeff, path[this.startPos + i + 5] * dKoeff);
                    y = oTransform.TransformPointY(path[this.startPos + i + 4] * dKoeff, path[this.startPos + i + 5] * dKoeff);
                    path[this.startPos + i + 4] = x;
                    path[this.startPos + i + 5] = y;
                    i += 5;
                    break;
                }
                case bezier4: {
                    x = oTransform.TransformPointX(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
                    y = oTransform.TransformPointY(path[this.startPos + i + 2] * dKoeff, path[this.startPos + i + 3] * dKoeff);
                    path[this.startPos + i + 2] = x;
                    path[this.startPos + i + 3] = y;
                    x = oTransform.TransformPointX(path[this.startPos + i + 4] * dKoeff, path[this.startPos + i + 5] * dKoeff);
                    y = oTransform.TransformPointY(path[this.startPos + i + 4] * dKoeff, path[this.startPos + i + 5] * dKoeff);
                    path[this.startPos + i + 4] = x;
                    path[this.startPos + i + 5] = y;
                    x = oTransform.TransformPointX(path[this.startPos + i + 6] * dKoeff, path[this.startPos + i + 7] * dKoeff);
                    y = oTransform.TransformPointY(path[this.startPos + i + 6] * dKoeff, path[this.startPos + i + 7] * dKoeff);
                    path[this.startPos + i + 6] = x;
                    path[this.startPos + i + 7] = y;
                    i += 7;
                    break;
                }
                case arcTo: {
                    i += 7;
                    break;
                }
                case close: {
                    i += 1;
                    break;
                }
            }
        }
    };
    Path2.prototype.convertToBezierCurves = function (oPath, oTransform) {
        let dX0, dY0, dX1, dY1, dX2, dY2;
        let oLastMoveTo = null;
        let dLastX, dLastY;
        let path = this.ArrPathCommand;
        let i = 0;
        let len = path[this.startPos];
        if(len === 0) {
            return null;
        }
        let X, Y, X0, Y0, X1, Y1, X2, Y2, stX, stY, wR, hR, stAng, swAng;
        while (i < len) {
            var cmd = path[this.startPos + i + 1];
            switch (cmd) {
                case moveTo: {
                    X = path[this.startPos + i + 2];
                    Y = path[this.startPos + i + 3];
                    dX0 = oTransform.TransformPointX(X, Y)*36000 >> 0;
                    dY0 = oTransform.TransformPointY(X, Y)*36000 >> 0;
                    oPath.moveTo(dX0, dY0);
                    oLastMoveTo = i;
                    dLastX = X;
                    dLastY = Y;
                    i += 3;
                    break;
                }
                case lineTo: {
                    X = path[this.startPos + i + 2];
                    Y = path[this.startPos + i + 3];
                    dX0 = oTransform.TransformPointX(X, Y)*36000 >> 0;
                    dY0 = oTransform.TransformPointY(X, Y)*36000 >> 0;
                    dX0 = oTransform.TransformPointX(dLastX + (X - dLastX)*(1/3), dLastY + (Y - dLastY)*(1/3))*36000 >> 0;
                    dY0 = oTransform.TransformPointY(dLastX + (X - dLastX)*(1/3), dLastY + (Y - dLastY)*(1/3))*36000 >> 0;
                    dX1 = oTransform.TransformPointX(dLastX + (X - dLastX)*(2/3), dLastY + (Y - dLastY)*(2/3))*36000 >> 0;
                    dY1 = oTransform.TransformPointY(dLastX + (X - dLastX)*(2/3), dLastY + (Y - dLastY)*(2/3))*36000 >> 0;
                    dX2 = oTransform.TransformPointX(X, Y)*36000 >> 0;
                    dY2 = oTransform.TransformPointY(X, Y)*36000 >> 0;
                    oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
                    dLastX = X;
                    dLastY = Y;
                    i += 3;
                    break;
                }
                case bezier3: {
                    X0 = path[this.startPos + i + 2];
                    Y0 = path[this.startPos + i + 3];
                    X1 = path[this.startPos + i + 4];
                    Y1 = path[this.startPos + i + 5];
                    dX0 = oTransform.TransformPointX(X0, Y0)*36000 >> 0;
                    dY0 = oTransform.TransformPointY(X0, Y0)*36000 >> 0;
                    dX1 = oTransform.TransformPointX(X1, Y1)*36000 >> 0;
                    dY1 = oTransform.TransformPointY(X1, Y1)*36000 >> 0;
                    oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX1, dY1);
                    dLastX = X1;
                    dLastY = Y1;

                    i += 5;
                    break;
                }
                case bezier4: {

                    X0 = path[this.startPos + i + 2];
                    Y0 = path[this.startPos + i + 3];
                    X1 = path[this.startPos + i + 4];
                    Y1 = path[this.startPos + i + 5];
                    X2 = path[this.startPos + i + 6];
                    Y2 = path[this.startPos + i + 7];

                    dX0 = oTransform.TransformPointX(X0, Y0)*36000 >> 0;
                    dY0 = oTransform.TransformPointY(X0, Y0)*36000 >> 0;
                    dX1 = oTransform.TransformPointX(X1, Y1)*36000 >> 0;
                    dY1 = oTransform.TransformPointY(X1, Y1)*36000 >> 0;
                    dX2 = oTransform.TransformPointX(X2, Y2)*36000 >> 0;
                    dY2 = oTransform.TransformPointY(X2, Y2)*36000 >> 0;
                    oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
                    dLastX = X2;
                    dLastY = Y2;
                    i += 7;
                    break;
                }
                case arcTo: {
                    stX = path[this.startPos + i + 2];
                    stY = path[this.startPos + i + 3]
                    wR = path[this.startPos + i + 4];
                    hR = path[this.startPos + i + 5];
                    stAng = path[this.startPos + i + 6];
                    swAng = path[this.startPos + i + 7];
                    let oPathAccumulator = new AscFormat.PathAccumulator();
                    ArcToCurvers(oPathAccumulator, stX, stY, wR, hR, stAng, swAng);
                    let aArcToCommands = oPathAccumulator.pathCommand;
                    for(let nArcCmd = 0; nArcCmd < aArcToCommands.length; ++nArcCmd)  {
                        let oArcToCmd = aArcToCommands[nArcCmd];
                        switch (oArcToCmd.id) {
                            case AscFormat.moveTo: {
                                break;
                            }
                            case AscFormat.bezier4: {
                                dX0 = oTransform.TransformPointX(oArcToCmd.X0, oArcToCmd.Y0)*36000 >> 0;
                                dY0 = oTransform.TransformPointY(oArcToCmd.X0, oArcToCmd.Y0)*36000 >> 0;
                                dX1 = oTransform.TransformPointX(oArcToCmd.X1, oArcToCmd.Y1)*36000 >> 0;
                                dY1 = oTransform.TransformPointY(oArcToCmd.X1, oArcToCmd.Y1)*36000 >> 0;
                                dX2 = oTransform.TransformPointX(oArcToCmd.X2, oArcToCmd.Y2)*36000 >> 0;
                                dY2 = oTransform.TransformPointY(oArcToCmd.X2, oArcToCmd.Y2)*36000 >> 0;
                                oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);

                                dLastX = oArcToCmd.X2;
                                dLastY = oArcToCmd.Y2;
                                break;
                            }
                        }
                    }
                    i += 7;
                    break;
                }
                case close: {

                    if(AscFormat.isRealNumber(oLastMoveTo)) {
                        X = path[this.startPos + oLastMoveTo + 2];
                        Y = path[this.startPos + oLastMoveTo + 3];
                        let dXM = oTransform.TransformPointX(X, Y);
                        let dYM = oTransform.TransformPointY(X, Y);
                        let dLastXM = oTransform.TransformPointX(dLastX, dLastY);
                        let dLastYM = oTransform.TransformPointY(dLastX, dLastY);
                        dX0 = (dLastXM + (dXM - dLastXM) / 4) * 36000 >> 0;
                        dY0 = (dLastYM + (dYM - dLastYM) / 4) * 36000 >> 0;
                        dX1 = (dLastXM + 3 * (dXM - dLastXM) / 4) * 36000 >> 0;
                        dY1 = (dLastYM + 3 * (dYM - dLastYM) / 4) * 36000 >> 0;
                        dX2 = (dXM) * 36000 >> 0;
                        dY2 = (dYM) * 36000 >> 0;
                        oPath.cubicBezTo(dX0, dY0, dX1, dY1, dX2, dY2);
                    }
                    oPath.close();
                    i += 1;
                    break;
                }
            }
        }
        oPath.recalculate({}, true);
    };




function partition_bezier3(x0, y0, x1, y1, x2, y2, epsilon)
{
    var dx01 = x1 - x0;
    var dy01 = y1 - y0;
    var dx12 = x2 - x1;
    var dy12 = y2 - y1;

    var r01 = Math.sqrt(dx01*dx01 + dy01*dy01);
    var r12 = Math.sqrt(dx12*dx12 + dy12*dy12);
    if(Math.max(r01, r12) < epsilon)
    {
        return [{x: x0, y: y0}, {x: x1, y: y1}, {x: x2, y: y2}];
    }

    var x01 = (x0 + x1)*0.5;
    var y01 = (y0 + y1)*0.5;

    var x12 = (x1 + x2)*0.5;
    var y12 = (y1 + y2)*0.5;

    var x012 = (x01 + x12)*0.5;
    var y012 = (y01 + y12)*0.5;

    return  partition_bezier3(x0, y0, x01, y01, x012, y012, epsilon).concat(partition_bezier3(x012, y012, x12, y12, x2, y2, epsilon));
}

function partition_bezier4(x0, y0, x1, y1, x2, y2, x3, y3, epsilon)
{
    var dx01 = x1 - x0;
    var dy01 = y1 - y0;
    var dx12 = x2 - x1;
    var dy12 = y2 - y1;
    var dx23 = x3 - x2;
    var dy23 = y3 - y2;

    var r01 = Math.sqrt(dx01*dx01 + dy01*dy01);
    var r12 = Math.sqrt(dx12*dx12 + dy12*dy12);
    var r23 = Math.sqrt(dx23*dx23 + dy23*dy23);

    if(Math.max(r01, r12, r23) < epsilon)
        return [{x: x0, y: y0}, {x: x1, y: y1}, {x: x2, y: y2}, {x: x3, y: y3}];


    var x01 = (x0 + x1)*0.5;
    var y01 = (y0 + y1)*0.5;

    var x12 = (x1 + x2)*0.5;
    var y12 = (y1 + y2)*0.5;

    var x23 = (x2 + x3)*0.5;
    var y23 = (y2 + y3)*0.5;

    var x012 = (x01 + x12)*0.5;
    var y012 = (y01 + y12)*0.5;

    var x123 = (x12 + x23)*0.5;
    var y123 = (y12 + y23)*0.5;

    var x0123 = (x012 + x123)*0.5;
    var y0123 = (y012 + y123)*0.5;

    return partition_bezier4(x0, y0, x01, y01, x012, y012, x0123, y0123, epsilon).concat(partition_bezier4(x0123, y0123, x123, y123, x23, y23, x3, y3, epsilon));
}

    function splitBezier4(x0, y0, x1, y1, x2, y2, x3, y3, parameters) {
        const aResult = [[x0, y0, x1, y1, x2, y2, x3, y3]];
        if(!Array.isArray(parameters) || parameters.length === 0) {
            return aResult;
        }
        const aWorkingParameters =  [].concat(parameters);
        aWorkingParameters.sort(function (a, b) {return a - b});
        const isN = AscFormat.isRealNumber;
        const isE = AscFormat.fApproxEqual;
        let dLastParam = 1.0;
        for(let nParamIdx = aWorkingParameters.length - 1; nParamIdx > -1; --nParamIdx) {
            let dParam = aWorkingParameters[nParamIdx];
            if(!isN(dParam) || dParam <= 0 || dParam >= 1.0 || isE(dParam, dLastParam)) {
                aWorkingParameters.splice(nParamIdx, 1);
            }
            else {
                dLastParam = dParam;
            }
        }

        dLastParam = 1.0;
        for(let nParamIdx = aWorkingParameters.length - 1; nParamIdx > -1; --nParamIdx) {
            let dParam = aWorkingParameters[nParamIdx];
            let dWorkingParam = dParam/dLastParam;
            let oCurrentCurve = aResult[0];

            let x0c = oCurrentCurve[0];
            let y0c = oCurrentCurve[1];
            let x1c = oCurrentCurve[2];
            let y1c = oCurrentCurve[3];
            let x2c = oCurrentCurve[4];
            let y2c = oCurrentCurve[5];
            let x3c = oCurrentCurve[6];
            let y3c = oCurrentCurve[7];
            let t = dWorkingParam;

            //De Casteljau's algorithm
            let x01 = x0c + (x1c - x0c) * t;
            let y01 = y0c + (y1c - y0c) * t;
            let x12 = x1c + (x2c - x1c) * t;
            let y12 = y1c + (y2c - y1c) * t;
            let x23 = x2c + (x3c - x2c) * t;
            let y23 = y2c + (y3c - y2c) * t;
            let x0112 = x01 + (x12 - x01) * t;
            let y0112 = y01 + (y12 - y01) * t;
            let x1223 = x12 + (x23 - x12) * t;
            let y1223 = y12 + (y23 - y12) * t;
            let x01121223 = x0112 + (x1223 - x0112) * t;
            let y01121223 = y0112 + (y1223 - y0112) * t;
            let oCurve1 = [x0c, y0c, x01, y01,  x0112, y0112, x01121223, y01121223];
            let oCurve2 = [x01121223, y01121223, x1223, y1223, x23, y23, x3c, y3c];
            aResult.splice(0, 1, oCurve1, oCurve2);
            dLastParam = dParam;
        }
        return aResult;
    }
    function splitBezier4OnParts(x0, y0, x1, y1, x2, y2, x3, y3, nPartsCount) {
        if(!AscFormat.isRealNumber(nPartsCount) || nPartsCount < 2) {
            return splitBezier4(x0, y0, x1, y1, x2, y2, x3, y3, []);
        }
        let aParameters = [];
        const dDist = 1/nPartsCount;
        for(let nPartIdx = 1; nPartIdx < nPartsCount; ++nPartIdx) {
            aParameters.push(nPartIdx*dDist);
        }
        return splitBezier4(x0, y0, x1, y1, x2, y2, x3, y3, aParameters);
    }

    //--------------------------------------------------------export----------------------------------------------------
    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].moveTo = moveTo;
    window['AscFormat'].lineTo = lineTo;
    window['AscFormat'].arcTo = arcTo;
    window['AscFormat'].bezier3 = bezier3;
    window['AscFormat'].bezier4 = bezier4;
    window['AscFormat'].close = close;
    window['AscFormat'].cToRad2 = cToRad2;
    window['AscFormat'].Path = Path;
    window['AscFormat'].Path2 = Path2;
    window['AscFormat'].CPathCmd = CPathCmd;
    window['AscFormat'].partition_bezier3 = partition_bezier3;
    window['AscFormat'].partition_bezier4 = partition_bezier4;
    window['AscFormat'].splitBezier4 = splitBezier4;
    window['AscFormat'].splitBezier4OnParts = splitBezier4OnParts;
})(window);
