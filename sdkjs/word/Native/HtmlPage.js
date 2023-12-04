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

var Page_Width     = 210;
var Page_Height    = 297;

var X_Left_Margin   = 30;  // 3   cm
var X_Right_Margin  = 15;  // 1.5 cm
var Y_Bottom_Margin = 20;  // 2   cm
var Y_Top_Margin    = 20;  // 2   cm

var X_Right_Field  = Page_Width  - X_Right_Margin;
var Y_Bottom_Field = Page_Height - Y_Bottom_Margin;

var docpostype_Content     = 0x00;
var docpostype_HdrFtr      = 0x02;

var selectionflag_Common        = 0x00;
var selectionflag_Numbering     = 0x01;
var selectionflag_DrawingObject = 0x002;

var tableSpacingMinValue = 0.02;//0.02мм

function CEditorPage(api)
{
    this.Name = "";

    this.X = 0;
    this.Y = 0;
    this.Width      = 10;
    this.Height     = 10;

    this.m_oDrawingDocument = new AscCommon.CDrawingDocument();
    this.m_oLogicDocument   = null;

    this.m_oDrawingDocument.m_oWordControl = this;
    this.m_oDrawingDocument.m_oLogicDocument = this.m_oLogicDocument;
    this.m_oApi = api;

    this.Init = function()
    {
    };

    this.CheckRetinaDisplay = function()
    {
    };

    this.ShowOverlay = function()
    {
    };
    this.UnShowOverlay = function()
    {
    };
    this.CheckUnShowOverlay = function()
    {
    };
    this.CheckShowOverlay = function()
    {
    };

    this.initEvents = function()
    {
    };

    this.onButtonRulersClick = function()
    {
    };

    this.HideRulers = function()
    {
    };

    this.zoom_FitToWidth = function()
    {
    };
    this.zoom_FitToPage = function()
    {
    };

    this.zoom_Fire = function(type, old_zoom)
    {
    };

    this.zoom_Out = function()
    {
    };

    this.zoom_In = function()
    {
    };

    this.ToSearchResult = function()
    {
    };

    this.ScrollToPosition = function(x, y, PageNum, height)
    {
    };

    this.ScrollToAbsolutePosition = function(x, y, PageNum, isBottom)
    {
    };

    this.onButtonTabsClick = function()
    {
    };

    this.onPrevPage = function()
    {
    };
    this.onNextPage = function()
    {
    };

    this.horRulerMouseDown = function(e)
    {
    };
    this.horRulerMouseUp = function(e)
    {
    };
    this.horRulerMouseMove = function(e)
    {
    };

    this.verRulerMouseDown = function(e)
    {
    };
    this.verRulerMouseUp = function(e)
    {
    };
    this.verRulerMouseMove = function(e)
    {
    };

    this.SelectWheel = function()
    {
    };

    this.onMouseDown = function(e)
    {
    };

    this.onMouseMove = function(e)
    {
    };
    this.onMouseMove2 = function()
    {
    };
    this.onMouseUp = function(e, bIsWindow)
    {
    };

    this.onMouseUpMainSimple = function()
    {
    };

    this.onMouseUpExternal = function(x, y)
    {
    };

    this.onMouseWhell = function(e)
    {
    };

    this.checkViewerModeKeys = function(e)
    {
    };

    this.IncreaseReaderFontSize = function()
    {
    };
    this.DecreaseReaderFontSize = function()
    {
    };

    this.EnableReaderMode = function()
    {
    };

    this.DisableReaderMode = function()
    {
    };

    this.CheckDestroyReader = function()
    {
    };

    this.TransformDivUseAnimation = function(_div, topPos)
    {
    };

    this.onKeyDown = function(e)
    {
    };

    this.onKeyDownNoActiveControl = function(e)
    {
    };

    this.onKeyDownTBIM = function(e)
    {
    };

    this.DisableTextEATextboxAttack = function()
    {
    };

    this.onKeyUp = function(e)
    {
    };
    this.onKeyPress = function(e)
    {
    };

    this.verticalScroll = function(sender,scrollPositionY,maxY,isAtTop,isAtBottom)
    {
    };
    this.CorrectSpeedVerticalScroll = function(newScrollPos)
    {
    };

    this.horizontalScroll = function(sender,scrollPositionX,maxX,isAtLeft,isAtRight)
    {
    };

    this.UpdateScrolls = function()
    {
    };

    this.OnRePaintAttack = function()
    {
    };

    this.OnResize = function(isAttack)
    {
    };

    this.checkNeedRules = function()
    {
    };
    this.checkNeedHorScroll = function()
    {
    };

    this.getScrollMaxX = function(size)
    {
    };
    this.getScrollMaxY = function(size)
    {
    };

    this.StartUpdateOverlay = function()
    {
    };
    this.EndUpdateOverlay = function()
    {
    };

    this.OnUpdateOverlay = function()
    {
    };

    this.OnUpdateSelection = function()
    {
    };

    this.OnCalculatePagesPlace = function()
    {
    };

    this.OnPaint = function()
    {
    };

    this.CheckRetinaElement = function(htmlElem)
    {
    };

    this.GetDrawingPageInfo = function(nPageIndex)
    {
    };

    this.CheckFontCache = function()
    {
    };
    this.OnScroll = function()
    {
    };

    this.CheckZoom = function()
    {
    };

    this.ChangeHintProps = function()
    {
    };

    this.CalculateDocumentSize = function()
    {
    };

    this.InitDocument = function(bIsEmpty)
    {
    };

    this.InitControl = function()
    {
    };

    this.OpenDocument = function(info)
    {
    };

    this.AnimationFrame = function()
    {
    };

    this.onTimerScroll = function()
    {
            var oWordControl = editor.WordControl;
            if(oWordControl.m_oLogicDocument)
            {
                oWordControl.m_oLogicDocument.ContinueSpellCheck();
            }
            oWordControl.m_nPaintTimerId = setTimeout(oWordControl.onTimerScroll, 500);
    };

    this.StartMainTimer = function()
    {
        this.onTimerScroll();
    };

    this.onTimerScroll2 = function()
    {
    };

    this.onTimerScroll2_sync = function()
    {
    };

    this.UpdateHorRuler = function()
    {
    };
    this.UpdateVerRuler = function()
    {
    };

    this.SetCurrentPage = function(isNoUpdateRulers)
    {
    };
    this.SetCurrentPage2 = function()
    {
    };

    this.UpdateHorRulerBack = function()
    {
    };
    this.UpdateVerRulerBack = function()
    {
    };

    this.GoToPage = function(lPageNum)
    {
    };

    this.GetVerticalScrollTo = function(y, page)
    {
    };

    this.GetHorizontalScrollTo = function(x, page)
    {
    };

    this.ReinitTB = function()
    {
    };

    this.SetTextBoxMode = function(isEA)
    {
    };

    this.TextBoxFocus = function()
    {
    };

    this.OnTextBoxInput = function()
    {
    };

    this.CheckTextBoxSize = function()
    {
    };

    this.TextBoxOnKeyDown = function(e)
    {
    };

    this.onChangeTB = function()
    {
    };
    this.CheckTextBoxInputPos = function()
    {
    };

    this.checkMouseHandMode = function()
    {
    };

    this.ReaderModeCurrent = 0;
    this.ReaderFontSizeCur = 2;
    this.ReaderFontSizes   = [12, 14, 16, 18, 22, 28, 36, 48, 72];
    this.ChangeReaderMode = function()
    {
        if (!this.m_oLogicDocument)
            return;

        if (this.ReaderModeCurrent)
        {
            this.m_oLogicDocument.SetDocumentPrintMode();
            this.ReaderModeCurrent = 0;
        }
        else
        {
            this.SetNewMobileMode();
            this.ReaderModeCurrent = 1;
        }
    };

    this.SetNewMobileMode = function()
    {
        if (this.m_oLogicDocument)
        {
            let sectPr = this.m_oLogicDocument.GetSectionsInfo().Get(0).SectPr;
            let scale = 2; // TODO: get deviceScale
            const nPageW = sectPr.GetPageWidth() / scale;
            const nPageH = sectPr.GetPageHeight() / scale;
            const nScale = this.ReaderFontSizes[this.ReaderFontSizeCur] / 16;
            this.m_oLogicDocument.SetDocumentReadMode(nPageW, nPageH, nScale);
            return true;
        }
        return false;
    };

    this.IncreaseReaderFontSize = function()
    {
        if (this.ReaderFontSizeCur >= (this.ReaderFontSizes.length - 1))
        {
            this.ReaderFontSizeCur = this.ReaderFontSizes.length - 1;
            return false;
        }
        this.ReaderFontSizeCur++;
        return this.SetNewMobileMode();
    };
    this.DecreaseReaderFontSize = function()
    {
        if (this.ReaderFontSizeCur <= 0)
        {
            this.ReaderFontSizeCur = 0;
            return false;
        }
        this.ReaderFontSizeCur--;
        return this.SetNewMobileMode();
    };
}

//------------------------------------------------------------export----------------------------------------------------
window['AscCommon'] = window['AscCommon'] || {};
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].CEditorPage = CEditorPage;

window['AscCommon'].Page_Width = Page_Width;
window['AscCommon'].Page_Height = Page_Height;
window['AscCommon'].X_Left_Margin = X_Left_Margin;
window['AscCommon'].X_Right_Margin = X_Right_Margin;
window['AscCommon'].Y_Bottom_Margin = Y_Bottom_Margin;
window['AscCommon'].Y_Top_Margin = Y_Top_Margin;
