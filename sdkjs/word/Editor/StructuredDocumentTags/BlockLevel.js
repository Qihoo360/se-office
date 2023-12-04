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

var type_BlockLevelSdt = 0x0003;

var c_oAscSdtLockType = Asc.c_oAscSdtLockType;

/**
 * @param oParent - родительский класс
 * @param oLogicDocument {CDocument} - главный класс документа
 * @constructor
 * @extends {CDocumentContentElementBase}
 */
function CBlockLevelSdt(oLogicDocument, oParent)
{
	CDocumentContentElementBase.call(this, oParent);

	this.LogicDocument = oLogicDocument;
	this.Content       = new CDocumentContent(this, oLogicDocument ? oLogicDocument.Get_DrawingDocument() : null, 0, 0, 0, 0, true, false, false);
	this.Pr            = new CSdtPr();
	this.Lock          = new AscCommon.CLock();

	// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
	AscCommon.g_oTableId.Add(this, this.Id);

	this.SkipSpecialLock = false;
	this.Current         = false;
}

CBlockLevelSdt.prototype = Object.create(CDocumentContentElementBase.prototype);
CBlockLevelSdt.prototype.constructor = CBlockLevelSdt;
AscCommon.ExtendPrototype(CBlockLevelSdt, CSdtBase);

CBlockLevelSdt.prototype.IsInlineLevel = function()
{
	return false;
};
CBlockLevelSdt.prototype.IsBlockLevel = function()
{
	return true;
};
CBlockLevelSdt.prototype.Copy = function(Parent, DrawingDocument, oPr)
{
	var oNew = new CBlockLevelSdt(this.LogicDocument, Parent ? Parent : this.Parent);

	oNew.Content.Copy2(this.Content, oPr);
	oNew.SetPr(this.Pr);

	if (undefined !== this.Pr.CheckBox)
		oNew.SetCheckBoxPr(this.Pr.CheckBox);

	if (undefined !== this.Pr.Picture)
		oNew.SetPicturePr(this.Pr.Picture);

	if (undefined !== this.Pr.ComboBox)
		oNew.SetComboBoxPr(this.Pr.ComboBox);

	if (undefined !== this.Pr.DropDown)
		oNew.SetDropDownListPr(this.Pr.DropDown);

	if (undefined !== this.Pr.Date)
		oNew.SetDatePickerPr(this.Pr.Date);

	if (oNew.IsEmpty())
		oNew.private_ReplaceContentWithPlaceHolder();

	oNew.SetShowingPlcHdr(this.Pr.ShowingPlcHdr);
	oNew.SetPlaceholder(this.private_CopyPlaceholder(oPr));
	oNew.SetContentControlEquation(this.Pr.Equation);
	oNew.SetContentControlTemporary(this.Pr.Temporary);

	return oNew;
};
CBlockLevelSdt.prototype.GetType = function()
{
	return type_BlockLevelSdt;
};
/**
 * Получаем содержимое данного контейнера
 * @returns {CDocumentContent}
 */
CBlockLevelSdt.prototype.GetContent = function()
{
	return this.Content;
};
CBlockLevelSdt.prototype.Is_Inline = function()
{
	return true;
};
CBlockLevelSdt.prototype.IsInline = function()
{
	return true;
};
CBlockLevelSdt.prototype.Reset = function(X, Y, XLimit, YLimit, PageAbs, ColumnAbs, ColumnsCount)
{
	this.Content.Reset(X, Y, XLimit, YLimit);
	this.Content.Set_StartPage(0);

	this.X            = X;
	this.Y            = Y;
	this.XLimit       = XLimit;
	this.YLimit       = YLimit;
	this.PageNum      = PageAbs;
	this.ColumnNum    = ColumnAbs ? ColumnAbs : 0;
	this.ColumnsCount = ColumnsCount ? ColumnsCount : 1;
};
CBlockLevelSdt.prototype.Recalculate_Page = function(CurPage)
{
	this.SetIsRecalculated(true);

	this.Content.RecalcInfo = this.Parent.RecalcInfo;

	var RecalcResult = this.Content.Recalculate_Page(CurPage, true);

	if (recalcresult2_End === RecalcResult && window['AscCommon'].g_specialPasteHelper && window['AscCommon'].g_specialPasteHelper.showButtonIdParagraph === this.GetId())
		window['AscCommon'].g_specialPasteHelper.SpecialPasteButtonById_Show();

	if (recalcresult2_End === RecalcResult)
		return recalcresult_NextElement;
	else if (recalcresult2_NextPage === RecalcResult)
		return recalcresult_NextPage;
	else if (recalcresult2_CurPage === RecalcResult)
		return recalcresult_CurPage;
};
CBlockLevelSdt.prototype.Get_PageBounds = function(CurPage)
{
	return this.Content.Get_PageBounds(CurPage);
};
CBlockLevelSdt.prototype.GetContentBounds = function(CurPage)
{
	return this.Content.GetContentBounds(CurPage);
};
CBlockLevelSdt.prototype.IsEmptyPage = function(nCurPage)
{
	return this.Content.IsEmptyPage(nCurPage);
};
CBlockLevelSdt.prototype.Get_PagesCount = function()
{
	return this.Content.Get_PagesCount();
};
CBlockLevelSdt.prototype.Reset_RecalculateCache = function()
{
	this.Content.Reset_RecalculateCache();
};
CBlockLevelSdt.prototype.Write_ToBinary2 = function(Writer)
{
	Writer.WriteLong(AscDFH.historyitem_type_BlockLevelSdt);
	// String : Id
	// String : Content id
	Writer.WriteString2(this.GetId());
	Writer.WriteString2(this.Content.GetId());
};
CBlockLevelSdt.prototype.Read_FromBinary2 = function(Reader)
{
	this.LogicDocument = editor.WordControl.m_oLogicDocument;

	// String : Id
	// String : Content id
	this.Id          = Reader.GetString2();
	this.Content     = this.LogicDocument.Get_TableId().Get_ById(Reader.GetString2());
};
CBlockLevelSdt.prototype.Draw = function(CurPage, oGraphics)
{
	if (this.LogicDocument.GetSdtGlobalShowHighlight() && undefined === oGraphics.RENDERER_PDF_FLAG)
	{
		var oBounds = this.GetContentBounds(CurPage);
		var oColor  = this.LogicDocument.GetSdtGlobalColor();

		oGraphics.b_color1(oColor.r, oColor.g, oColor.b, 255);
		oGraphics.rect(oBounds.Left, oBounds.Top, oBounds.Right - oBounds.Left, oBounds.Bottom - oBounds.Top);
		oGraphics.df();
	}

	var isPlaceHolder = this.IsPlaceHolder();
	var nTextAlpha;
	if (isPlaceHolder && oGraphics.setTextGlobalAlpha)
	{
		nTextAlpha = oGraphics.getTextGlobalAlpha();
		oGraphics.setTextGlobalAlpha(0.5);
	}

	this.Content.Draw(CurPage, oGraphics);

	if (isPlaceHolder && oGraphics.setTextGlobalAlpha)
		oGraphics.setTextGlobalAlpha(nTextAlpha);

	if (AscCommon.locktype_None !== this.Lock.Get_Type())
	{
		var oBounds = this.GetContentBounds(CurPage);
		oGraphics.DrawLockObjectRect(this.Lock.Get_Type(), oBounds.Left, oBounds.Top, oBounds.Right - oBounds.Left, oBounds.Bottom - oBounds.Top);
	}
};
CBlockLevelSdt.prototype.Get_CurrentPage_Absolute = function()
{
	return this.Content.Get_CurrentPage_Absolute();
};
CBlockLevelSdt.prototype.Get_CurrentPage_Relative = function()
{
	return this.Content.Get_CurrentPage_Relative();
};
CBlockLevelSdt.prototype.IsInText = function(X, Y, CurPage)
{
	return this.Content.IsInText(X, Y, CurPage);
};
CBlockLevelSdt.prototype.IsInDrawing = function(X, Y, CurPage)
{
	return this.Content.IsInDrawing(X, Y, CurPage);
};
CBlockLevelSdt.prototype.IsTableBorder = function(X, Y, CurPage)
{
	return this.Content.IsTableBorder(X, Y, CurPage);
};
CBlockLevelSdt.prototype.UpdateCursorType = function(X, Y, CurPage)
{
	var oBounds = this.GetContentBounds(CurPage);
	if (true === this.Lock.Is_Locked() && X < oBounds.Right && X > oBounds.Left && Y > oBounds.Top && Y < oBounds.Bottom)
	{
		var MMData              = new AscCommon.CMouseMoveData();
		var Coords              = this.LogicDocument.DrawingDocument.ConvertCoordsToCursorWR(oBounds.Left, oBounds.Top, this.Get_AbsolutePage(CurPage), this.Get_ParentTextTransform());
		MMData.X_abs            = Coords.X - 5;
		MMData.Y_abs            = Coords.Y;
		MMData.Type             = Asc.c_oAscMouseMoveDataTypes.LockedObject;
		MMData.UserId           = this.Lock.Get_UserId();
		MMData.HaveChanges      = this.Lock.Have_Changes();
		MMData.LockedObjectType = c_oAscMouseMoveLockedObjectType.Common;
		this.LogicDocument.Api.sync_MouseMoveCallback(MMData);
	}

	this.DrawContentControlsTrack(AscCommon.ContentControlTrack.Hover, X, Y, CurPage);
	return this.Content.UpdateCursorType(X, Y, CurPage);
};
CBlockLevelSdt.prototype.Selection_SetStart = function(X, Y, CurPage, MouseEvent, isTableBorder)
{
	if (this.IsPlaceHolder())
	{
		var nDirection = this.Parent && this.Parent.GetSelectDirection ? this.Parent.GetSelectDirection() : 1;
		this.SelectAll(nDirection);
		return;
	}

	this.Content.Selection_SetStart(X, Y, CurPage, MouseEvent, isTableBorder);
};
CBlockLevelSdt.prototype.Selection_SetEnd = function(X, Y, CurPage, MouseEvent, isTableBorder)
{
	if (this.IsPlaceHolder())
	{
		var nDirection = this.Parent && this.Parent.GetSelectDirection ? this.Parent.GetSelectDirection() : 1;
		this.SelectAll(nDirection);
		return;
	}

	this.Content.Selection_SetEnd(X, Y, CurPage, MouseEvent, isTableBorder);
};
CBlockLevelSdt.prototype.IsSelectionEmpty = function(isCheckHidden)
{
	return this.Content.IsSelectionEmpty(isCheckHidden);
};
CBlockLevelSdt.prototype.GetSelectedElementsInfo = function(oInfo)
{
	if (!oInfo.IsSkipTOC() || !this.IsBuiltInTableOfContents())
		oInfo.SetBlockLevelSdt(this);

	this.Content.GetSelectedElementsInfo(oInfo);
};
CBlockLevelSdt.prototype.IsSelectionUse = function()
{
	return this.Content.IsSelectionUse();
};
CBlockLevelSdt.prototype.IsSelectionToEnd = function()
{
	return this.Content.IsSelectionToEnd();
};
CBlockLevelSdt.prototype.RemoveSelection = function()
{
	this.Content.RemoveSelection();
};
CBlockLevelSdt.prototype.SetSelectionUse = function(isUse)
{
	this.Content.SetSelectionUse(isUse);
};
CBlockLevelSdt.prototype.SetSelectionToBeginEnd = function(isSelectionStart, isElementStart)
{
	this.Content.SetSelectionToBeginEnd(isSelectionStart, isElementStart);
};
CBlockLevelSdt.prototype.SelectAll = function(nDirection)
{
	this.Content.SelectAll(nDirection);
};
CBlockLevelSdt.prototype.GetCalculatedTextPr = function()
{
	return this.Content.GetCalculatedTextPr(true);
};
CBlockLevelSdt.prototype.GetCalculatedParaPr = function()
{
	return this.Content.GetCalculatedParaPr();
};
CBlockLevelSdt.prototype.GetDirectParaPr = function()
{
	return this.Content.GetDirectParaPr();
};
CBlockLevelSdt.prototype.GetDirectTextPr = function()
{
	return this.Content.GetDirectTextPr();
};
CBlockLevelSdt.prototype.DrawSelectionOnPage = function(CurPage)
{
	this.Content.DrawSelectionOnPage(CurPage);
};
CBlockLevelSdt.prototype.GetSelectionBounds = function()
{
	return this.Content.GetSelectionBounds();
};
CBlockLevelSdt.prototype.RecalculateCurPos = function(bUpdateX, bUpdateY, isUpdateTarget)
{
	return this.Content.RecalculateCurPos(bUpdateX, bUpdateY, isUpdateTarget);
};
CBlockLevelSdt.prototype.Can_CopyCut = function()
{
	return this.Content.Can_CopyCut();
};
CBlockLevelSdt.prototype.CheckPosInSelection = function(X, Y, CurPage, NearPos)
{
	return this.Content.CheckPosInSelection(X, Y, CurPage, NearPos);
};
CBlockLevelSdt.prototype.Get_NearestPos = function(CurPage, X, Y, bAnchor, Drawing)
{
	return this.Content.Get_NearestPos(CurPage, X, Y, bAnchor, Drawing);
};
CBlockLevelSdt.prototype.CanUpdateTarget = function(CurPage)
{
	return this.Content.CanUpdateTarget(CurPage);
};
CBlockLevelSdt.prototype.MoveCursorLeft = function(AddToSelect, Word)
{
	if (this.IsPlaceHolder())
	{
		if (AddToSelect)
			this.SelectAll(-1);

		return false;
	}

	var bResult = this.Content.MoveCursorLeft(AddToSelect, Word);
	if (!bResult && this.LogicDocument.IsFillingFormMode())
		return true;

	return bResult;
};
CBlockLevelSdt.prototype.MoveCursorLeftWithSelectionFromEnd = function(Word)
{
	return this.Content.MoveCursorLeftWithSelectionFromEnd(Word);
};
CBlockLevelSdt.prototype.MoveCursorRight = function(AddToSelect, Word)
{
	if (this.IsPlaceHolder())
	{
		if (AddToSelect)
			this.SelectAll(1);

		return false;
	}

	var bResult = this.Content.MoveCursorRight(AddToSelect, Word, false);
	if (!bResult && this.LogicDocument.IsFillingFormMode())
		return true;

	return bResult;
};
CBlockLevelSdt.prototype.MoveCursorRightWithSelectionFromStart = function(Word)
{
	return this.Content.MoveCursorRightWithSelectionFromStart(Word);
};
CBlockLevelSdt.prototype.MoveCursorToStartPos = function(AddToSelect)
{
	return this.Content.MoveCursorToStartPos(AddToSelect);
};
CBlockLevelSdt.prototype.MoveCursorToEndPos = function(AddToSelect, StartSelectFromEnd)
{
	return this.Content.MoveCursorToEndPos(AddToSelect, StartSelectFromEnd);
};
CBlockLevelSdt.prototype.MoveCursorUp = function(AddToSelect)
{
	return this.Content.MoveCursorUp(AddToSelect);
};
CBlockLevelSdt.prototype.MoveCursorUpToLastRow = function(X, Y, AddToSelect)
{
	return this.Content.MoveCursorUpToLastRow(X, Y, AddToSelect);
};
CBlockLevelSdt.prototype.MoveCursorDown = function(AddToSelect)
{
	return this.Content.MoveCursorDown(AddToSelect);
};
CBlockLevelSdt.prototype.MoveCursorDownToFirstRow = function(X, Y, AddToSelect)
{
	return this.Content.MoveCursorDownToFirstRow(X, Y, AddToSelect);
};
CBlockLevelSdt.prototype.MoveCursorToEndOfLine = function(AddToSelect)
{
	return this.Content.MoveCursorToEndOfLine(AddToSelect);
};
CBlockLevelSdt.prototype.MoveCursorToStartOfLine = function(AddToSelect)
{
	return this.Content.MoveCursorToStartOfLine(AddToSelect);
};
CBlockLevelSdt.prototype.MoveCursorToXY = function(X, Y, bLine, bDontChangeRealPos, CurPage)
{
	return this.Content.MoveCursorToXY(X, Y, bLine, bDontChangeRealPos, CurPage);
};
CBlockLevelSdt.prototype.MoveCursorToCell = function(bNext)
{
	return this.Content.MoveCursorToCell(bNext);
};
CBlockLevelSdt.prototype.GetSelectionState = function()
{
	return this.Content.GetSelectionState();
};
CBlockLevelSdt.prototype.SetSelectionState = function(State, StateIndex)
{
	return this.Content.SetSelectionState(State, StateIndex);
};
CBlockLevelSdt.prototype.IsCursorAtBegin = function(bOnlyPara)
{
	return this.Content.IsCursorAtBegin(bOnlyPara);
};
CBlockLevelSdt.prototype.IsCursorAtEnd = function()
{
	return this.Content.IsCursorAtEnd();
};
CBlockLevelSdt.prototype.AddNewParagraph = function()
{
	this.private_ReplacePlaceHolderWithContent();
	return this.Content.AddNewParagraph();
};
CBlockLevelSdt.prototype.Get_SelectionState2 = function()
{
	var oState  = new CDocumentSelectionState();
	oState.Id   = this.GetId();
	oState.Data = this.Content.Get_SelectionState2();
	return oState;
};
CBlockLevelSdt.prototype.Set_SelectionState2 = function(State)
{
	if (State.Data)
		this.Content.Set_SelectionState2(State.Data);
};
CBlockLevelSdt.prototype.IsStartFromNewPage = function()
{
	this.Content.IsStartFromNewPage();
};
CBlockLevelSdt.prototype.GetAllParagraphs = function(Props, ParaArray)
{
	return this.Content.GetAllParagraphs(Props, ParaArray);
};
CBlockLevelSdt.prototype.GetAllTables = function(oProps, arrTables)
{
	return this.Content.GetAllTables(oProps, arrTables);
};
CBlockLevelSdt.prototype.SetContentSelection = function(StartDocPos, EndDocPos, Depth, StartFlag, EndFlag)
{
	this.Content.SetContentSelection(StartDocPos, EndDocPos, Depth, StartFlag, EndFlag);
};
CBlockLevelSdt.prototype.GetContentPosition = function(bSelection, bStart, PosArray)
{
	return this.Content.GetContentPosition(bSelection, bStart, PosArray);
};
CBlockLevelSdt.prototype.SetContentPosition = function(DocPos, Depth, Flag)
{
	this.Content.SetContentPosition(DocPos, Depth, Flag);
};
CBlockLevelSdt.prototype.GetNumberingInfo = function(oNumberingEngine)
{
	return this.Content.GetNumberingInfo(oNumberingEngine);
};
CBlockLevelSdt.prototype.AddInlineImage = function(W, H, Img, GraphicObject, bFlow)
{
	this.private_ReplacePlaceHolderWithContent();
	this.Content.AddInlineImage(W, H, Img, GraphicObject, bFlow);
};
CBlockLevelSdt.prototype.AddImages = function(aImages)
{
	this.private_ReplacePlaceHolderWithContent();
	this.Content.AddImages(aImages);
};
CBlockLevelSdt.prototype.AddSignatureLine = function(oSignatureDrawing)
{
	this.private_ReplacePlaceHolderWithContent();
	this.Content.AddSignatureLine(oSignatureDrawing);
};
CBlockLevelSdt.prototype.AddOleObject = function(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect, arrImagesForAddToHistory)
{
	this.private_ReplacePlaceHolderWithContent();
	return this.Content.AddOleObject(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect, arrImagesForAddToHistory);
};
CBlockLevelSdt.prototype.AddTextArt = function(nStyle)
{
	this.private_ReplacePlaceHolderWithContent();
	this.Content.AddTextArt(nStyle);
};
CBlockLevelSdt.prototype.AddInlineTable = function(nCols, nRows, nMode)
{
	this.private_ReplacePlaceHolderWithContent();
	return this.Content.AddInlineTable(nCols, nRows, nMode);
};
CBlockLevelSdt.prototype.Remove = function(nCount, isRemoveWholeElement, bRemoveOnlySelection, bOnAddText, isWord)
{
	if (this.IsPlaceHolder())
	{
		let logicDocument = this.GetLogicDocument();
		
		if (!this.CanBeDeleted() && !bOnAddText)
			return true;
		
		if (bOnAddText || !(logicDocument && logicDocument.IsDocumentEditor() && logicDocument.IsFillingFormMode()))
			this.private_ReplacePlaceHolderWithContent();
		
		return !!bOnAddText;
	}

	var bResult = this.Content.Remove(nCount, isRemoveWholeElement, bRemoveOnlySelection, bOnAddText, isWord);

	if (bResult && this.IsContentControlTemporary())
	{
		this.RemoveContentControlWrapper();
	}
	else if (this.IsEmpty()
		&& !bOnAddText
		&& (true !== isRemoveWholeElement || bResult)
		&& this.CanBeEdited())
	{
		this.private_ReplaceContentWithPlaceHolder();
		return true;
	}

	return bResult;
};
CBlockLevelSdt.prototype.Is_Empty = function()
{
	return this.Content.Is_Empty();
};
CBlockLevelSdt.prototype.Add = function(oParaItem)
{
	var isRemoveWrapper = false;
	if (oParaItem && oParaItem.Type !== para_TextPr)
	{
		isRemoveWrapper = this.IsContentControlTemporary();
		this.private_ReplacePlaceHolderWithContent(true);
	}
	else if (oParaItem && oParaItem.Type !== para_TextPr && this.IsPlaceHolder())
	{
		var oTempTextPr = this.Pr.TextPr.Copy();
		oTempTextPr.Merge(oParaItem.Value);
		this.SetDefaultTextPr(oTempTextPr);
	}

	if (oParaItem && para_TextPr === oParaItem.Type && (this.IsComboBox() || this.IsDropDownList()))
	{
		this.Content.SetApplyToAll(true);
		this.Content.AddToParagraph(oParaItem);
		this.Content.SetApplyToAll(false);
	}
	else
	{
		this.Content.AddToParagraph(oParaItem);
	}

	if (isRemoveWrapper)
		this.RemoveContentControlWrapper();
};
CBlockLevelSdt.prototype.PreDelete = function()
{
	if (this.IsPlaceHolder())
		return;

	this.Content.PreDelete();
};
CBlockLevelSdt.prototype.ClearParagraphFormatting = function(isClearParaPr, isClearTextPr)
{
	if (this.IsPlaceHolder())
		return;

	this.Content.ClearParagraphFormatting(isClearParaPr, isClearTextPr);
};
CBlockLevelSdt.prototype.GetCursorPosXY = function()
{
	return this.Content.GetCursorPosXY();
};
CBlockLevelSdt.prototype.StartSelectionFromCurPos = function()
{
	this.Content.StartSelectionFromCurPos();
};
CBlockLevelSdt.prototype.SetParagraphPr = function(oParaPr)
{
	return this.Content.SetParagraphPr(oParaPr);
};
CBlockLevelSdt.prototype.SetParagraphAlign = function(Align)
{
	return this.Content.SetParagraphAlign(Align);
};
CBlockLevelSdt.prototype.SetParagraphSpacing = function(Spacing)
{
	return this.Content.SetParagraphSpacing(Spacing);
};
CBlockLevelSdt.prototype.SetParagraphTabs = function(Tabs)
{
	return this.Content.SetParagraphTabs(Tabs);
};
CBlockLevelSdt.prototype.SetParagraphIndent = function(Ind)
{
	return this.Content.SetParagraphIndent(Ind);
};
CBlockLevelSdt.prototype.SetParagraphShd = function(Shd)
{
	return this.Content.SetParagraphShd(Shd);
};
CBlockLevelSdt.prototype.SetParagraphStyle = function(Name)
{
	return this.Content.SetParagraphStyle(Name);
};
CBlockLevelSdt.prototype.SetParagraphContextualSpacing = function(Value)
{
	return this.Content.SetParagraphContextualSpacing(Value);
};
CBlockLevelSdt.prototype.SetParagraphPageBreakBefore = function(Value)
{
	return this.Content.SetParagraphPageBreakBefore(Value);
};
CBlockLevelSdt.prototype.SetParagraphKeepLines = function(Value)
{
	return this.Content.SetParagraphKeepLines(Value);
};
CBlockLevelSdt.prototype.SetParagraphKeepNext = function(Value)
{
	return this.Content.SetParagraphKeepNext(Value);
};
CBlockLevelSdt.prototype.SetParagraphWidowControl = function(Value)
{
	return this.Content.SetParagraphWidowControl(Value);
};
CBlockLevelSdt.prototype.SetParagraphBorders = function(Borders)
{
	return this.Content.SetParagraphBorders(Borders);
};
CBlockLevelSdt.prototype.SetParagraphFramePr = function(FramePr, bDelete)
{
	return this.Content.SetParagraphFramePr(FramePr, bDelete);
};
CBlockLevelSdt.prototype.IncreaseDecreaseFontSize = function(bIncrease)
{
	return this.Content.IncreaseDecreaseFontSize(bIncrease);
};
CBlockLevelSdt.prototype.IncreaseDecreaseIndent = function(bIncrease)
{
	return this.Content.IncreaseDecreaseIndent(bIncrease);
};
CBlockLevelSdt.prototype.SetImageProps = function(oProps)
{
	return this.Content.SetImageProps(oProps);
};
CBlockLevelSdt.prototype.SetTableProps = function(oProps)
{
	return this.Content.SetTableProps(oProps);
};
CBlockLevelSdt.prototype.GetSelectedContent = function(oSelectedContent)
{
	if (this.Content.IsSelectedAll() || this.IsPlaceHolder())
	{
		oSelectedContent.Add(new AscCommonWord.CSelectedElement(this.Copy(this.Parent)));
	}
	else
	{
		return this.Content.GetSelectedContent(oSelectedContent);
	}
};
CBlockLevelSdt.prototype.PasteFormatting = function(oData)
{
	return this.Content.PasteFormatting(oData);
};
CBlockLevelSdt.prototype.GetCurPosXY = function()
{
	return this.Content.GetCurPosXY();
};
CBlockLevelSdt.prototype.GetSelectedText = function(bClearText, oPr)
{
	if (oPr && oPr.MathAdd && this.IsContentControlEquation() && this.IsPlaceHolder())
		return "";

	return this.Content.GetSelectedText(bClearText, oPr);
};
CBlockLevelSdt.prototype.GetCurrentParagraph = function(bIgnoreSelection, arrSelectedParagraphs, oPr)
{
	if (oPr && true === oPr.ReplacePlaceHolder)
		this.private_ReplacePlaceHolderWithContent();

	return this.Content.GetCurrentParagraph(bIgnoreSelection, arrSelectedParagraphs, oPr);
};
CBlockLevelSdt.prototype.GetCurrentTablesStack = function(arrTables)
{
	if (!arrTables)
		arrTables = [];

	return this.Content.GetCurrentTablesStack(arrTables);
};
CBlockLevelSdt.prototype.AddTableRow = function(bBefore, nCount)
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.AddTableRow(bBefore, nCount);
};
CBlockLevelSdt.prototype.AddTableColumn = function(bBefore, nCount)
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.AddTableColumn(bBefore, nCount);
};
CBlockLevelSdt.prototype.RemoveTableRow = function(nRowIndex)
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.RemoveTableRow(nRowIndex);
};
CBlockLevelSdt.prototype.RemoveTableColumn = function()
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.RemoveTableColumn();
};
CBlockLevelSdt.prototype.MergeTableCells = function()
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.MergeTableCells();
};
CBlockLevelSdt.prototype.SplitTableCells = function(nColsCount, nRowsCount)
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.SplitTableCells(nColsCount, nRowsCount);
};
CBlockLevelSdt.prototype.RemoveTableCells = function()
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.RemoveTableCells();
};
CBlockLevelSdt.prototype.RemoveTable = function()
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.RemoveTable();
};
CBlockLevelSdt.prototype.SelectTable = function(Type)
{
	if (this.IsPlaceHolder())
		return;

	return this.Content.SelectTable(Type);
};
CBlockLevelSdt.prototype.CanMergeTableCells = function()
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.CanMergeTableCells();
};
CBlockLevelSdt.prototype.CanSplitTableCells = function()
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.CanSplitTableCells();
};
CBlockLevelSdt.prototype.DistributeTableCells = function(isHorizontally)
{
	return this.Content.DistributeTableCells(isHorizontally);
};
CBlockLevelSdt.prototype.Document_UpdateInterfaceState = function()
{
	if (!this.IsBuiltInTableOfContents())
		this.LogicDocument.Api.sync_ContentControlCallback(this.GetContentControlPr());

	this.Content.Document_UpdateInterfaceState();
};
CBlockLevelSdt.prototype.Document_UpdateRulersState = function(CurPage)
{
	this.Content.Document_UpdateRulersState(CurPage);
};
CBlockLevelSdt.prototype.GetTableProps = function()
{
	return this.Content.GetTableProps();
};
CBlockLevelSdt.prototype.AddHyperlink = function(Props)
{
	this.private_ReplacePlaceHolderWithContent();
	return this.Content.AddHyperlink(Props);
};
CBlockLevelSdt.prototype.ModifyHyperlink = function(Props)
{
	this.private_ReplacePlaceHolderWithContent();
	this.Content.ModifyHyperlink(Props);
};
CBlockLevelSdt.prototype.RemoveHyperlink = function()
{
	this.private_ReplacePlaceHolderWithContent();
	this.Content.RemoveHyperlink();
};
CBlockLevelSdt.prototype.CanAddHyperlink = function(bCheckInHyperlink)
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.CanAddHyperlink(bCheckInHyperlink);
};
CBlockLevelSdt.prototype.IsCursorInHyperlink = function(bCheckEnd)
{
	if (this.IsPlaceHolder())
		return null;

	return this.Content.IsCursorInHyperlink(bCheckEnd);
};
CBlockLevelSdt.prototype.AddComment = function(Comment, bStart, bEnd)
{
	if (this.IsPlaceHolder())
		return false;

	return this.Content.AddComment(Comment, bStart, bEnd);
};
CBlockLevelSdt.prototype.CanAddComment = function()
{
	return this.Content.CanAddComment();
};
CBlockLevelSdt.prototype.GetSelectionAnchorPos = function()
{
	return this.Content.GetSelectionAnchorPos();
};
CBlockLevelSdt.prototype.GetBoundingPolygonFirstLineH = function()
{
	// TODO: Когда специальные формы будут реализованы и с помощью блочного контрола, тогда
	//       надо будет рассчитать тут
	return 0;
};
/**
 * Функция возвращает рект, содержащий все отрезки данного контрола
 * @return {{X:number, Y:number, W:number, H:number, Page:number, Transform:object}}
 */
CBlockLevelSdt.prototype.GetBoundingRect = function()
{
	var nL = null;
	var nR = null;
	var nT = null;
	var nB = null;

	var nCurPage = -1;
	for (var nPageIndex = 0, nPagesCount = this.GetPagesCount(); nPageIndex < nPagesCount; ++nPageIndex)
	{
		if (this.IsEmptyPage(nPageIndex))
			continue;

		if (-1 === nCurPage)
			nCurPage = nPageIndex;
		else if (nCurPage !== nPageIndex)
			break;

		var oBounds = this.Content.GetContentBounds(nPageIndex);

		if (null === nL || nL > oBounds.Left)
			nL = oBounds.Left;

		if (null === nR || nR < oBounds.Right)
			nR = oBounds.Right;

		if (null === nT || nT > oBounds.Top)
			nT = oBounds.Top;

		if (null === nB || nB < oBounds.Bottom)
			nB = oBounds.Bottom;
	}

	if (-1 === nCurPage || null === nL || null === nT || null === nR || null === nB)
		return {X : 0, Y : 0, W : 0, H : 0, Page : -1, Transform : null};

	return {
		X         : nL,
		Y         : nT,
		W         : nR - nL,
		H         : nB - nT,
		Page      : this.GetAbsolutePage(nCurPage),
		Transform : this.Get_ParentTextTransform()
	};
};
CBlockLevelSdt.prototype.DrawContentControlsTrack = function(nType, X, Y, nCurPage, isCheckHit)
{
	if (!this.IsRecalculated() || !this.LogicDocument)
		return;

	var oLogicDocument   = this.LogicDocument;
	var oDrawingDocument = oLogicDocument.GetDrawingDocument();
	var arrRects = [];

	// TODO: Нужно отрисовать рамку формулы, но для этого нужно чтобы селект плейсхолдера был не целиком на параграф,
	//       а только на формулу
	if (this.IsContentControlEquation() && !this.IsPlaceHolder())
	{
		oDrawingDocument.OnDrawContentControl(null, nType);
		return;
	}

	if (this.IsHideContentControlTrack())
	{
		oDrawingDocument.OnDrawContentControl(null, nType);
		return;
	}

	var oHdrFtr     = this.IsHdrFtr(true);
	var nHdrFtrPage = oHdrFtr ? oHdrFtr.GetContent().GetAbsolutePage(0) : null;
	let isFullWidth = oLogicDocument.IsFillingFormMode();

	for (var nPageIndex = 0, nPagesCount = this.GetPagesCount(); nPageIndex < nPagesCount; ++nPageIndex)
	{
		if (this.IsEmptyPage(nPageIndex))
			continue;

		var nPageAbs = this.GetAbsolutePage(nPageIndex);
		if (null === nHdrFtrPage || nHdrFtrPage === nPageAbs)
		{
			let contentBounds = this.Content.GetContentBounds(nPageIndex);
			let pageBounds    = this.Content.GetPageBounds(nPageIndex);
			if (isFullWidth)
				arrRects.push({X : pageBounds.Left, Y : contentBounds.Top, R : pageBounds.Right, B : contentBounds.Bottom, Page : nPageAbs});
			else
				arrRects.push({X : contentBounds.Left, Y : contentBounds.Top, R : contentBounds.Right, B : contentBounds.Bottom, Page : nPageAbs});
		}
	}

	if (undefined !== X && undefined !== Y && undefined !== nCurPage)
	{
		var nPageAbs = this.GetAbsolutePage(nCurPage);
		var isHit    = false;

		for (var nIndex = 0, nCount = arrRects.length; nIndex < nCount; ++nIndex)
		{
			var oRect = arrRects[nIndex];
			if (nPageAbs === oRect.Page && oRect.X <= X && X <= oRect.R && oRect.Y <= Y && Y <= oRect.B)
			{
				isHit = true;
				break;
			}
		}

		if (false !== isCheckHit && !isHit)
			return;

		var sHelpText = "";
		if (AscCommon.ContentControlTrack.Hover === nType && this.IsForm() && (sHelpText = this.GetFormPr().HelpText))
		{
			var oMMData   = new AscCommon.CMouseMoveData();
			var oCoords   = oDrawingDocument.ConvertCoordsToCursorWR(X, Y, this.Content.GetAbsolutePage(nCurPage), this.Get_ParentTextTransform());
			oMMData.X_abs = oCoords.X - 5;
			oMMData.Y_abs = oCoords.Y;
			oMMData.Type  = Asc.c_oAscMouseMoveDataTypes.Form;
			oMMData.Text  = sHelpText;
			oLogicDocument.GetApi().sync_MouseMoveCallback(oMMData);
		}
	}

	oDrawingDocument.OnDrawContentControl(this, nType, arrRects);
};
CBlockLevelSdt.prototype.AddContentControl = function(nContentControlType)
{
	this.private_ReplacePlaceHolderWithContent();
	return this.Content.AddContentControl(nContentControlType);
};
CBlockLevelSdt.prototype.RecalculateMinMaxContentWidth = function(isRotated)
{
	return this.Content.RecalculateMinMaxContentWidth(isRotated);
};
CBlockLevelSdt.prototype.Shift = function(CurPage, dX, dY)
{
	this.Content.Shift(CurPage, dX, dY);
};
CBlockLevelSdt.prototype.UpdateEndInfo = function()
{
	this.Content.UpdateEndInfo();
};
CBlockLevelSdt.prototype.PrepareRecalculateObject = function()
{
	this.Content.PrepareRecalculateObject();
};
CBlockLevelSdt.prototype.SaveRecalculateObject = function()
{
	return this.Content.SaveRecalculateObject();
};
CBlockLevelSdt.prototype.LoadRecalculateObject = function(RecalcObj)
{
	return this.Content.LoadRecalculateObject(RecalcObj);
};
CBlockLevelSdt.prototype.SetApplyToAll = function(bValue)
{
	this.Content.SetApplyToAll(bValue);
};
CBlockLevelSdt.prototype.IsApplyToAll = function()
{
	return this.Content.IsApplyToAll();
};
CBlockLevelSdt.prototype.RecalculateAllTables = function()
{
	this.Content.RecalculateAllTables();
};
CBlockLevelSdt.prototype.GetAllFloatElements = function(FloatObjects)
{
	return this.Content.GetAllFloatElements(FloatObjects);
};
CBlockLevelSdt.prototype.Get_FirstParagraph = function()
{
	return this.Content.Get_FirstParagraph();
};
CBlockLevelSdt.prototype.StartFromNewPage = function()
{
	this.Content.StartFromNewPage();
};
CBlockLevelSdt.prototype.CollectDocumentStatistics = function(Stats)
{
	return this.Content.CollectDocumentStatistics(Stats);
};
CBlockLevelSdt.prototype.CompareDrawingsLogicPositions = function(CompareObject)
{
	return this.Content.CompareDrawingsLogicPositions(CompareObject);
};
CBlockLevelSdt.prototype.GetStyleFromFormatting = function()
{
	return this.Content.GetStyleFromFormatting();
};
CBlockLevelSdt.prototype.GetAllContentControls = function(arrContentControls)
{
	if (!arrContentControls)
		arrContentControls = [];
	else
		arrContentControls.push(this);

	this.Content.GetAllContentControls(arrContentControls);

	return arrContentControls;
};
CBlockLevelSdt.prototype.IsSelectedAll = function()
{
	return this.Content.IsSelectedAll();
};
CBlockLevelSdt.prototype.IsApplyToAll = function()
{
	return this.Content.IsApplyToAll();
};
CBlockLevelSdt.prototype.GetLastRangeVisibleBounds = function()
{
	return this.Content.GetLastRangeVisibleBounds();
};
CBlockLevelSdt.prototype.FindNextFillingForm = function(isNext, isCurrent, isStart)
{
	if (isCurrent && true === this.IsSelectedAll())
	{
		if (isNext)
			return this.Content.FindNextFillingForm(isNext, isCurrent, isStart);

		return null;
	}

	if (!isCurrent && isNext)
		return this;

	var oRes = this.Content.FindNextFillingForm(isNext, isCurrent, isStart);
	if (oRes)
		return oRes;

	if (!isNext)
		return this;

	return null;
};
CBlockLevelSdt.prototype.GetRevisionsChangeElement = function(SearchEngine)
{
	return this.Content.GetRevisionsChangeElement(SearchEngine);
};
CBlockLevelSdt.prototype.AcceptRevisionChanges = function(Type, bAll)
{
	this.Content.AcceptRevisionChanges(Type, bAll);
};
CBlockLevelSdt.prototype.RejectRevisionChanges = function(Type, bAll)
{
	this.Content.RejectRevisionChanges(Type, bAll);
};
CBlockLevelSdt.prototype.IsContentOnFirstPage = function()
{
	return this.Content.IsContentOnFirstPage();
};
//----------------------------------------------------------------------------------------------------------------------
CBlockLevelSdt.prototype.IsHdrFtr = function(bReturnHdrFtr)
{
	return this.Parent.IsHdrFtr(bReturnHdrFtr);
};
CBlockLevelSdt.prototype.IsFootnote = function(bReturnFootnote)
{
	return this.Parent.IsFootnote(bReturnFootnote);
};
CBlockLevelSdt.prototype.Is_TopDocument = function(bReturnTopDocument)
{
	return this.Parent.Is_TopDocument(bReturnTopDocument);
};
CBlockLevelSdt.prototype.IsCell = function(isReturnCell)
{
	return this.Parent.IsTableCellContent(isReturnCell);
};
CBlockLevelSdt.prototype.Is_DrawingShape = function(bRetShape)
{
	return this.Parent ? this.Parent.Is_DrawingShape(bRetShape) : (bRetShape ? null : false);
};
CBlockLevelSdt.prototype.Get_Numbering = function()
{
	return this.LogicDocument.Get_Numbering();
};
CBlockLevelSdt.prototype.Get_Styles = function()
{
	return this.LogicDocument.Get_Styles();
};
CBlockLevelSdt.prototype.Get_TableStyleForPara = function()
{
	return this.Parent ? this.Parent.Get_TableStyleForPara() : null;
};
CBlockLevelSdt.prototype.Get_ShapeStyleForPara = function()
{
	return this.Parent ? this.Parent.Get_ShapeStyleForPara() : null;
};
CBlockLevelSdt.prototype.Get_Theme = function()
{
	return this.Parent.Get_Theme();
};
CBlockLevelSdt.prototype.GetPrevElementEndInfo = function()
{
	return this.Parent.GetPrevElementEndInfo(this);
};
CBlockLevelSdt.prototype.GetEndInfo = function()
{
	return this.Content.GetEndInfo();
};
CBlockLevelSdt.prototype.IsUseInDocument = function(Id)
{
	if (Id === this.Content.GetId() && this.Parent)
		return this.Parent.IsUseInDocument(this.GetId());
	else if (this.Parent && this.Parent.IsUseInDocument)
		return this.Parent.IsUseInDocument(this.Get_Id());

	return false;
};
CBlockLevelSdt.prototype.Get_ColorMap = function()
{
	return this.Parent.Get_ColorMap();
};
CBlockLevelSdt.prototype.Get_TextBackGroundColor = function()
{
	return this.Parent.Get_TextBackGroundColor();
};
CBlockLevelSdt.prototype.IsThisElementCurrent = function(oElement)
{
	if (oElement === this.Content)
		return this.Parent.IsThisElementCurrent();

	return false;
};
CBlockLevelSdt.prototype.OnContentReDraw = function(StartPageAbs, EndPageAbs)
{
	this.Parent.OnContentReDraw(StartPageAbs, EndPageAbs);
};
CBlockLevelSdt.prototype.Document_CreateFontMap = function(FontMap)
{
	this.Content.Document_CreateFontMap(FontMap);
};
CBlockLevelSdt.prototype.Document_CreateFontCharMap = function(FontCharMap)
{
	this.Content.Document_CreateFontCharMap(FontCharMap);
};
CBlockLevelSdt.prototype.Document_Get_AllFontNames = function(AllFonts)
{
	this.Content.Document_Get_AllFontNames(AllFonts);
};
CBlockLevelSdt.prototype.Get_ParentTextTransform = function()
{
	return this.Parent.Get_ParentTextTransform();
};
CBlockLevelSdt.prototype.Set_CurrentElement = function(bUpdateStates, PageAbs, oDocContent)
{
	if (oDocContent === this.Content)
	{
		if (this.IsPlaceHolder())
			this.SelectAll(1);

		var nIndex = this.GetIndex();
		if (-1 !== nIndex)
			this.Parent.Set_CurrentElement(nIndex, bUpdateStates);
	}
};
CBlockLevelSdt.prototype.Refresh_RecalcData2 = function(CurPage)
{
	if (!this.Parent)
		return;

	this.Parent.Refresh_RecalcData2(this.Index, this.private_GetRelativePageIndex(CurPage));
};
CBlockLevelSdt.prototype.Refresh_RecalcData = function(Data)
{
};
CBlockLevelSdt.prototype.Check_AutoFit = function()
{
	return this.Parent.Check_AutoFit();
};
CBlockLevelSdt.prototype.IsInTable = function(bReturnTopTable)
{
	return this.Parent.IsInTable(bReturnTopTable);
};
CBlockLevelSdt.prototype.Get_PageContentStartPos = function(CurPage)
{
	var StartPage   = this.Get_AbsolutePage(0);
	var StartColumn = this.Get_AbsoluteColumn(0);

	if (this.Parent instanceof CDocumentContent)
	{
		StartPage   = this.Parent.Get_AbsolutePage(0) - StartPage;
		StartColumn = this.Parent.StartColumn;

		// Такого не должно быть, но на всякий случай
		if (StartPage < 0)
			StartPage = 0;
	}

	return this.Parent.Get_PageContentStartPos2(StartPage, StartColumn, CurPage, this.Index);
};
CBlockLevelSdt.prototype.CheckTableCoincidence = function(Table)
{
	return this.Parent.CheckTableCoincidence(Table);
};
CBlockLevelSdt.prototype.CheckRange = function(X0, Y0, X1, Y1, _Y0, _Y1, X_lf, X_rf, CurPage, Inner, bMathWrap)
{
	if (true === Inner)
	{
		var PageRel = this.Get_AbsolutePage(CurPage) - this.Get_AbsolutePage(0) + this.Get_StartPage_Relative();
		return this.Parent.CheckRange(X0, Y0, X1, Y1, _Y0, _Y1, X_lf, X_rf, PageRel, Inner, bMathWrap);
	}
	else
	{
		return this.Content.CheckRange(X0, Y0, X1, Y1, _Y0, _Y1, X_lf, X_rf, CurPage, Inner, bMathWrap);
	}
};
CBlockLevelSdt.prototype.GetTopDocumentContent = function(isOneLevel)
{
	return this.Parent.GetTopDocumentContent(isOneLevel);
};
CBlockLevelSdt.prototype.GetAllDrawingObjects = function(AllDrawingObjects)
{
	return this.Content.GetAllDrawingObjects(AllDrawingObjects);
};
CBlockLevelSdt.prototype.GetAllComments = function(AllComments)
{
	return this.Content.GetAllComments(AllComments);
};
CBlockLevelSdt.prototype.GetAllMaths = function(AllMaths)
{
	return this.Content.GetAllMaths(AllMaths);
};

CBlockLevelSdt.prototype.GetAllSeqFieldsByType = function(sType, aFields)
{
	return this.Content.GetAllSeqFieldsByType(sType, aFields);
};
CBlockLevelSdt.prototype.Get_ParentTextTransform = function()
{
	return this.Parent.Get_ParentTextTransform();
};
CBlockLevelSdt.prototype.Get_SectPr = function()
{
	if (this.Parent && this.Parent.Get_SectPr)
	{
		this.Parent.Update_ContentIndexing();
		return this.Parent.Get_SectPr(this.Index);
	}

	return null;
};
CBlockLevelSdt.prototype.GetMargins = function()
{
	return this.Parent.GetMargins();
};
CBlockLevelSdt.prototype.SelectContentControl = function()
{
	this.SelectAll(1);
	this.Set_CurrentElement(false, 0, this.Content);
};
CBlockLevelSdt.prototype.MoveCursorToContentControl = function(isBegin)
{
	if (isBegin)
		this.MoveCursorToStartPos(false);
	else
		this.MoveCursorToEndPos(false);

	this.SetThisElementCurrent();
};
CBlockLevelSdt.prototype.SetThisElementCurrent = function()
{
	this.Set_CurrentElement(false, 0, this.Content);
};
CBlockLevelSdt.prototype.RemoveContentControlWrapper = function()
{
	if (!this.Parent)
		return;

	this.Parent.Update_ContentIndexing();
	var nElementPos = this.GetIndex();

	if (this.Parent.Content[nElementPos] !== this)
		return;

	var nParentCurPos            = this.Parent.CurPos.ContentPos;
	var nParentSelectionStartPos = this.Parent.Selection.StartPos;
	var nParentSelectionEndPos   = this.Parent.Selection.EndPos;

	this.Parent.Remove_FromContent(nElementPos, 1);
	for (var nIndex = 0, nCount = this.Content.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Parent.Add_ToContent(nElementPos + nIndex, this.Content.Content[nIndex]);
	}

	if (nParentCurPos === nElementPos)
		this.Parent.CurPos.ContentPos = nParentCurPos + this.Content.CurPos.ContentPos;
	else if (nParentCurPos > nElementPos)
		this.Parent.CurPos.ContentPos = nParentCurPos + nCount - 1;

	if (nParentSelectionStartPos === nElementPos)
		this.Parent.Selection.StartPos = nParentSelectionStartPos + this.Content.Selection.StartPos;
	else if (nParentSelectionStartPos > nElementPos)
		this.Parent.Selection.StartPos = nParentSelectionStartPos + nCount - 1;

	if (nParentSelectionEndPos === nElementPos)
		this.Parent.Selection.EndPos = nParentSelectionEndPos + this.Content.Selection.EndPos;
	else if (nParentSelectionEndPos > nElementPos)
		this.Parent.Selection.EndPos = nParentSelectionEndPos + nCount - 1;

	this.Content.RemoveFromContent(0, this.Content.Content.length, false);
};
CBlockLevelSdt.prototype.IsTableFirstRowOnNewPage = function()
{
	return this.Parent.IsTableFirstRowOnNewPage();
};
CBlockLevelSdt.prototype.UpdateBookmarks = function(oManager)
{
	this.Content.UpdateBookmarks(oManager);
};
CBlockLevelSdt.prototype.GetSimilarNumbering = function(oContinueEngine)
{
	if (oContinueEngine.IsFound())
		return;

	this.Content.GetSimilarNumbering(oContinueEngine);
};
CBlockLevelSdt.prototype.GetTableOfContents = function(isUnique, isCheckFields)
{
	if (this.IsBuiltInTableOfContents() && (!isUnique || this.IsBuiltInUnique()))
		return this;

	return this.Content.GetTableOfContents(isCheckFields);
};
CBlockLevelSdt.prototype.GetInnerTableOfContents = function()
{
	var oTOC = this.Content.GetTableOfContents(false, true);
	if (oTOC instanceof CBlockLevelSdt)
		return oTOC.GetInnerTableOfContents();

	return oTOC;
};
CBlockLevelSdt.prototype.GetTablesOfFigures = function(arrComplexFields)
{
	this.Content.GetTablesOfFigures(arrComplexFields);
};
CBlockLevelSdt.prototype.IsBlockLevelSdtFirstOnNewPage = function()
{
	if ((null != this.Get_DocumentPrev() && !this.Parent.IsElementStartOnNewPage(this.GetIndex()))
		|| (true === this.Parent.IsTableCellContent() && true !== this.Parent.IsTableFirstRowOnNewPage())
		|| (true === this.Parent.IsBlockLevelSdtContent() && true !== this.Parent.IsBlockLevelSdtFirstOnNewPage()))
		return false;

	return true;
};
//----------------------------------------------------------------------------------------------------------------------
CBlockLevelSdt.prototype.GetContentControlType = function()
{
	return c_oAscSdtLevelType.Block;
};
CBlockLevelSdt.prototype.SetPr = function(oPr)
{
	if (!oPr || this.IsBuiltInTableOfContents())
		return;

	this.SetAlias(oPr.Alias);
	this.SetTag(oPr.Tag);
	this.SetLabel(oPr.Label);
	this.SetContentControlLock(oPr.Lock);
	this.SetContentControlId(oPr.Id);

	if (undefined !== oPr.DocPartObj)
		this.SetDocPartObj(oPr.DocPartObj.Category, oPr.DocPartObj.Gallery, oPr.DocPartObj.Unique);

	if (undefined !== oPr.Appearance)
		this.SetAppearance(oPr.Appearance);

	if (undefined !== oPr.Color)
		this.SetColor(oPr.Color);
};
/**
 * Выставляем настройки текста по умолчанию для данного контрола
 * @param {CTextPr} oTextPr
 */
CBlockLevelSdt.prototype.SetDefaultTextPr = function(oTextPr)
{
	if (oTextPr && !this.Pr.TextPr.IsEqual(oTextPr))
	{
		History.Add(new CChangesSdtPrTextPr(this, this.Pr.TextPr, oTextPr));
		this.Pr.TextPr = oTextPr;
	}
};
/**
 * Получаем настройки для текста по умолчанию
 * @returns {CTextPr}
 */
CBlockLevelSdt.prototype.GetDefaultTextPr = function()
{
	return this.Pr.TextPr;
};
CBlockLevelSdt.prototype.SetAlias = function(sAlias)
{
	if (sAlias !== this.Pr.Alias)
	{
		History.Add(new CChangesSdtPrAlias(this, this.Pr.Alias, sAlias));
		this.Pr.Alias = sAlias;
	}
};
CBlockLevelSdt.prototype.GetAlias = function()
{
	return (undefined !== this.Pr.Alias ? this.Pr.Alias : "");
};
CBlockLevelSdt.prototype.SetContentControlId = function(Id)
{
	if (this.Pr.Id !== Id)
	{
		History.Add(new CChangesSdtPrId(this, this.Pr.Id, Id));
		this.Pr.Id = Id;
	}
};
CBlockLevelSdt.prototype.GetContentControlId = function()
{
	return this.Pr.Id;
};
CBlockLevelSdt.prototype.SetTag = function(sTag)
{
	if (this.Pr.Tag !== sTag)
	{
		History.Add(new CChangesSdtPrTag(this, this.Pr.Tag, sTag));
		this.Pr.Tag = sTag;
	}
};
CBlockLevelSdt.prototype.GetTag = function()
{
	return (undefined !== this.Pr.Tag ? this.Pr.Tag : "");
};
CBlockLevelSdt.prototype.SetLabel = function(sLabel)
{
	if (this.Pr.Label !== sLabel)
	{
		History.Add(new CChangesSdtPrLabel(this, this.Pr.Label, sLabel));
		this.Pr.Label = sLabel;
	}
};
CBlockLevelSdt.prototype.GetLabel = function()
{
	return (undefined !== this.Pr.Label ? this.Pr.Label : "");
};
CBlockLevelSdt.prototype.SetAppearance = function(nType)
{
	if (this.Pr.Appearance !== nType)
	{
		History.Add(new CChangesSdtPrAppearance(this, this.Pr.Appearance, nType));
		this.Pr.Appearance = nType;
	}
};
CBlockLevelSdt.prototype.GetAppearance = function()
{
	return this.Pr.Appearance;
};
CBlockLevelSdt.prototype.SetColor = function(oColor)
{
	if (null === oColor || undefined === oColor)
	{
		if (undefined !== this.Pr.Color)
		{
			History.Add(new CChangesSdtPrColor(this, this.Pr.Color, undefined));
			this.Pr.Color = undefined;
		}
	}
	else
	{
		History.Add(new CChangesSdtPrColor(this, this.Pr.Color, oColor));
		this.Pr.Color = oColor;
	}
};
CBlockLevelSdt.prototype.GetColor = function()
{
	return this.Pr.Color;
};
CBlockLevelSdt.prototype.SetDocPartObj = function(sCategory, sGallery, isUnique)
{
	History.Add(new CChangesSdtPrDocPartObj(this, this.Pr.DocPartObj, {Category : sCategory, Gallery : sGallery, Unique : isUnique}));
	this.Pr.DocPartObj.Category = sCategory;
	this.Pr.DocPartObj.Gallery  = sGallery;
	this.Pr.DocPartObj.Unique   = isUnique;
};
CBlockLevelSdt.prototype.SetContentControlLock = function(nLockType)
{
	if (this.Pr.Lock !== nLockType)
	{
		History.Add(new CChangesSdtPrLock(this, this.Pr.Lock, nLockType));
		this.Pr.Lock = nLockType;
	}
};
CBlockLevelSdt.prototype.GetContentControlLock = function()
{
	return (undefined !== this.Pr.Lock ? this.Pr.Lock : c_oAscSdtLockType.Unlocked);
};
CBlockLevelSdt.prototype.SetContentControlPr = function(oPr)
{
	if (!oPr || this.IsBuiltInTableOfContents())
		return;

	if (oPr && !(oPr instanceof CContentControlPr))
	{
		var oTemp = new CContentControlPr(c_oAscSdtLockType.Block);
		oTemp.FillFromObject(oPr);
		oPr = oTemp;
	}

	oPr.SetToContentControl(this);
};
CBlockLevelSdt.prototype.GetContentControlPr = function()
{
	var oPr = new CContentControlPr(c_oAscSdtLevelType.Block);
	oPr.FillFromContentControl(this);
	return oPr;
};
CBlockLevelSdt.prototype.RestartSpellCheck = function()
{
	this.Content.RestartSpellCheck();
};
//----------------------------------------------------------------------------------------------------------------------
// Search
//----------------------------------------------------------------------------------------------------------------------
CBlockLevelSdt.prototype.Search = function(oSearchEngine, nType)
{
	this.Content.Search(oSearchEngine, nType);
};
CBlockLevelSdt.prototype.GetSearchElementId = function(bNext, bCurrent)
{
	return this.Content.GetSearchElementId(bNext, bCurrent);
};
//----------------------------------------------------------------------------------------------------------------------
CBlockLevelSdt.prototype.ClearContentControl = function()
{
	var oParagraph = new Paragraph(this.LogicDocument.Get_DrawingDocument(), this.Content);
	oParagraph.Correct_Content();

	oParagraph.SelectAll();
	oParagraph.ApplyTextPr(this.Pr.TextPr);
	oParagraph.RemoveSelection();

	this.Content.AddToContent(0, oParagraph);
	this.Content.RemoveFromContent(1, this.Content.GetElementsCount() - 1);
	this.Content.RemoveSelection();
	this.Content.MoveCursorToStartPos();
};
CBlockLevelSdt.prototype.GotoFootnoteRef = function(isNext, isCurrent, isStepFootnote, isStepEndnote)
{
	return this.Content.GotoFootnoteRef(isNext, isCurrent, isStepFootnote, isStepEndnote);
};
/**
 * Получаем последний элемент содержимого
 * @returns {?CDocumentContentElementBase}
 */
CBlockLevelSdt.prototype.GetLastElement = function()
{
	var nCount = this.Content.GetElementsCount();
	if (nCount <= 0)
		return null;

	return this.Content.GetElement(nCount - 1);
};
CBlockLevelSdt.prototype.GetElement = function(nIndex)
{
	return this.Content.GetElement(nIndex);
};
CBlockLevelSdt.prototype.GetElementsCount = function()
{
	return this.Content.GetElementsCount();
};
CBlockLevelSdt.prototype.GetLastParagraph = function()
{
	return this.Content.GetLastParagraph();
};
CBlockLevelSdt.prototype.GetOutlineParagraphs = function(arrOutline, oPr)
{
	this.Content.GetOutlineParagraphs(arrOutline, oPr);
};
CBlockLevelSdt.prototype.IsLastTableCellInRow = function(isSelection)
{
	return this.Parent.IsLastTableCellInRow(isSelection);
};
/**
 * Можно ли удалить данный контейнер
 * @returns {boolean}
 */
CBlockLevelSdt.prototype.CanBeDeleted = function()
{
	return (undefined === this.Pr.Lock || c_oAscSdtLockType.Unlocked === this.Pr.Lock || c_oAscSdtLockType.ContentLocked === this.Pr.Lock);
};
/**
 * Можно ли редактировать данный контейнер
 * @returns {boolean}
 */
CBlockLevelSdt.prototype.CanBeEdited = function()
{
	if (!this.SkipSpecialLock && (this.IsCheckBox() || this.IsPicture() || this.IsDropDownList()))
		return false;

	return (undefined === this.Pr.Lock || c_oAscSdtLockType.Unlocked === this.Pr.Lock || c_oAscSdtLockType.SdtLocked === this.Pr.Lock);
};
/**
 * Активен PlaceHolder сейчас или нет
 * @returns {boolean}
 */
CBlockLevelSdt.prototype.IsPlaceHolder = function()
{
	return this.Pr.ShowingPlcHdr;
};
CBlockLevelSdt.prototype.private_ReplacePlaceHolderWithContent = function(isSkipTemporaryCheck)
{
	if (!this.IsPlaceHolder())
		return;

	this.SetShowingPlcHdr(false);

	this.Content.RemoveFromContent(0, this.Content.GetElementsCount(), false);

	var oParagraph = new Paragraph(this.LogicDocument ? this.LogicDocument.GetDrawingDocument() : null, this.Content, false);
	oParagraph.Correct_Content();

	oParagraph.SelectAll();
	oParagraph.ApplyTextPr(this.Pr.TextPr);
	oParagraph.RemoveSelection();

	this.Content.AddToContent(0, oParagraph);
	this.Content.RemoveSelection();
	this.Content.MoveCursorToStartPos();

	if (this.IsContentControlEquation())
	{
		var oParaMath = new ParaMath();
		oParaMath.Root.Load_FromMenu(c_oAscMathType.Default_Text, oParagraph, null);
		oParaMath.Root.Correct_Content(true);
		oParagraph.AddToContent(0, oParaMath);
		oParaMath.SetThisElementCurrentInParagraph();
		oParaMath.MoveCursorToStartPos();
	}

	if (true !== isSkipTemporaryCheck && this.IsContentControlTemporary())
		this.RemoveContentControlWrapper();
};
CBlockLevelSdt.prototype.private_ReplaceContentWithPlaceHolder = function(isSelect, isForceUpdate)
{
	if (this.IsPlaceHolder())
	{
		if (isForceUpdate)
			this.private_FillPlaceholderContent();
		
		return;
	}

	this.SetShowingPlcHdr(true);
	this.private_FillPlaceholderContent();

	if (false !== isSelect)
		this.SelectContentControl();
};
CBlockLevelSdt.prototype.private_FillPlaceholderContent = function()
{
	this.Content.RemoveFromContent(0, this.Content.GetElementsCount(), false);

	var oLogicDocument = this.GetLogicDocument() ? this.GetLogicDocument() : editor.WordControl.m_oLogicDocument;
	var oDocPart       = oLogicDocument.GetGlossaryDocument().GetDocPartByName(this.GetPlaceholder());
	if (oDocPart)
	{
		if (this.IsContentControlEquation())
		{
			var oFirstParagraph = oDocPart.GetFirstParagraph();

			var oNewParagraph = oFirstParagraph.Copy();
			oNewParagraph.RemoveFromContent(0, oNewParagraph.GetElementsCount());

			var oParaMath = new ParaMath();
			oParaMath.Root.Load_FromMenu(c_oAscMathType.Default_Text, oNewParagraph, null, oFirstParagraph.GetText());
			oParaMath.Root.Correct_Content(true);
			oNewParagraph.AddToContent(0, oParaMath);
			oNewParagraph.CorrectContent();

			this.Content.AddToContent(0, oNewParagraph);
		}
		else
		{
			for (var nIndex = 0, nCount = oDocPart.GetElementsCount(); nIndex < nCount; ++nIndex)
			{
				this.Content.AddToContent(0, oDocPart.GetElement(nIndex).Copy());
			}
		}
	}
	else
	{
		var oParagraph = new Paragraph(this.LogicDocument ? this.LogicDocument.GetDrawingDocument() : null, this.Content, false);
		oParagraph.Correct_Content();

		oParagraph.SelectAll();
		oParagraph.ApplyTextPr(this.Pr.TextPr);
		oParagraph.RemoveSelection();

		this.Content.AddToContent(0, oParagraph);
		if (this.IsContentControlEquation())
		{
			var oParaMath = new ParaMath();
			oParaMath.Root.Load_FromMenu(c_oAscMathType.Default_Text, oParagraph, null, String.fromCharCode(nbsp_charcode, nbsp_charcode, nbsp_charcode, nbsp_charcode));
			oParaMath.Root.Correct_Content(true);
			oParagraph.AddToContent(0, oParaMath);
			oParagraph.CorrectContent();
		}
		else
		{
			var oRun = new ParaRun(oParagraph, false);
			oRun.AddText(String.fromCharCode(nbsp_charcode, nbsp_charcode, nbsp_charcode, nbsp_charcode));
			oParagraph.AddToContent(0, oRun);
		}
	}
};
CBlockLevelSdt.prototype.GetPlaceHolderObject = function()
{
	if (this.IsPlaceHolder())
		return this;

	return this.Content.GetPlaceHolderObject();
};
CBlockLevelSdt.prototype.GetAllFields = function(isUseSelection, arrFields)
{
	if (this.IsPlaceHolder())
		return arrFields ? arrFields : [];

	return this.Content.GetAllFields(isUseSelection, arrFields);
};
CBlockLevelSdt.prototype.ReplacePlaceHolderWithContent = function()
{
	return this.private_ReplacePlaceHolderWithContent();
};
CBlockLevelSdt.prototype.ReplaceContentWithPlaceHolder = function(isSelect, isForceUpdate)
{
	return this.private_ReplaceContentWithPlaceHolder(isSelect, isForceUpdate);
};
CBlockLevelSdt.prototype.CheckRunContent = function(fCheck)
{
	return this.Content.CheckRunContent(fCheck);
};
CBlockLevelSdt.prototype.CheckSelectedRunContent = function(fCheck)
{
	return this.Content.CheckSelectedRunContent(fCheck);
};
CBlockLevelSdt.prototype.IsTableCellSelection = function()
{
	return this.Content.IsTableCellSelection();
};
/**
 * Проверяем, является ли данный контейнер чекбоксом
 * @returns {boolean}
 */
CBlockLevelSdt.prototype.IsCheckBox = function()
{
	return !!(this.Pr.CheckBox);
};
/**
 * Выключаем проверку невозможности редактирования данного объекта, из-за того что это чекбокс
 * @param isSkip {boolean}
 */
CBlockLevelSdt.prototype.SkipSpecialContentControlLock = function(isSkip)
{
	this.SkipSpecialLock = isSkip;
};
/**
 * @retuns {boolean}
 */
CBlockLevelSdt.prototype.IsSkipSpecialContentControlLock = function()
{
	return this.SkipSpecialLock;
};
/**
 * Применяем заданные настройки для чекобокса
 * @param {CSdtCheckBoxPr} oCheckBoxPr
 * @param {CTextPr} oTextPr
 */
CBlockLevelSdt.prototype.ApplyCheckBoxPr = function(oCheckBoxPr, oTextPr)
{
	if (undefined === this.Pr.CheckBox || !this.Pr.CheckBox.IsEqual(oCheckBoxPr))
	{
		if (this.IsPlaceHolder())
			this.private_ReplacePlaceHolderWithContent(false);

		this.SetCheckBoxPr(oCheckBoxPr);
	}

	if (this.IsCheckBox())
	{
		if (oTextPr)
		{
			if (this.Content.GetElementsCount() >= 1 && this.Content.GetElement(0).IsParagraph())
			{
				var oPara = this.Content.GetElement(0);

				if (oPara.Content[0] && para_Run === oPara.Content[0].Type)
					oPara.Content[0].SetPr(oTextPr);
			}
		}

		this.private_UpdateCheckBoxContent();
		this.SetShowingPlcHdr(false);
		this.SetPlaceholder(undefined);
	}
};
/**
 * Выставляем настройки чекбокса
 * @param {CSdtCheckBoxPr} oCheckBoxPr
 */
CBlockLevelSdt.prototype.SetCheckBoxPr = function(oCheckBoxPr)
{
	if (undefined === this.Pr.CheckBox || !this.Pr.CheckBox.IsEqual(oCheckBoxPr))
	{
		History.Add(new CChangesSdtPrCheckBox(this, this.Pr.CheckBox, oCheckBoxPr));
		this.Pr.CheckBox = oCheckBoxPr;
	}
};
/**
 * Получаем настройки для чекбокса
 * @returns {?CSdtCheckBoxPr}
 */
CBlockLevelSdt.prototype.GetCheckBoxPr = function()
{
	return this.Pr.CheckBox;
};
/**
 * Выставляем состояние чекбокса
 * @param {boolean|undefined} [isChecked=undefined]
 */
CBlockLevelSdt.prototype.ToggleCheckBox = function(isChecked)
{
	if (!this.IsCheckBox())
		return;

	if (undefined !== isChecked && this.Pr.CheckBox.Checked === isChecked)
		return;

	let oLogicDocument = this.GetLogicDocument();
	if (oLogicDocument && oLogicDocument.IsDocumentEditor())
		oLogicDocument.OnChangeForm(this);

	if (undefined === isChecked && this.IsRadioButton() && true === this.Pr.CheckBox.Checked)
		return;

	this.SetCheckBoxChecked(!this.Pr.CheckBox.Checked);
};
CBlockLevelSdt.prototype.SetCheckBoxChecked = function(isChecked)
{
	History.Add(new CChangesSdtPrCheckBoxChecked(this, this.Pr.CheckBox.Checked, isChecked));
	this.Pr.CheckBox.Checked = isChecked;

	this.private_UpdateCheckBoxContent();
};
CBlockLevelSdt.prototype.private_UpdateCheckBoxContent = function()
{
	var isChecked = this.Pr.CheckBox.Checked;

	var oRun;
	if (this.LogicDocument && this.LogicDocument.IsTrackRevisions())
	{
		var oFirstParagraph = this.GetFirstParagraph();
		var oFirstRun       = oFirstParagraph ? oFirstParagraph.GetFirstRun() : null;
		var oTextPr         = oFirstRun ? oFirstRun.GetDirectTextPr() : new CTextPr();

		this.SelectAll();
		this.Remove();
		this.RemoveSelection();

		oFirstParagraph = null;

		var oDocContent = this.Content;
		if (oDocContent.Content.length <= 0 || !oDocContent.Content[0].IsParagraph())
		{
			oFirstParagraph = new Paragraph(this.LogicDocument.GetDrawingDocument(), oDocContent);
			oDocContent.AddToContent(0, oFirstParagraph);
		}
		else
		{
			oFirstParagraph = oDocContent.Content[0];
		}
		oFirstParagraph.SetReviewType(reviewtype_Common);

		oRun = new ParaRun(oFirstParagraph, false);
		oRun.SetPr(oTextPr);
		oFirstParagraph.AddToContent(0, oRun);

		if (3 === oFirstParagraph.Content.length
			&& para_Run === oFirstParagraph.Content[0].Type
			&& para_Run === oFirstParagraph.Content[1].Type
			&& reviewtype_Add === oFirstParagraph.Content[0].GetReviewType()
			&& reviewtype_Remove === oFirstParagraph.Content[1].GetReviewType()
			&& oFirstParagraph.Content[0].GetReviewInfo().IsCurrentUser()
			&& oFirstParagraph.Content[1].GetReviewInfo().IsCurrentUser()
			&& ((isChecked
			&& String.fromCharCode(this.Pr.CheckBox.CheckedSymbol) === oFirstParagraph.Content[1].GetText())
			|| (!isChecked
			&& String.fromCharCode(this.Pr.CheckBox.UncheckedSymbol) === oFirstParagraph.Content[1].GetText())))
		{
			oFirstParagraph.RemoveFromContent(1, 1);
			oRun.SetReviewType(reviewtype_Common);
		}
	}
	else
	{
		var oPara = this.Content.MakeSingleParagraphContent();
		oRun  = oPara.MakeSingleRunParagraph();
		oRun.ClearContent();
	}

	oRun.AddText(String.fromCharCode(isChecked ? this.Pr.CheckBox.CheckedSymbol : this.Pr.CheckBox.UncheckedSymbol));

	if (isChecked && this.Pr.CheckBox.CheckedFont)
	{
		oRun.SetRFontsAscii({Index : -1, Name : this.Pr.CheckBox.CheckedFont});
		oRun.SetRFontsHAnsi({Index : -1, Name : this.Pr.CheckBox.CheckedFont});
		oRun.SetRFontsCS({Index : -1, Name : this.Pr.CheckBox.CheckedFont});
		oRun.SetRFontsEastAsia({Index : -1, Name : this.Pr.CheckBox.CheckedFont});
	}
	else if (!isChecked && this.Pr.CheckBox.UncheckedFont)
	{
		oRun.SetRFontsAscii({Index : -1, Name : this.Pr.CheckBox.UncheckedFont});
		oRun.SetRFontsHAnsi({Index : -1, Name : this.Pr.CheckBox.UncheckedFont});
		oRun.SetRFontsCS({Index : -1, Name : this.Pr.CheckBox.UncheckedFont});
		oRun.SetRFontsEastAsia({Index : -1, Name : this.Pr.CheckBox.UncheckedFont});
	}
};
/**
 * Проверяем, является ли данный класс специальным контейнером для картинки
 * @returns {boolean}
 */
CBlockLevelSdt.prototype.IsPicture = function()
{
	return (!!this.Pr.Picture);
};
/**
 * Выставляем настройку того, что это контент контрол с картинкой
 * @param isPicture {boolean}
 */
CBlockLevelSdt.prototype.SetPicturePr = function(isPicture)
{
	if (this.Pr.Picture !== isPicture)
	{
		History.Add(new CChangesSdtPrPicture(this, this.Pr.Picture, isPicture));
		this.Pr.Picture = isPicture;
	}
};
CBlockLevelSdt.prototype.private_UpdatePictureContent = function(_nW, _nH)
{
	if (!this.IsPicture())
		return;

	var arrDrawings = this.GetAllDrawingObjects();

	if (this.IsPlaceHolder())
	{
		this.ReplacePlaceHolderWithContent();
		this.SetShowingPlcHdr(true);
	}

	var oDrawing;
	for (var nIndex = 0, nCount = arrDrawings.length; nIndex < nCount; ++nIndex)
	{
		if (arrDrawings[nIndex].IsPicture())
		{
			oDrawing = arrDrawings[nIndex];
			break;
		}
	}

	var oPara = this.Content.MakeSingleParagraphContent();
	var oRun  = oPara.MakeSingleRunParagraph();

	if (!oDrawing)
	{
		var oDrawingObjects = this.LogicDocument ? this.LogicDocument.DrawingObjects : null;
		if (!oDrawingObjects)
			return;

		var nW = _nW ? _nW : 50;
		var nH = _nH ? _nH : 50;

		oDrawing   = new ParaDrawing(nW, nH, null, oDrawingObjects, this.LogicDocument, null);
		var oImage = oDrawingObjects.createImage(AscCommon.g_sWordPlaceholderImage, 0, 0, nW, nH);
		oImage.setParent(oDrawing);
		oDrawing.Set_GraphicObject(oImage);
	}

	oRun.AddToContent(0, oDrawing);
};
/**
 * Применяме к данному контейнеру настройку того, что это специальный контейнер для картинок
 * @param isPicture {boolean}
 * @param [nW=-1] {number} если не задано (или значение не положительное), тогда используется стандартное значение
 * @param [nH=-1] {number} если не задано (или значение не положительное), тогда используется стандартное значение
 */
CBlockLevelSdt.prototype.ApplyPicturePr = function(isPicture, nW, nH)
{
	this.SetPicturePr(isPicture);
	this.SetPlaceholder(undefined);
	this.private_UpdatePictureContent(nW, nH);
};
/**
 * Выделяем изображение, если это специальный контейнер для изображения
 * @returns {boolean}
 */
CBlockLevelSdt.prototype.SelectPicture = function()
{
	if (!this.IsPicture())
		return false;

	var arrDrawings = this.GetAllDrawingObjects();
	if (arrDrawings.length <= 0)
		return false;

	this.Content.Select_DrawingObject(arrDrawings[0].GetId());
	return true;
};
/**
 * Проверяем является ли данный контейнер специальным для поля со списком
 * @returns {boolean}
 */
CBlockLevelSdt.prototype.IsComboBox = function()
{
	return (undefined !== this.Pr.ComboBox);
};
/**
 * @param oPr {AscWord.CSdtComboBoxPr}
 */
CBlockLevelSdt.prototype.SetComboBoxPr = function(oPr)
{
	if (undefined === this.Pr.ComboBox || !this.Pr.ComboBox.IsEqual(oPr))
	{
		History.Add(new CChangesSdtPrComboBox(this, this.Pr.ComboBox, oPr));
		this.Pr.ComboBox = oPr;
	}
};
/**
 * @returns {?AscWord.CSdtComboBoxPr}
 */
CBlockLevelSdt.prototype.GetComboBoxPr = function()
{
	return this.Pr.ComboBox;
};
/**
 * Проверяем является ли данный контейнер специальным для выпадающего списка
 * @returns {boolean}
 */
CBlockLevelSdt.prototype.IsDropDownList = function()
{
	return (undefined !== this.Pr.DropDown);
};
/**
 * @param oPr {AscWord.CSdtComboBoxPr}
 */
CBlockLevelSdt.prototype.SetDropDownListPr = function(oPr)
{
	if (undefined === this.Pr.DropDown || !this.Pr.DropDown.IsEqual(oPr))
	{
		History.Add(new CChangesSdtPrDropDownList(this, this.Pr.DropDown, oPr));
		this.Pr.DropDown = oPr;
	}
};
/**
 * @returns {?AscWord.CSdtComboBoxPr}
 */
CBlockLevelSdt.prototype.GetDropDownListPr = function()
{
	return this.Pr.DropDown;
};
/**
 * Применяем к данному контейнеру настройки того, что это специальный контйенер для поля со списком
 * @param oPr {AscWord.CSdtComboBoxPr}
 */
CBlockLevelSdt.prototype.ApplyComboBoxPr = function(oPr)
{
	this.SetPlaceholder(c_oAscDefaultPlaceholderName.List);
	if (this.IsPlaceHolder())
		this.private_FillPlaceholderContent();

	this.SetComboBoxPr(oPr);
	this.SelectListItem();
};
/**
 * Применяем к данному контейнеру настройки того, что это специальный контейнер для выпадающего списка
 * @param oPr {AscWord.CSdtComboBoxPr}
 */
CBlockLevelSdt.prototype.ApplyDropDownListPr = function(oPr)
{
	this.SetPlaceholder(c_oAscDefaultPlaceholderName.List);
	if (this.IsPlaceHolder())
		this.private_FillPlaceholderContent();

	this.SetDropDownListPr(oPr);
	this.SelectListItem();
};
/**
 * Заполняем контейнер текстом в зависимости от выбранного элемента в списке
 * @param sValue {string}
 */
CBlockLevelSdt.prototype.SelectListItem = function(sValue)
{
	var oList = null;
	if (this.IsComboBox())
		oList = this.Pr.ComboBox;
	else if (this.IsDropDownList())
		oList = this.Pr.DropDown;

	if (!oList)
		return;

	var sText = oList.GetTextByValue(sValue);

	if (this.LogicDocument && this.LogicDocument.IsTrackRevisions())
	{
		if (!sText && this.IsPlaceHolder())
		{
			this.private_FillPlaceholderContent();
			return;
		}

		var oFirstParagraph = this.GetFirstParagraph();
		var oFirstRun       = oFirstParagraph ? oFirstParagraph.GetFirstRun() : null;
		var oTextPr         = oFirstRun ? oFirstRun.GetDirectTextPr() : new CTextPr();

		if (!this.IsPlaceHolder())
		{
			this.SelectAll();
			this.Remove();
			this.RemoveSelection();
		}
		else
		{
			this.ReplacePlaceHolderWithContent();
		}

		if (!sText && this.IsEmpty())
		{
			this.ReplaceContentWithPlaceHolder();
		}

		if (sText)
		{
			var oRun, oParagraph;
			if (this.IsEmpty())
			{
				oParagraph = this.Content.MakeSingleParagraphContent();
				if (!oParagraph)
					return;

				oRun = oParagraph.MakeSingleRunParagraph();
				if (!oRun)
					return;

				oRun.SetReviewType(reviewtype_Add);
			}
			else
			{
				if (this.Content.Content[this.Content.Content.length - 1].IsParagraph())
				{
					oParagraph = this.Content.Content[this.Content.Content.length - 1];
				}
				else
				{
					oParagraph = new Paragraph(this.LogicDocument.GetDrawingDocument(), this.Content);
					this.Content.AddToParagraph(this.Content.length, oParagraph);
				}

				oRun = new ParaRun(oParagraph, false);
				oParagraph.AddToContent(oParagraph.GetContentLength() - 1, oRun);
			}
			oParagraph.SetReviewType(reviewtype_Common);

			oRun.SetPr(oTextPr);
			oRun.AddText(sText);
		}
	}
	else
	{
		if (null === sText)
		{
			this.private_ReplaceContentWithPlaceHolder();
		}
		else
		{
			this.private_ReplacePlaceHolderWithContent();
			var oRun = this.private_UpdateListContent();
			if (oRun)
				oRun.AddText(sText);
		}
	}
};
CBlockLevelSdt.prototype.private_UpdateListContent = function()
{
	if (this.IsPlaceHolder())
		return null;

	var oParagraph = this.Content.MakeSingleParagraphContent();
	if (!oParagraph)
		return null;

	var oRun = oParagraph.MakeSingleRunParagraph();
	if (!oRun)
		return null;

	return oRun;
};
/**
 * Проверяем является ли данный контейнер специальным для даты
 * @returns {boolean}
 */
CBlockLevelSdt.prototype.IsDatePicker = function()
{
	return (undefined !== this.Pr.Date);
};
/**
 * @param oPr {AscWord.CSdtDatePickerPr}
 */
CBlockLevelSdt.prototype.SetDatePickerPr = function(oPr)
{
	if (undefined === this.Pr.Date || !this.Pr.Date.IsEqual(oPr))
	{
		var _oPr = oPr ? oPr.Copy() : undefined;
		History.Add(new CChangesSdtPrDatePicker(this, this.Pr.Date, _oPr));
		this.Pr.Date = _oPr;
	}
};
/**
 * @returns {?AscWord.CSdtDatePickerPr}
 */
CBlockLevelSdt.prototype.GetDatePickerPr = function()
{
	return this.Pr.Date;
};
/**
 * Применяем к данному контейнеру настройки того, что это специальный контйенер для даты
 * @param oPr {AscWord.CSdtDatePickerPr}
 * @param updateValue {boolean}
 */
CBlockLevelSdt.prototype.ApplyDatePickerPr = function(oPr, updateValue)
{
	this.SetDatePickerPr(oPr);

	if (!this.IsDatePicker())
		return;

	if (true === updateValue || !this.IsPlaceHolder())
		this.private_UpdateDatePickerContent();
	else
		this.private_FillPlaceholderContent();
};
CBlockLevelSdt.prototype.private_UpdateDatePickerContent = function()
{
	if (!this.Pr.Date)
		return;

	if (this.IsPlaceHolder())
		return this.ReplacePlaceHolderWithContent();

	var oRun;
	var sText = this.Pr.Date.ToString();

	if (this.LogicDocument && this.LogicDocument.IsTrackRevisions())
	{
		if (!sText && this.IsPlaceHolder())
			return;

		var oFirstParagraph = this.GetFirstParagraph();
		var oFirstRun       = oFirstParagraph ? oFirstParagraph.GetFirstRun() : null;
		var oTextPr         = oFirstRun ? oFirstRun.GetDirectTextPr() : new CTextPr();

		if (!this.IsPlaceHolder())
		{
			this.SelectAll();
			this.Remove();
			this.RemoveSelection();
		}
		else
		{
			this.ReplacePlaceHolderWithContent();
		}

		if (!sText && this.IsEmpty())
		{
			this.ReplaceContentWithPlaceHolder();
		}

		if (sText)
		{
			var oRun, oParagraph;
			if (this.IsEmpty())
			{
				oParagraph = this.Content.MakeSingleParagraphContent();
				if (!oParagraph)
					return;

				oRun = oParagraph.MakeSingleRunParagraph();
				if (!oRun)
					return;

				oRun.SetReviewType(reviewtype_Add);
			}
			else
			{
				if (this.Content.Content[this.Content.Content.length - 1].IsParagraph())
				{
					oParagraph = this.Content.Content[this.Content.Content.length - 1];
				}
				else
				{
					oParagraph = new Paragraph(this.LogicDocument.GetDrawingDocument(), this.Content);
					this.Content.AddToParagraph(this.Content.length, oParagraph);
				}

				oRun = new ParaRun(oParagraph, false);
				oParagraph.AddToContent(oParagraph.GetContentLength() - 1, oRun);
			}
			oParagraph.SetReviewType(reviewtype_Common);

			oRun.SetPr(oTextPr);
		}
	}
	else
	{
		var oParagraph = this.Content.MakeSingleParagraphContent();
		if (!oParagraph)
			return;

		oRun = oParagraph.MakeSingleRunParagraph();
		if (!oRun)
			return;
	}

	if (oRun)
		oRun.AddText(sText);
};
CBlockLevelSdt.prototype.Document_Is_SelectionLocked = function(CheckType, bCheckInner)
{
	if (AscCommon.changestype_Document_Content_Add === CheckType
		&& ((this.Content.IsCursorAtBegin() && !this.Get_DocumentPrev())
			|| (this.Content.IsCursorAtEnd() && !this.Get_DocumentNext())))
	{
		return AscCommon.CollaborativeEditing.Add_CheckLock(false);
	}

	var isCheckContentControlLock = this.LogicDocument ? this.LogicDocument.IsCheckContentControlsLock() : true;

	if (AscCommon.changestype_Paragraph_TextProperties === CheckType
		|| ((AscCommon.changestype_Drawing_Props === CheckType || AscCommon.changestype_Image_Properties === CheckType)
		&& this.IsPicture()))
	{
		this.SkipSpecialContentControlLock(true);
		if (!this.CanBeEdited())
			AscCommon.CollaborativeEditing.Add_CheckLock(true);
		this.SkipSpecialContentControlLock(false);

		isCheckContentControlLock = false;
	}

	var nContentControlLock = this.GetContentControlLock();

	if (AscCommon.changestype_ContentControl_Properties === CheckType)
		return this.Lock.Check(this.GetId());

	if (AscCommon.changestype_ContentControl_Remove === CheckType)
	{
		isCheckContentControlLock = false;
		this.Lock.Check(this.GetId());
	}

	if (isCheckContentControlLock
		&& (AscCommon.changestype_Paragraph_Content === CheckType
			|| AscCommon.changestype_Paragraph_AddText === CheckType
			|| AscCommon.changestype_ContentControl_Add === CheckType
			|| AscCommon.changestype_Remove === CheckType
			|| AscCommon.changestype_Delete === CheckType
			|| AscCommon.changestype_Document_Content === CheckType
			|| AscCommon.changestype_Document_Content_Add === CheckType
			|| AscCommon.changestype_ContentControl_Remove === CheckType)
		&& ((this.IsSelectionUse()
			&& this.IsSelectedAll())
			|| this.IsApplyToAll()))
	{
		var bSelectedOnlyThis = false;
		// Если это происходит на добавлении текста, тогда проверяем, что выделен только данный элемент

		if (AscCommon.changestype_Remove !== CheckType && AscCommon.changestype_Delete !== CheckType)
		{
			var oInfo = this.LogicDocument.GetSelectedElementsInfo();
			bSelectedOnlyThis = oInfo.GetBlockLevelSdt() === this ? true : false;
		}

		if (c_oAscSdtLockType.SdtContentLocked === nContentControlLock
			|| (c_oAscSdtLockType.SdtLocked === nContentControlLock && true !== bSelectedOnlyThis)
			|| (!this.CanBeEdited() && true === bSelectedOnlyThis))
		{
			return AscCommon.CollaborativeEditing.Add_CheckLock(true);
		}
		else
		{
			AscCommon.CollaborativeEditing.AddContentControlForSkippingOnCheckEditingLock(this);
			this.Content.Document_Is_SelectionLocked(CheckType, bCheckInner);
			AscCommon.CollaborativeEditing.RemoveContentControlForSkippingOnCheckEditingLock(this);
			return;
		}
	}
	else if (isCheckContentControlLock
		&& !this.CanBeEdited())
	{
		return AscCommon.CollaborativeEditing.Add_CheckLock(true);
	}
	else
	{
		return this.Content.Document_Is_SelectionLocked(CheckType, bCheckInner);
	}
};
CBlockLevelSdt.prototype.CheckContentControlEditingLock = function()
{
	var isCheckContentControlLock = this.LogicDocument ? this.LogicDocument.IsCheckContentControlsLock() : true;
	if (!isCheckContentControlLock)
		return;

	var nContentControlLock = this.GetContentControlLock();

	if (false === AscCommon.CollaborativeEditing.IsNeedToSkipContentControlOnCheckEditingLock(this)
		&& (c_oAscSdtLockType.SdtContentLocked === nContentControlLock || c_oAscSdtLockType.ContentLocked === nContentControlLock))
		return AscCommon.CollaborativeEditing.Add_CheckLock(true);

	if (this.Parent && this.Parent.CheckContentControlEditingLock)
		this.Parent.CheckContentControlEditingLock();
};
/**
 * Получаем типа данного контейнера
 * @returns {Asc.c_oAscContentControlSpecificType}
 */
CBlockLevelSdt.prototype.GetSpecificType = function()
{
	if (this.IsBuiltInTableOfContents())
		return Asc.c_oAscContentControlSpecificType.TOC;

	if (this.IsCheckBox())
		return Asc.c_oAscContentControlSpecificType.CheckBox;

	if (this.IsPicture())
		return Asc.c_oAscContentControlSpecificType.Picture;

	if (this.IsComboBox())
		return Asc.c_oAscContentControlSpecificType.ComboBox;

	if (this.IsDropDownList())
		return Asc.c_oAscContentControlSpecificType.DropDownList;

	if (this.IsDatePicker())
		return Asc.c_oAscContentControlSpecificType.DateTime;

	return Asc.c_oAscContentControlSpecificType.None;
};
CBlockLevelSdt.prototype.GetAllTablesOnPage = function(nPageAbs, arrTables)
{
	return this.Content.GetAllTablesOnPage(nPageAbs, arrTables);
};
CBlockLevelSdt.prototype.ProcessComplexFields = function()
{
	return this.Content.ProcessComplexFields();
};
CBlockLevelSdt.prototype.RecalculateEndInfo = function()
{
	this.Content.RecalculateEndInfo();
};
CBlockLevelSdt.prototype.SetPlaceholder = function(sDocPartName)
{
	if (this.Pr.Placeholder !== sDocPartName)
	{
		History.Add(new CChangesSdtPrPlaceholder(this, this.Pr.Placeholder, sDocPartName));
		this.Pr.Placeholder = sDocPartName;
	}
};
CBlockLevelSdt.prototype.GetPlaceholder = function()
{
	return this.Pr.Placeholder;
};
CBlockLevelSdt.prototype.SetShowingPlcHdr = function(isShow)
{
	if (this.Pr.ShowingPlcHdr !== isShow)
	{
		History.Add(new CChangesSdtPrShowingPlcHdr(this, this.Pr.ShowingPlcHdr, isShow));
		this.Pr.ShowingPlcHdr = isShow;
	}
};
CBlockLevelSdt.prototype.IsShowingPlcHdr = function()
{
	return this.Pr.ShowingPlcHdr;
};
CBlockLevelSdt.prototype.GetFramePr = function()
{
	return this.Content.GetFramePr();
};
CBlockLevelSdt.prototype.SetCalculatedFrame = function(oFrame)
{
	this.Content.SetCalculatedFrame(oFrame);
};
CBlockLevelSdt.prototype.GetPrevParagraphForLineNumbers = function(isPrev, isNewSection)
{
	var oPrevPara = null;

	if (!isPrev)
		oPrevPara = this.Content.GetPrevParagraphForLineNumbers(-1, isNewSection);

	if (!oPrevPara)
	{
		var oParent = this.GetParent();
		if (oParent && oParent.GetPrevParagraphForLineNumbers)
			return oParent.GetPrevParagraphForLineNumbers(this.GetIndex(), isNewSection);
	}

	return oPrevPara;
};
CBlockLevelSdt.prototype.UpdateLineNumbersInfo = function()
{
	this.Content.UpdateLineNumbersInfo();
};
CBlockLevelSdt.prototype.private_OnAddFormPr = function()
{
	this.Content.Recalc_AllParagraphs_CompiledPr();
};
CBlockLevelSdt.prototype.CheckHitInContentControlByXY = function(X, Y, nPageAbs)
{
	var oTransform = this.Get_ParentTextTransform();

	var _X = X;
	var _Y = Y;
	if (oTransform)
	{
		oTransform = oTransform.Invert();
		_X = oTransform.TransformPointX(X, Y);
		_Y = oTransform.TransformPointY(X, Y);
	}

	let logicDocument = this.GetLogicDocument();
	let isFullWidth   = !!(logicDocument && logicDocument.IsFillingFormMode());
	for (var nPageIndex = 0, nPagesCount = this.GetPagesCount(); nPageIndex < nPagesCount; ++nPageIndex)
	{
		if (this.IsEmptyPage(nPageIndex))
			continue;

		if (this.GetAbsolutePage(nPageIndex) === nPageAbs)
		{
			let contentBounds = this.Content.GetContentBounds(nPageIndex);
			let pageBounds    = this.Content.GetPageBounds(nPageIndex);
			
			if ((isFullWidth && pageBounds.Left <= _X && _X <= pageBounds.Right && contentBounds.Top <= _Y && _Y <= contentBounds.Bottom)
				|| (!isFullWidth && contentBounds.Left <= _X && _X <= contentBounds.Right && contentBounds.Top <= _Y && _Y <= contentBounds.Bottom))
				return true;
		}
	}

	return false;
};
CBlockLevelSdt.prototype.CalculateTextToTable = function(oEngine)
{
	this.Content.CalculateTextToTable(oEngine);
};
CBlockLevelSdt.prototype.CollectSelectedReviewChanges = function(oTrackManager)
{
	return this.Content.CollectSelectedReviewChanges(oTrackManager);
};
CBlockLevelSdt.prototype.MoveCursorOutsideForm = function(isBefore)
{
	let logicDocument = this.GetLogicDocument();
	logicDocument.RemoveSelection();
	
	if (isBefore)
	{
		let prevElement = this.GetPrevDocumentElement();
		if (prevElement && prevElement.Document_SetThisElementCurrent)
		{
			prevElement.Document_SetThisElementCurrent();
			prevElement.MoveCursorToEndPos();
		}
	}
	else
	{
		let nextElement = this.GetNextDocumentElement();
		if (nextElement && nextElement.Document_SetThisElementCurrent)
		{
			nextElement.Document_SetThisElementCurrent();
			nextElement.MoveCursorToStartPos();
		}
	}
};
CBlockLevelSdt.prototype.OnContentChange = function()
{
	let logicDocument = this.GetLogicDocument();
	if (logicDocument)
		logicDocument.OnChangeContentControl(this);
	
	if (this.Parent && this.Parent.OnContentChange)
		this.Parent.OnContentChange();
};
//--------------------------------------------------------export--------------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].CBlockLevelSdt = CBlockLevelSdt;
window['AscCommonWord'].type_BlockLevelSdt = type_BlockLevelSdt;
window["AscWord"].CBlockLevelSdt = CBlockLevelSdt;
