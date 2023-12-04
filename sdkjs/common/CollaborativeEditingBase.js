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

(function(window, undefined){

    var FOREIGN_CURSOR_LABEL_HIDETIME = 1500;

    function CCollaborativeChanges()
    {
        this.m_pData  = null;
        this.m_oColor = null;
    }
    CCollaborativeChanges.prototype.Set_Data = function(pData)
    {
        this.m_pData = pData;
    };
    CCollaborativeChanges.prototype.Set_Color = function(oColor)
    {
        this.m_oColor = oColor;
    };
    CCollaborativeChanges.prototype.Set_FromUndoRedo = function(Class, Data, Binary)
    {
        if (!Class.Get_Id)
            return false;

        this.m_pData = this.private_SaveData(Binary);
        return true;
    };
    CCollaborativeChanges.prototype.Apply_Data = function()
    {
        var CollaborativeEditing = AscCommon.CollaborativeEditing;

        var Reader  = this.private_LoadData(this.m_pData);
        if ((Asc.editor || editor).binaryChanges) {
            Reader.GetULong();
        }
        var ClassId = Reader.GetString2();
        var Class   = AscCommon.g_oTableId.Get_ById(ClassId);

        if (!Class)
            return false;

        //------------------------------------------------------------------------------------------------------------------
        // Новая схема
        var nReaderPos   = Reader.GetCurPos();
        var nChangesType = Reader.GetLong();

        var fChangesClass = AscDFH.changesFactory[nChangesType];
        if (fChangesClass)
        {
            var oChange = new fChangesClass(Class);
            oChange.ReadFromBinary(Reader);

            if (true === CollaborativeEditing.private_AddOverallChange(oChange))
            {
                oChange.Load(this.m_oColor);

                if (oChange.GetClass() && oChange.GetClass().SetIsRecalculated && oChange.IsNeedRecalculate())
                    oChange.GetClass().SetIsRecalculated(false);

            }

            return true;
        }
        else
        {
            CollaborativeEditing.private_AddOverallChange(this.m_pData);
            // Сюда мы попадаем, когда у данного изменения нет класса и он все еще работает по старой схеме через объект

            Reader.Seek2(nReaderPos);
            //\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            // Старая схема

            if (!Class.Load_Changes)
                return false;

            return Class.Load_Changes(Reader, null, this.m_oColor);
        }
    };
    CCollaborativeChanges.prototype.private_LoadData = function(szSrc)
    {
        return this.GetStream(szSrc);
    };
    CCollaborativeChanges.prototype.GetStream = function(szSrc, offset)
    {
        if ((Asc.editor || editor).binaryChanges) {
            return new AscCommon.FT_Stream2(szSrc, szSrc.length);
        } else {
            var memoryData = AscCommon.Base64.decode(szSrc, true, undefined, offset);
            return new AscCommon.FT_Stream2(memoryData, memoryData.length);
        }
    };
    CCollaborativeChanges.prototype.private_SaveData = function(Binary)
    {
        var Writer = AscCommon.History.BinaryWriter;
        var Pos    = Binary.Pos;
        var Len    = Binary.Len;
        if ((Asc.editor || editor).binaryChanges) {
            return Writer.GetDataUint8(Pos, Len);
        } else {
            return Len + ";" + Writer.GetBase64Memory2(Pos, Len);
        }
    };
	CCollaborativeChanges.prototype.ToHistoryChange = function()
	{
		let binaryReader = this.private_LoadData(this.m_pData);

		let objectId = binaryReader.GetString2();
		let object   = AscCommon.g_oTableId.Get_ById(objectId);

		if (!object)
		{
			let changeType = binaryReader.GetLong();
			let change = new AscDFH.changesFactory[changeType](object);
			change.ReadFromBinary(binaryReader);

			return null;
		}

		let changeType = binaryReader.GetLong();

		if (!AscDFH.changesFactory[changeType])
			return null;

		let change = new AscDFH.changesFactory[changeType](object);
		change.ReadFromBinary(binaryReader);
		return change;
	};
	CCollaborativeChanges.ToBase64 = function(pos, len)
	{
		let writer = AscCommon.History.BinaryWriter;
		if ((Asc.editor || editor).binaryChanges)
			return writer.GetDataUint8(pos, len);
		else
			return len + ";" + writer.GetBase64Memory2(pos, len);
	};


	/**
	 * Базовый класс для совместного редактирования в разных редакторах
	 * @constructor
	 */
    function CCollaborativeEditingBase()
    {
        this.m_nUseType     = 1;  // 1 - 1 клиент и мы сохраняем историю, -1 - несколько клиентов, 0 - переход из -1 в 1

        this.m_aUsers       = []; // Список текущих пользователей, редактирующих данный документ
        this.m_aChanges     = []; // Массив с изменениями других пользователей

        this.m_aNeedUnlock  = []; // Массив со списком залоченных объектов(которые были залочены другими пользователями)
        this.m_aNeedUnlock2 = []; // Массив со списком залоченных объектов(которые были залочены на данном клиенте)
        this.m_aNeedLock    = []; // Массив со списком залоченных объектов(которые были залочены, но еще не были добавлены на данном клиенте)

        this.m_aLinkData    = []; // Массив, указателей, которые нам надо выставить при загрузке чужих изменений
        this.m_aEndActions  = []; // Массив действий, которые надо выполнить после принятия чужих изменений


        this.m_bGlobalLock          = 0; // Запрещаем производить любые "редактирующие" действия (т.е. то, что в историю запишется)
        this.m_bGlobalLockSelection = 0; // Запрещаем изменять селект и курсор

        this.m_aCheckLocks         = []; // Массив для проверки залоченности объектов, которые мы собираемся изменять
        this.m_aCheckLocksInstance = []; // Массив для проверки залоченности объектов в случае сложного действия

        this.m_aNewObjects  = []; // Массив со списком чужих новых объектов
        this.m_aNewImages   = []; // Массив со списком картинок, которые нужно будет загрузить на сервере
        this.m_aDC          = {}; // Массив(ассоциативный) классов DocumentContent
        this.m_aChangedClasses = {}; // Массив(ассоциативный) классов, в которых есть изменения выделенные цветом

        this.m_aCursorsToUpdate        = {}; // Курсоры, которые нужно обновить после принятия изменений
        this.m_aCursorsToUpdateShortId = {};

        // // CollaborativeEditing LOG
        // this.m_nErrorLog_PointChangesCount = 0;
        // this.m_nErrorLog_SavedPCC          = 0;
        // this.m_nErrorLog_CurPointIndex     = -1;
        // this.m_nErrorLog_SumIndex          = 0;

        this.m_bFast  = false;

        this.m_oLogicDocument     = null;
        this.m_aDocumentPositions = new CDocumentPositionsManager();
        this.m_aForeignCursorsPos = new CDocumentPositionsManager();
        this.m_aForeignCursors    = {};
        this.m_aForeignCursorsId  = {};
        this.m_aForeignCursorsXY     = {};
        this.m_aForeignCursorsToShow = {};


        this.m_nAllChangesSavedIndex = 0;

        this.m_nRecalcIndexStart  = 0;
        this.m_nRecalcIndexEnd    = 0;

		this.CoHistory = new AscCommon.CCollaborativeHistory(this);

        this.m_aAllChanges        = []; // Список всех изменений
        this.m_aOwnChangesIndexes = []; // Список номеров своих изменений в общем списке, которые мы можем откатить

        this.m_oOwnChanges        = [];

		this.m_fEndLoadCallBack   = null;

		this.m_aExternalChanges = [];
		this.m_nExternalIndex0  = -1;
		this.m_nExternalIndex1  = -1;
		this.m_nExternalCounter = 0;
    }

    CCollaborativeEditingBase.prototype.GetEditorApi = function()
    {
        return null;
    };
    CCollaborativeEditingBase.prototype.GetDrawingDocument = function()
    {
        return null;
    };
	CCollaborativeEditingBase.prototype.GetLogicDocument = function()
	{
		return this.m_oLogicDocument;
	};
    CCollaborativeEditingBase.prototype.Clear = function()
    {
        this.m_nUseType = 1;

        this.m_aUsers = [];
        this.m_aChanges = [];
        this.m_aNeedUnlock = [];
        this.m_aNeedUnlock2 = [];
        this.m_aNeedLock = [];
        this.m_aLinkData = [];
        this.m_aEndActions = [];
        this.m_aCheckLocks = [];
        this.m_aCheckLocksInstance = [];
        this.m_aNewObjects = [];
        this.m_aNewImages = [];
    };
    CCollaborativeEditingBase.prototype.Set_Fast = function(bFast)
    {
        this.m_bFast = bFast;

        if (false === bFast)
        {
            this.Remove_AllForeignCursors();
            this.RemoveMyCursorFromOthers();
        }
    };
    CCollaborativeEditingBase.prototype.Is_Fast = function()
    {
        return this.m_bFast;
    };
    CCollaborativeEditingBase.prototype.Is_SingleUser = function()
    {
        return (1 === this.m_nUseType);
    };
	CCollaborativeEditingBase.prototype.getCoHistory = function()
	{
		return this.CoHistory;
	};
    CCollaborativeEditingBase.prototype.getCollaborativeEditing = function()
    {
        return !this.Is_SingleUser();
    };
    CCollaborativeEditingBase.prototype.Start_CollaborationEditing = function()
    {
        this.m_nUseType = -1;
    };
    CCollaborativeEditingBase.prototype.End_CollaborationEditing = function()
    {
        if (this.m_nUseType <= 0)
            this.m_nUseType = 0;
    };
    CCollaborativeEditingBase.prototype.Add_User = function(UserId)
    {
        if (-1 === this.Find_User(UserId))
            this.m_aUsers.push(UserId);
    };
    CCollaborativeEditingBase.prototype.Find_User = function(UserId)
    {
        var Len = this.m_aUsers.length;
        for (var Index = 0; Index < Len; Index++)
        {
            if (this.m_aUsers[Index] === UserId)
                return Index;
        }

        return -1;
    };
    CCollaborativeEditingBase.prototype.Remove_User = function(UserId)
    {
        var Pos = this.Find_User( UserId );
        if ( -1 != Pos )
            this.m_aUsers.splice( Pos, 1 );
    };
    CCollaborativeEditingBase.prototype.Add_Changes = function(Changes)
    {
        this.m_aChanges.push(Changes);
    };
    CCollaborativeEditingBase.prototype.Add_Unlock = function(LockClass)
    {
        this.m_aNeedUnlock.push( LockClass );
    };
    CCollaborativeEditingBase.prototype.Add_Unlock2 = function(Lock)
    {
        this.m_aNeedUnlock2.push(Lock);
        editor._onUpdateDocumentCanSave();
    };
    CCollaborativeEditingBase.prototype.Have_OtherChanges = function()
    {
        return (0 < this.m_aChanges.length);
    };
    CCollaborativeEditingBase.prototype.Apply_Changes = function(fEndCallBack)
    {
        if (this.m_aChanges.length > 0)
        {
            this.GetEditorApi().sendEvent("asc_onBeforeApplyChanges");
            AscFonts.IsCheckSymbols = true;
            editor.WordControl.m_oLogicDocument.PauseRecalculate();
            editor.WordControl.m_oLogicDocument.EndPreview_MailMergeResult();

            editor.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.ApplyChanges);

            var DocState = this.private_SaveDocumentState();
            this.Clear_NewImages();

            this.Apply_OtherChanges();

            // После того как мы приняли чужие изменения, мы должны залочить новые объекты, которые были залочены
            this.Lock_NeedLock();
            this.private_RestoreDocumentState(DocState);
            this.OnStart_Load_Objects(fEndCallBack);
            AscFonts.IsCheckSymbols = false;
            this.GetEditorApi().sendEvent("asc_onApplyChanges");
        }
		else
		{
			if (fEndCallBack)
				fEndCallBack();
		}
    };
    CCollaborativeEditingBase.prototype.Apply_OtherChanges = function()
    {
        // Чтобы заново созданные параграфы не отображались залоченными
        AscCommon.g_oIdCounter.Set_Load( true );

        if (this.m_aChanges.length > 0)
            this.private_CollectOwnChanges();

        this.private_SaveRecalcChangeIndex(true);

        // Применяем изменения, пока они есть
        var _count = this.m_aChanges.length;
        for (var i = 0; i < _count; i++)
        {
            if (window["NATIVE_EDITOR_ENJINE"] === true && window["native"]["CheckNextChange"])
            {
                if (!window["native"]["CheckNextChange"]())
                    break;
            }

            var Changes = this.m_aChanges[i];
            Changes.Apply_Data();
            // // CollaborativeEditing LOG
            // this.m_nErrorLog_PointChangesCount++;
        }

        this.private_SaveRecalcChangeIndex(false);
        this.private_ClearChanges();

        // У новых элементов выставляем указатели на другие классы
        this.Apply_LinkData();

        // Делаем проверки корректности новых изменений
        this.Check_MergeData();

        this.OnEnd_ReadForeignChanges();
        AscCommon.g_oIdCounter.Set_Load( false );
    };
	CCollaborativeEditingBase.prototype.ValidateExternalChanges = function()
	{
		if (this.m_nExternalIndex0 < 0)
		{
			if (true === window['AscApplyChanges'] || !window['AscChanges'])
				return false;

			this.m_aExternalChanges = window['AscChanges'];
			this.m_nExternalIndex0  = 0;
			this.m_nExternalIndex1  = -1;
		}

		return true;
	};
	CCollaborativeEditingBase.prototype.GetNextExternalChange = function()
	{
		if (this.m_nExternalIndex0 < 0
			|| this.m_nExternalIndex0 >= this.m_aExternalChanges.length)
			return null;

		this.m_nExternalIndex1++;

		while (this.m_nExternalIndex1 >= this.m_aExternalChanges[this.m_nExternalIndex0].length)
		{
			this.m_nExternalIndex0++;
			this.m_nExternalIndex1 = 0;

			if (this.m_nExternalIndex0 >= this.m_aExternalChanges.length)
				return null;
		}

		return this.m_aExternalChanges[this.m_nExternalIndex0][this.m_nExternalIndex1];
	};
	/**
	 * Накатываем изменения заданные через LoadExternalChanges, причем не все сразу, а последовательно по 1-ой точке
	 * @returns {number} возвращаем количество примененных изменений
	 */
	CCollaborativeEditingBase.prototype.ApplyExternalChanges = function(pointCount)
	{
		if (!this.ValidateExternalChanges())
			return 0;

		// Чтобы заново созданные параграфы не отображались залоченными
		AscCommon.g_oIdCounter.Set_Load( true );

		if (this.m_aChanges.length > 0)
			this.private_CollectOwnChanges();

		this.private_SaveRecalcChangeIndex(true);

		pointCount = pointCount ? pointCount : 1;
		let counter = 0;
		for (let pointIndex = 0; pointIndex < pointCount; ++pointIndex)
		{
			let color   = new CDocumentColor(191, 255, 199);

			this.m_nExternalCounter++;

			let isStart = false;
			let binary;
			while (binary = this.GetNextExternalChange())
			{
				let change = AscDFH.GetChange(binary);
				if (!change)
				{
					console.log("Lost change!!!");
					continue;
				}

				if (change.IsDescriptionChange())
				{
					if (isStart)
					{
						this.m_nExternalIndex1--;
						break;
					}
					else
						isStart = true;
				}

				if (this.private_AddOverallChange(change))
				{
					change.Load(color);
					if (change.GetClass() && change.GetClass().SetIsRecalculated && change.IsNeedRecalculate())
						change.GetClass().SetIsRecalculated(false);
				}

				counter++;

				// // CollaborativeEditing LOG
				// this.m_nErrorLog_PointChangesCount++;
			}

			if (!binary)
				break;
		}


		this.private_SaveRecalcChangeIndex(false);
		this.private_ClearChanges();

		// У новых элементов выставляем указатели на другие классы
		this.Apply_LinkData();

		// Делаем проверки корректности новых изменений
		this.Check_MergeData();

		this.OnEnd_ReadForeignChanges();

		AscCommon.g_oIdCounter.Set_Load( false );

		editor.WordControl.m_oLogicDocument.RecalculateFromStart();

		return counter;
	};
    CCollaborativeEditingBase.prototype.getOwnLocksLength = function()
    {
        return this.m_aNeedUnlock2.length;
    };
    CCollaborativeEditingBase.prototype.Send_Changes = function()
    {
    };
    CCollaborativeEditingBase.prototype.Release_Locks = function()
    {
    };


    CCollaborativeEditingBase.prototype.CheckWaitingImages = function (aImages) {

    };

    CCollaborativeEditingBase.prototype.SendImagesUrlsFromChanges = function (aImages) {
        var rData = {}, oApi = editor || Asc['editor'], i;
        if(!oApi){
            return;
        }
        rData['type'] = 'pathurls';
        rData['data'] = [];
        for(i = 0; i < aImages.length; ++i)
        {
            rData['data'].push(aImages[i]);
        }
        var aImagesToLoad = [].concat(AscCommon.CollaborativeEditing.m_aNewImages);
        this.CheckWaitingImages(aImagesToLoad);
        AscCommon.CollaborativeEditing.m_aNewImages.length = 0;
        if(false === oApi.isSaveFonts_Images){
            oApi.isSaveFonts_Images = true;
        }
        var callback = function (isTimeout, oRes) {
            var aData, i, oUrls;
            if(oRes && oRes['status'] === 'ok')
            {
                aData = oRes['data'];
                oUrls= {};
                for(i = 0; i < aData.length; ++i)
                {
                    oUrls[aImages[i]] = aData[i];
                }
                AscCommon.g_oDocumentUrls.addUrls(oUrls);
            }
            AscCommon.CollaborativeEditing.SendImagesCallback(aImagesToLoad);
        };

        if (!oApi.CoAuthoringApi.callPRC(rData, Asc.c_nCommonRequestTime, callback)) {
            callback(false, undefined);
        }
    };

    CCollaborativeEditingBase.prototype.SendImagesCallback = function (aImages) {
        var oApi = editor || Asc['editor'];
        oApi.pre_Save(aImages);
    };


    CCollaborativeEditingBase.prototype.CollectImagesFromChanges = function () {
        var oApi = editor || Asc['editor'];
        var aImages = [], sImagePath, i, sImageFromChanges, oThemeUrls = {};
        var aNewImages = this.m_aNewImages;
        var oMap = {};
        for(i = 0; i < aNewImages.length; ++i)
        {
            sImageFromChanges = aNewImages[i];
            if(oMap[sImageFromChanges])
            {
                continue;
            }
            oMap[sImageFromChanges] = 1;
            if(sImageFromChanges.indexOf('theme') === 0 && oApi.ThemeLoader)
            {
                oThemeUrls[sImageFromChanges] = oApi.ThemeLoader.ThemesUrlAbs + sImageFromChanges;
            }
            else if (0 === sImageFromChanges.indexOf('http:') || 0 === sImageFromChanges.indexOf('data:') || 0 === sImageFromChanges.indexOf('https:') ||
                0 === sImageFromChanges.indexOf('file:') || 0 === sImageFromChanges.indexOf('ftp:'))
            {
            }
            else
            {
                sImagePath = AscCommon.g_oDocumentUrls.mediaPrefix + sImageFromChanges;
                if(!AscCommon.g_oDocumentUrls.getUrl(sImagePath))
                {
                    aImages.push(sImagePath);
                }
            }
        }
        AscCommon.g_oDocumentUrls.addUrls(oThemeUrls);
        return aImages;
    };


    CCollaborativeEditingBase.prototype.OnStart_Load_Objects = function(fEndCallBack)
    {
		if (fEndCallBack)
			this.m_fEndLoadCallBack = fEndCallBack;

        this.Set_GlobalLock(true);
        this.Set_GlobalLockSelection(true);
        // Вызываем функцию для загрузки необходимых элементов (новые картинки и шрифты)
        var aImages = this.CollectImagesFromChanges();
        if(aImages.length > 0)
        {
            this.SendImagesUrlsFromChanges(aImages);
        }
        else
        {
            this.SendImagesCallback([].concat(this.m_aNewImages));
            this.m_aNewImages.length = 0;
        }
    };
    CCollaborativeEditingBase.prototype.OnEnd_Load_Objects = function()
    {
		if (this.m_fEndLoadCallBack)
		{
			this.m_fEndLoadCallBack();
			this.m_fEndLoadCallBack = null;
		}
	};
    //-----------------------------------------------------------------------------------
    // Функции для работы с ссылками, у новых объектов
    //-----------------------------------------------------------------------------------
    CCollaborativeEditingBase.prototype.Clear_LinkData = function()
    {
        this.m_aLinkData.length = 0;
    };
    CCollaborativeEditingBase.prototype.Add_LinkData = function(Class, LinkData)
    {
        this.m_aLinkData.push( { Class : Class, LinkData : LinkData } );
    };
    CCollaborativeEditingBase.prototype.Apply_LinkData = function()
    {
        var Count = this.m_aLinkData.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            var Item = this.m_aLinkData[Index];
            Item.Class.Load_LinkData( Item.LinkData );
        }

        this.Clear_LinkData();
    };
    //-----------------------------------------------------------------------------------
    // Функции для проверки корректности новых изменений
    //-----------------------------------------------------------------------------------
    CCollaborativeEditingBase.prototype.Check_MergeData = function()
    {
    };
    //-----------------------------------------------------------------------------------
    // Функции для проверки залоченности объектов
    //-----------------------------------------------------------------------------------
    CCollaborativeEditingBase.prototype.Get_GlobalLock = function()
    {
        return (0 === this.m_bGlobalLock ? false : true);
    };
    CCollaborativeEditingBase.prototype.Set_GlobalLock = function(isLock)
    {
        if (isLock)
            this.m_bGlobalLock++;
        else
            this.m_bGlobalLock = Math.max(0, this.m_bGlobalLock - 1);
    };
    CCollaborativeEditingBase.prototype.Set_GlobalLockSelection = function(isLock)
    {
        if (isLock)
            this.m_bGlobalLockSelection++;
        else
            this.m_bGlobalLockSelection = Math.max(0, this.m_bGlobalLockSelection - 1);
    };
    CCollaborativeEditingBase.prototype.Get_GlobalLockSelection = function()
    {
        return (0 === this.m_bGlobalLockSelection ? false : true);
    };
    CCollaborativeEditingBase.prototype.OnStart_CheckLock = function()
    {
        this.m_aCheckLocks.length = 0;
        this.m_aCheckLocksInstance.length = 0;
    };
    CCollaborativeEditingBase.prototype.Add_CheckLock = function(oItem)
    {
        this.m_aCheckLocks.push(oItem);
        this.m_aCheckLocksInstance.push(oItem);
    };
    CCollaborativeEditingBase.prototype.OnEnd_CheckLock = function()
    {
    };
    CCollaborativeEditingBase.prototype.OnCallback_AskLock = function(result)
    {
    };
    CCollaborativeEditingBase.prototype.OnStartCheckLockInstance = function()
    {
        this.m_aCheckLocksInstance.length = 0;
    };
    CCollaborativeEditingBase.prototype.OnEndCheckLockInstance = function()
    {
        var isLocked = false;
        for (var nIndex = 0, nCount = this.m_aCheckLocksInstance.length; nIndex < nCount; ++nIndex)
        {
            if (true === this.m_aCheckLocksInstance[nIndex])
            {
                isLocked = true;
                break;
            }
        }

        if (isLocked)
        {
            var nCount = this.m_aCheckLocksInstance.length;
            this.m_aCheckLocks.splice(this.m_aCheckLocks.length - nCount, nCount);
        }

        this.m_aCheckLocksInstance.length = 0;
        return isLocked;
    };
    //-----------------------------------------------------------------------------------
    // Функции для работы с залоченными объектами, которые еще не были добавлены
    //-----------------------------------------------------------------------------------
    CCollaborativeEditingBase.prototype.Reset_NeedLock = function()
    {
        this.m_aNeedLock = {};
    };
    CCollaborativeEditingBase.prototype.Add_NeedLock = function(Id, sUser)
    {
        this.m_aNeedLock[Id] = sUser;
    };
    CCollaborativeEditingBase.prototype.Remove_NeedLock = function(Id)
    {
        delete this.m_aNeedLock[Id];
    };
    CCollaborativeEditingBase.prototype.Lock_NeedLock = function()
    {
        for ( var Id in this.m_aNeedLock )
        {
            var Class = AscCommon.g_oTableId.Get_ById( Id );

            if ( null != Class )
            {
                var Lock = Class.Lock;
                Lock.Set_Type( AscCommon.locktype_Other, false );
                if(Class.getObjectType && Class.getObjectType() === AscDFH.historyitem_type_Slide)
                {
                    editor.WordControl.m_oLogicDocument.DrawingDocument.UnLockSlide && editor.WordControl.m_oLogicDocument.DrawingDocument.UnLockSlide(Class.num);
                }
                Lock.Set_UserId( this.m_aNeedLock[Id] );
            }
        }

        this.Reset_NeedLock();
    };
    //-----------------------------------------------------------------------------------
    // Функции для работы с новыми объектами, созданными на других клиентах
    //-----------------------------------------------------------------------------------
    CCollaborativeEditingBase.prototype.Clear_NewObjects = function()
    {
        this.m_aNewObjects.length = 0;
    };
    CCollaborativeEditingBase.prototype.Add_NewObject = function(Class)
    {
        this.m_aNewObjects.push(Class);
        Class.FromBinary = true;
    };
    CCollaborativeEditingBase.prototype.Clear_EndActions = function()
    {
        this.m_aEndActions.length = 0;
    };
    CCollaborativeEditingBase.prototype.Add_EndActions = function(Class, Data)
    {
        this.m_aEndActions.push({Class : Class, Data : Data});
    };
    CCollaborativeEditingBase.prototype.OnEnd_ReadForeignChanges = function()
    {
        var Count = this.m_aNewObjects.length;

        for (var Index = 0; Index < Count; Index++)
        {
            var Class = this.m_aNewObjects[Index];
            Class.FromBinary = false;
        }

        Count = this.m_aEndActions.length;
        for (var Index = 0; Index < Count; Index++)
        {
            var Item = this.m_aEndActions[Index];
            Item.Class.Process_EndLoad(Item.Data);
        }

        this.Clear_EndActions();
        this.Clear_NewObjects();
    };
    //-----------------------------------------------------------------------------------
    // Функции для работы с новыми объектами, созданными на других клиентах
    //-----------------------------------------------------------------------------------
    CCollaborativeEditingBase.prototype.Clear_NewImages = function()
    {
        this.m_aNewImages.length = 0;
    };
    CCollaborativeEditingBase.prototype.Add_NewImage = function(Url)
    {
        this.m_aNewImages.push( Url );
    };
    //-----------------------------------------------------------------------------------
    // Функции для работы с массивом m_aDC
    //-----------------------------------------------------------------------------------
    CCollaborativeEditingBase.prototype.Add_NewDC = function(Class)
    {
        var Id = Class.Get_Id();
        this.m_aDC[Id] = Class;
    };
    CCollaborativeEditingBase.prototype.Clear_DCChanges = function()
    {
        for (var Id in this.m_aDC)
        {
            this.m_aDC[Id].Clear_ContentChanges();
        }

        // Очищаем массив
        this.m_aDC = {};
    };
    CCollaborativeEditingBase.prototype.Refresh_DCChanges = function()
    {
        for (var Id in this.m_aDC)
        {
            this.m_aDC[Id].Refresh_ContentChanges();
        }

        this.Clear_DCChanges();
    };


    //-----------------------------------------------------------------------------------
    // Функции для работы с массивами PosExtChangesX, PosExtChangesY
    //-----------------------------------------------------------------------------------
        CCollaborativeEditingBase.prototype.AddPosExtChanges = function(Item, ChangeObject){

        };
        CCollaborativeEditingBase.prototype.RefreshPosExtChanges = function(){

        };
        CCollaborativeEditingBase.prototype.RewritePosExtChanges = function(changesArr, scale, Binary_Writer)
        {
        };

        CCollaborativeEditingBase.prototype.RefreshPosExtChanges = function()
        {
        };
    //-----------------------------------------------------------------------------------
    // Функции для работы с отметками изменений
    //-----------------------------------------------------------------------------------
    CCollaborativeEditingBase.prototype.Add_ChangedClass = function(Class)
    {
        var Id = Class.Get_Id();
        this.m_aChangedClasses[Id] = Class;
    };
    CCollaborativeEditingBase.prototype.Clear_CollaborativeMarks = function(bRepaint)
    {
        for ( var Id in this.m_aChangedClasses )
        {
            this.m_aChangedClasses[Id].Clear_CollaborativeMarks();
        }

        // Очищаем массив
        this.m_aChangedClasses = {};


        if (true === bRepaint)
        {
            var oDrawingDocument = this.GetDrawingDocument();
            if(oDrawingDocument)
            {
                oDrawingDocument.ClearCachePages();
                oDrawingDocument.FirePaint();
            }
        }
    };
    //----------------------------------------------------------------------------------------------------------------------
    // Функции для работы с обновлением курсоров после принятия изменений
    //----------------------------------------------------------------------------------------------------------------------
	CCollaborativeEditingBase.prototype.UpdateForeignCursorByAdditionalInfo = function(info)
	{
		if (!info)
			return;
		
		let userId      = undefined !== info["UserId"] ? info["UserId"] : info.UserId;
		let cursorInfo  = undefined !== info["CursorInfo"] ? info["CursorInfo"] : info.CursorInfo;
		let shortUserId = undefined !== info["UserShortId"] ? info["UserShortId"] : info.UserShortId;
		
		if (!userId || !cursorInfo || !shortUserId)
			return;
		
		this.Add_ForeignCursorToUpdate(userId, cursorInfo, shortUserId);
	};
    CCollaborativeEditingBase.prototype.Add_ForeignCursorToUpdate = function(UserId, CursorInfo, UserShortId)
    {
        this.m_aCursorsToUpdate[UserId] = CursorInfo;
        this.m_aCursorsToUpdateShortId[UserId] = UserShortId;
    };
    CCollaborativeEditingBase.prototype.private_UpdateForeignCursor = function(CursorInfo, UserId, Show, UserShortId)
    {
    };
    CCollaborativeEditingBase.prototype.Refresh_ForeignCursors = function()
    {
        for (var UserId in this.m_aCursorsToUpdate)
        {
            var CursorInfo = this.m_aCursorsToUpdate[UserId];
            this.private_UpdateForeignCursor(CursorInfo, UserId, false, this.m_aCursorsToUpdateShortId[UserId]);
            this.Add_ForeignCursorToShow(UserId);
        }
        this.m_aCursorsToUpdate = {};
        this.m_aCursorsToUpdateShortId = {};
    };
    //----------------------------------------------------------------------------------------------------------------------
    // Функции для работы с сохраненными позициями в Word-документах. Они объявлены в базовом классе, потому что вызываются
    // из общих классов Paragraph, Run, Table. Вообщем, для совместимости.
    //----------------------------------------------------------------------------------------------------------------------
    CCollaborativeEditingBase.prototype.Clear_DocumentPositions = function(){
        this.m_aDocumentPositions.Clear_DocumentPositions();
    };
    CCollaborativeEditingBase.prototype.Add_DocumentPosition = function(DocumentPos){
        this.m_aDocumentPositions.Add_DocumentPosition(DocumentPos);
    };
    CCollaborativeEditingBase.prototype.Add_ForeignCursor = function(UserId, DocumentPos, UserShortId){
        this.m_aForeignCursorsPos.Remove_DocumentPosition(this.m_aCursorsToUpdate[UserId]);
        this.m_aForeignCursors[UserId] = DocumentPos;
        this.m_aForeignCursorsPos.Add_DocumentPosition(DocumentPos);
        this.m_aForeignCursorsId[UserId] = UserShortId;
    };
    CCollaborativeEditingBase.prototype.Remove_ForeignCursor = function(UserId){
        this.m_aForeignCursorsPos.Remove_DocumentPosition(this.m_aCursorsToUpdate[UserId]);
        delete this.m_aForeignCursors[UserId];
    };
    CCollaborativeEditingBase.prototype.Remove_AllForeignCursors = function() {
        for (var UserId in this.m_aForeignCursors) {
            this.Remove_ForeignCursor(UserId);
        }
    };
    CCollaborativeEditingBase.prototype.RemoveMyCursorFromOthers = function() {
        // Удаляем свой курсор у других пользователей
        var oApi = this.GetEditorApi();
        if(!oApi) {
            return;
        }
        oApi.CoAuthoringApi.sendCursor("");
    };
    CCollaborativeEditingBase.prototype.Update_DocumentPositionsOnAdd = function(Class, Pos){
        this.m_aDocumentPositions.Update_DocumentPositionsOnAdd(Class, Pos);
        this.m_aForeignCursorsPos.Update_DocumentPositionsOnAdd(Class, Pos);
    };
    CCollaborativeEditingBase.prototype.Update_DocumentPositionsOnRemove = function(Class, Pos, Count){
        this.m_aDocumentPositions.Update_DocumentPositionsOnRemove(Class, Pos, Count);
        this.m_aForeignCursorsPos.Update_DocumentPositionsOnRemove(Class, Pos, Count);
    };
    CCollaborativeEditingBase.prototype.OnStart_SplitRun = function(SplitRun, SplitPos){
        this.m_aDocumentPositions.OnStart_SplitRun(SplitRun, SplitPos);
        this.m_aForeignCursorsPos.OnStart_SplitRun(SplitRun, SplitPos);
    };
    CCollaborativeEditingBase.prototype.OnEnd_SplitRun = function(NewRun){
        this.m_aDocumentPositions.OnEnd_SplitRun(NewRun);
        this.m_aForeignCursorsPos.OnEnd_SplitRun(NewRun);
    };
    CCollaborativeEditingBase.prototype.Update_DocumentPosition = function(DocPos){
        this.m_aDocumentPositions.Update_DocumentPosition(DocPos);
    };
    CCollaborativeEditingBase.prototype.Update_ForeignCursorPosition = function(UserId, Run, InRunPos, isRemoveLabel) {
    };
    CCollaborativeEditingBase.prototype.Update_ForeignCursorsPositions = function(){
    };
    CCollaborativeEditingBase.prototype.Check_ForeignCursorsLabels = function(X, Y, PageIndex) {
        var DrawingDocument = this.GetDrawingDocument();
        if(!DrawingDocument) {
            return;
        }
        var Px7 = DrawingDocument.GetMMPerDot(7);
        var Px3 = DrawingDocument.GetMMPerDot(3);

        for (var UserId in this.m_aForeignCursorsXY) {
            var Cursor = this.m_aForeignCursorsXY[UserId];
            if ((true === Cursor.Transform && Cursor.PageIndex === PageIndex && Cursor.X0 - Px3 < X && X < Cursor.X1 + Px3 && Cursor.Y0 - Px3 < Y && Y < Cursor.Y1 + Px3)
                || (Math.abs(X - Cursor.X) < Px7 && Cursor.Y - Px3 < Y && Y < Cursor.Y + Cursor.H + Px3 && Cursor.PageIndex === PageIndex)) {
                this.Show_ForeignCursorLabel(UserId);
            }
        }
    };
    CCollaborativeEditingBase.prototype.Show_ForeignCursorLabel = function(UserId) {
        var Api = this.GetEditorApi();
        if(!Api) {
            return;
        }
        var DrawingDocument = this.GetDrawingDocument();
        if(!DrawingDocument) {
            return;
        }

        if (!this.m_aForeignCursorsXY[UserId])
            return;

        var Cursor = this.m_aForeignCursorsXY[UserId];
        if (Cursor.ShowId)
            clearTimeout(Cursor.ShowId);

        Cursor.ShowId = setTimeout(function()
        {
            Cursor.ShowId = null;
            Api.sendEvent("asc_onHideForeignCursorLabel", UserId);
        }, AscCommon.FOREIGN_CURSOR_LABEL_HIDETIME);

        var UserShortId = this.m_aForeignCursorsId[UserId] ? this.m_aForeignCursorsId[UserId] : UserId;
        var Color  = AscCommon.getUserColorById(UserShortId, null, true);
        var Coords = DrawingDocument.Collaborative_GetTargetPosition(UserId);
        if (!Color || !Coords)
            return;

        this.Update_ForeignCursorLabelPosition(UserId, Coords.X, Coords.Y, Color);
    };
    CCollaborativeEditingBase.prototype.Add_ForeignCursorToShow = function(UserId) {
        this.m_aForeignCursorsToShow[UserId] = true;
    };
    CCollaborativeEditingBase.prototype.Remove_ForeignCursorToShow = function(UserId) {
        delete this.m_aForeignCursorsToShow[UserId];
    };
    CCollaborativeEditingBase.prototype.Add_ForeignCursorXY = function(UserId, X, Y, PageIndex, H, Paragraph, isRemoveLabel, SheetId)
    {
        var Cursor;
        if (!this.m_aForeignCursorsXY[UserId])
        {
            Cursor = {X: X, Y: Y, H: H, PageIndex: PageIndex, Transform: false, ShowId: null, SheetId: SheetId};
            this.m_aForeignCursorsXY[UserId] = Cursor;
        }
        else
        {
            Cursor = this.m_aForeignCursorsXY[UserId];
            if (Cursor.ShowId)
            {
                if (true === isRemoveLabel)
                {
                    var Api = this.GetEditorApi();
                    clearTimeout(Cursor.ShowId);
                    Cursor.ShowId = null;
                    Api.sendEvent("asc_onHideForeignCursorLabel", UserId);
                }
            }
            else
            {
                Cursor.ShowId = null;
            }

            Cursor.X         = X;
            Cursor.Y         = Y;
            Cursor.PageIndex = PageIndex;
            Cursor.H         = H;
            Cursor.SheetId   = SheetId;
        }

        var Transform = Paragraph.Get_ParentTextTransform();
        if (Transform)
        {
            Cursor.Transform = true;
            var X0 = Transform.TransformPointX(Cursor.X, Cursor.Y);
            var Y0 = Transform.TransformPointY(Cursor.X, Cursor.Y);
            var X1 = Transform.TransformPointX(Cursor.X, Cursor.Y + Cursor.H);
            var Y1 = Transform.TransformPointY(Cursor.X, Cursor.Y + Cursor.H);

            Cursor.X0 = Math.min(X0, X1);
            Cursor.Y0 = Math.min(Y0, Y1);
            Cursor.X1 = Math.max(X0, X1);
            Cursor.Y1 = Math.max(Y0, Y1);
        }
        else
        {
            Cursor.Transform = false;
        }
    };
    CCollaborativeEditingBase.prototype.Remove_ForeignCursorXY = function(UserId) {
        if (this.m_aForeignCursorsXY[UserId]) {
            if (this.m_aForeignCursorsXY[UserId].ShowId) {
                var Api = this.GetEditorApi();
                Api.sendEvent("asc_onHideForeignCursorLabel", UserId);
                clearTimeout(this.m_aForeignCursorsXY[UserId].ShowId);
            }
            delete this.m_aForeignCursorsXY[UserId];
        }
    };
    CCollaborativeEditingBase.prototype.private_SaveDocumentState = function() {
        var LogicDocument = editor.WordControl.m_oLogicDocument;
        var DocState;
        if (true !== this.Is_Fast()) {
            DocState = LogicDocument.Get_SelectionState2();
            this.m_aCursorsToUpdate = {};
        }
        else {
            DocState = LogicDocument.Save_DocumentStateBeforeLoadChanges(false, true);
        }
        return DocState;
    };
    CCollaborativeEditingBase.prototype.Remove_ForeignCursorToShow = function(UserId) {
        delete this.m_aForeignCursorsToShow[UserId];
    };
    CCollaborativeEditingBase.prototype.private_RestoreDocumentState = function(DocState) {
        var LogicDocument = editor.WordControl.m_oLogicDocument;
        if (true !== this.Is_Fast()) {
            LogicDocument.Set_SelectionState2(DocState);
        }
        else {
            LogicDocument.Load_DocumentStateAfterLoadChanges(DocState);
            this.Refresh_ForeignCursors();
        }
    };
    CCollaborativeEditingBase.prototype.WatchDocumentPositionsByState = function(DocState)
	{
        this.Clear_DocumentPositions();

        if (DocState.Pos)
            this.Add_DocumentPosition(DocState.Pos);
        if (DocState.StartPos)
            this.Add_DocumentPosition(DocState.StartPos);
        if (DocState.EndPos)
            this.Add_DocumentPosition(DocState.EndPos);

		if (DocState.AnchorPos)
			this.Add_DocumentPosition(DocState.AnchorPos);

        if (DocState.FootnotesStart && DocState.FootnotesStart.Pos)
            this.Add_DocumentPosition(DocState.FootnotesStart.Pos);
        if (DocState.FootnotesStart && DocState.FootnotesStart.StartPos)
            this.Add_DocumentPosition(DocState.FootnotesStart.StartPos);
        if (DocState.FootnotesStart && DocState.FootnotesStart.EndPos)
            this.Add_DocumentPosition(DocState.FootnotesStart.EndPos);
        if (DocState.FootnotesEnd && DocState.FootnotesEnd.Pos)
            this.Add_DocumentPosition(DocState.FootnotesEnd.Pos);
        if (DocState.FootnotesEnd && DocState.FootnotesEnd.StartPos)
            this.Add_DocumentPosition(DocState.FootnotesEnd.StartPos);
        if (DocState.FootnotesEnd && DocState.FootnotesEnd.EndPos)
            this.Add_DocumentPosition(DocState.FootnotesEnd.EndPos);
    };
    CCollaborativeEditingBase.prototype.UpdateDocumentPositionsByState = function(DocState)
	{
        if (DocState.Pos)
            this.Update_DocumentPosition(DocState.Pos);
        if (DocState.StartPos)
            this.Update_DocumentPosition(DocState.StartPos);
        if (DocState.EndPos)
            this.Update_DocumentPosition(DocState.EndPos);

		if (DocState.AnchorPos)
			this.Update_DocumentPosition(DocState.AnchorPos);

        if (DocState.FootnotesStart && DocState.FootnotesStart.Pos)
            this.Update_DocumentPosition(DocState.FootnotesStart.Pos);
        if (DocState.FootnotesStart && DocState.FootnotesStart.StartPos)
            this.Update_DocumentPosition(DocState.FootnotesStart.StartPos);
        if (DocState.FootnotesStart && DocState.FootnotesStart.EndPos)
            this.Update_DocumentPosition(DocState.FootnotesStart.EndPos);
        if (DocState.FootnotesEnd && DocState.FootnotesEnd.Pos)
            this.Update_DocumentPosition(DocState.FootnotesEnd.Pos);
        if (DocState.FootnotesEnd && DocState.FootnotesEnd.StartPos)
            this.Update_DocumentPosition(DocState.FootnotesEnd.StartPos);
        if (DocState.FootnotesEnd && DocState.FootnotesEnd.EndPos)
            this.Update_DocumentPosition(DocState.FootnotesEnd.EndPos);
    };
    CCollaborativeEditingBase.prototype.Update_ForeignCursorLabelPosition = function(UserId, X, Y, Color)
    {
        var oApi = this.GetEditorApi();
        if(!oApi) {
            return;
        }
        var Cursor = this.m_aForeignCursorsXY[UserId];
        if (!Cursor || !Cursor.ShowId)
            return;
        if (this.m_oLogicDocument && this.m_oLogicDocument.IsFocusOnNotes && this.m_oLogicDocument.IsFocusOnNotes())
        {
            Y += parseInt(oApi.WordControl.m_oNotesContainer.HtmlElement.style.top);
        }
        oApi.sendEvent("asc_onShowForeignCursorLabel", UserId, X, Y, new AscCommon.CColor(Color.r, Color.g, Color.b, 255));
    };
    CCollaborativeEditingBase.prototype.GetDocumentPositionBinary = function(oWriter, PosInfo) {
        if (!PosInfo)
            return null;
        var BinaryPos = oWriter.GetCurPosition();
        oWriter.WriteString2(PosInfo.Class.Get_Id());
        oWriter.WriteLong(PosInfo.Position);
        var BinaryLen = oWriter.GetCurPosition() - BinaryPos;
        return  (BinaryLen + ";" + oWriter.GetBase64Memory2(BinaryPos, BinaryLen));
    };
    //----------------------------------------------------------------------------------------------------------------------
    // Private area
    //----------------------------------------------------------------------------------------------------------------------
	CCollaborativeEditingBase.prototype.private_ClearChanges      = function()
	{
		this.m_aChanges    = [];
		this.m_oOwnChanges = [];
	};
	CCollaborativeEditingBase.prototype.private_CollectOwnChanges = function()
	{
		var StartPoint = (null === AscCommon.History.SavedIndex ? 0 : AscCommon.History.SavedIndex + 1);
		var LastPoint  = -1;

		if (this.m_nUseType <= 0)
			LastPoint = AscCommon.History.Points.length - 1;
		else
			LastPoint = AscCommon.History.Index;

		for (var PointIndex = StartPoint; PointIndex <= LastPoint; PointIndex++)
		{
			var Point = AscCommon.History.Points[PointIndex];
			for (var Index = 0; Index < Point.Items.length; Index++)
			{
				var Item = Point.Items[Index];

				this.m_oOwnChanges.push(Item.Data);
			}
		}
	};
	CCollaborativeEditingBase.prototype.private_AddOverallChange  = function(oChange, isSave)
	{
		// Здесь мы должны смержить пришедшее изменение с одним из наших изменений
		for (var nIndex = 0, nCount = this.m_oOwnChanges.length; nIndex < nCount; ++nIndex)
		{
			if (oChange && oChange.Merge && false === oChange.Merge(this.m_oOwnChanges[nIndex]))
				return false;
		}

		if (false !== isSave)
			this.CoHistory.AddChange(oChange);

		return true;
	};
	CCollaborativeEditingBase.prototype.PreUndo = function()
	{
		let logicDocument = this.m_oLogicDocument;

		logicDocument.DrawingDocument.EndTrackTable(null, true);
		logicDocument.TurnOffCheckChartSelection();

		return this.private_SaveDocumentState();
	};
	CCollaborativeEditingBase.prototype.PostUndo = function(state, changes)
	{
		this.private_RestoreDocumentState(state);
		this.private_RecalculateDocument(changes);

		let logicDocument = this.m_oLogicDocument;
		logicDocument.TurnOnCheckChartSelection();
		logicDocument.UpdateSelection();
		logicDocument.UpdateInterface();
		logicDocument.UpdateRulers();
	};
	CCollaborativeEditingBase.prototype.UndoGlobal = function(count)
	{
		if (!count)
			return;

		let state   = this.PreUndo();
		let changes = this.CoHistory.UndoGlobalChanges(count);
		this.PostUndo(state, changes);
	};
	CCollaborativeEditingBase.prototype.UndoGlobalPoint = function()
	{
		let state   = this.PreUndo();
		let changes = this.CoHistory.UndoGlobalPoint();
		this.PostUndo(state, changes);
	};
	CCollaborativeEditingBase.prototype.Undo = function()
	{
		if (true === this.Get_GlobalLock())
			return;

		let state   = this.PreUndo();
		let changes = this.CoHistory.UndoOwnPoint();
		this.PostUndo(state, changes);
	};
	CCollaborativeEditingBase.prototype.CanUndoMultipleActions = function()
	{
		return this.CoHistory.CanUndoGlobal();
	};
	CCollaborativeEditingBase.prototype.GetAllChangesCount = function()
	{
		return this.CoHistory.GetChangeCount();
	};
	CCollaborativeEditingBase.prototype.CanUndo = function()
	{
		return this.CoHistory.CanUndoOwn();
	};
	CCollaborativeEditingBase.prototype.private_RecalculateDocument = function(oRecalcData)
	{
	};
	CCollaborativeEditingBase.prototype.private_SaveRecalcChangeIndex = function(isStart)
	{
		if (isStart)
			this.m_nRecalcIndexStart = this.CoHistory.GetChangeCount();
		else
			this.m_nRecalcIndexEnd = this.CoHistory.GetChangeCount() - 1;
	};
    //----------------------------------------------------------------------------------------------------------------------
    // Класс для работы с сохраненными позициями документа.
    //----------------------------------------------------------------------------------------------------------------------
    //   Принцип следующий. Заданная позиция - это Run + Позиция внутри данного Run.
    //   Если заданный ран был разбит (операция Split), тогда отслеживаем куда перешла
    //   заданная позиция, в новый ран или осталась в старом? Если в новый, тогда сохраняем
    //   новый ран как отдельную позицию в массив m_aDocumentPositions, и добавляем мап
    //   старой позиции в новую m_aDocumentPositionsMap. В конце действия, когда нам нужно
    //   определить где же находистся наша позиция, мы сначала проверяем Map, если в нем есть
    //   конечная позиция, проверяем является ли заданная позиция валидной в документе.
    //   Если да, тогда выставляем ее, если нет, тогда берем Run исходной позиции, и
    //   пытаемся сформировать полную позицию по данному Run. Если и это не получается,
    //   тогда восстанавливаем позицию по измененной полной исходной позиции.
    //----------------------------------------------------------------------------------------------------------------------
    function CDocumentPositionsManager()
    {
        this.m_aDocumentPositions      = [];
        this.m_aDocumentPositionsSplit = [];
        this.m_aDocumentPositionsMap   = [];
    }
    CDocumentPositionsManager.prototype.Clear_DocumentPositions = function()
    {
        this.m_aDocumentPositions      = [];
        this.m_aDocumentPositionsSplit = [];
        this.m_aDocumentPositionsMap   = [];
    };
    CDocumentPositionsManager.prototype.Add_DocumentPosition = function(Position)
    {
        this.m_aDocumentPositions.push(Position);
    };
    CDocumentPositionsManager.prototype.Update_DocumentPositionsOnAdd = function(Class, Pos)
    {
        for (var PosIndex = 0, PosCount = this.m_aDocumentPositions.length; PosIndex < PosCount; ++PosIndex)
        {
            var DocPos = this.m_aDocumentPositions[PosIndex];
            for (var ClassPos = 0, ClassLen = DocPos.length; ClassPos < ClassLen; ++ClassPos)
            {
                var _Pos = DocPos[ClassPos];
                if (Class === _Pos.Class
                    && undefined !== _Pos.Position
                    && (_Pos.Position > Pos
                    || (_Pos.Position === Pos && !(Class instanceof AscCommonWord.ParaRun))))
                {
                    _Pos.Position++;
                    break;
                }
            }
        }
    };
    CDocumentPositionsManager.prototype.Update_DocumentPositionsOnRemove = function(Class, Pos, Count)
    {
        for (var PosIndex = 0, PosCount = this.m_aDocumentPositions.length; PosIndex < PosCount; ++PosIndex)
        {
            var DocPos = this.m_aDocumentPositions[PosIndex];
            for (var ClassPos = 0, ClassLen = DocPos.length; ClassPos < ClassLen; ++ClassPos)
            {
                var _Pos = DocPos[ClassPos];
                if (Class === _Pos.Class && undefined !== _Pos.Position)
                {
                    if (_Pos.Position > Pos + Count)
                    {
                        _Pos.Position -= Count;
                    }
                    else if (_Pos.Position >= Pos)
                    {
                        // Элемент, в котором находится наша позиция, удаляется. Ставим специальную отметку об этом.
                        _Pos.Position = Pos;
                        _Pos.Deleted  = true;
                    }

                    break;
                }
            }
        }
    };
    CDocumentPositionsManager.prototype.OnStart_SplitRun = function(SplitRun, SplitPos)
    {
        this.m_aDocumentPositionsSplit = [];

        for (var PosIndex = 0, PosCount = this.m_aDocumentPositions.length; PosIndex < PosCount; ++PosIndex)
        {
            var DocPos = this.m_aDocumentPositions[PosIndex];
            for (var ClassPos = 0, ClassLen = DocPos.length; ClassPos < ClassLen; ++ClassPos)
            {
                var _Pos = DocPos[ClassPos];
                if (SplitRun === _Pos.Class && _Pos.Position && _Pos.Position >= SplitPos)
                {
                    this.m_aDocumentPositionsSplit.push({DocPos : DocPos, NewRunPos : _Pos.Position - SplitPos});
                }
            }
        }
    };
    CDocumentPositionsManager.prototype.OnEnd_SplitRun = function(NewRun)
    {
        if (!NewRun)
            return;

        for (var PosIndex = 0, PosCount = this.m_aDocumentPositionsSplit.length; PosIndex < PosCount; ++PosIndex)
        {
            var NewDocPos = [];
            NewDocPos.push({Class : NewRun, Position : this.m_aDocumentPositionsSplit[PosIndex].NewRunPos});
            this.m_aDocumentPositions.push(NewDocPos);
            this.m_aDocumentPositionsMap.push({StartPos : this.m_aDocumentPositionsSplit[PosIndex].DocPos, EndPos : NewDocPos});
        }
    };
    CDocumentPositionsManager.prototype.Update_DocumentPosition = function(DocPos)
    {
        // Смотрим куда мапится заданная позиция
        var NewDocPos = DocPos;
        for (var PosIndex = 0, PosCount = this.m_aDocumentPositionsMap.length; PosIndex < PosCount; ++PosIndex)
        {
            if (this.m_aDocumentPositionsMap[PosIndex].StartPos === NewDocPos)
                NewDocPos = this.m_aDocumentPositionsMap[PosIndex].EndPos;
        }

        // Нашли результирующую позицию. Проверим является ли она валидной для документа.
        if (NewDocPos !== DocPos && NewDocPos.length === 1 && NewDocPos[0].Class instanceof AscCommonWord.ParaRun)
        {
            var Run = NewDocPos[0].Class;
            var Para = Run.GetParagraph();
            if (AscCommonWord.CanUpdatePosition(Para, Run))
            {
                DocPos.length = 0;
                Run.GetDocumentPositionFromObject(DocPos);
                DocPos.push({Class : Run, Position : NewDocPos[0].Position});
            }
        }
        // Возможно ран с позицией переместился в другой класс
        else if (DocPos.length > 0 && DocPos[DocPos.length - 1].Class instanceof AscCommonWord.ParaRun)
        {
            var Run = DocPos[DocPos.length - 1].Class;
            var RunPos = DocPos[DocPos.length - 1].Position;
            var Para = Run.GetParagraph();
            if (AscCommonWord.CanUpdatePosition(Para, Run))
            {
                DocPos.length = 0;
                Run.GetDocumentPositionFromObject(DocPos);
                DocPos.push({Class : Run, Position : RunPos});
            }
        }
    };
    CDocumentPositionsManager.prototype.Remove_DocumentPosition = function(DocPos)
    {
        for (var Pos = 0, Count = this.m_aDocumentPositions.length; Pos < Count; ++Pos)
        {
            if (this.m_aDocumentPositions[Pos] === DocPos)
            {
                this.m_aDocumentPositions.splice(Pos, 1);
                return;
            }
        }
    };
    //--------------------------------------------------------export----------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].FOREIGN_CURSOR_LABEL_HIDETIME = FOREIGN_CURSOR_LABEL_HIDETIME;
    window['AscCommon'].CCollaborativeChanges = CCollaborativeChanges;
    window['AscCommon'].CCollaborativeEditingBase = CCollaborativeEditingBase;
    window['AscCommon'].CDocumentPositionsManager = CDocumentPositionsManager;
})(window);
