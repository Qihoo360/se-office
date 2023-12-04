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

/**
 *
 * PageLayout.js
 *
 * PageLayout controller
 *
 * Extra controller for toolbar
 *
 * Created by Maxim.Kadushkin on 3/31/2017.
 */

define([
    'core'
], function () {
    'use strict';

    DE.Controllers.PageLayout = Backbone.Controller.extend((function(){
        var _imgOriginalProps;

        return {
            initialize: function () {
            },

            onLaunch: function (view) {
                this.toolbar = view;
                this.editMode = true;
                this._state = {};
                return this;
            },

            onAppReady: function (config) {
                var me = this;

                me.toolbar.btnImgAlign.menu.on('item:click', me.onClickMenuAlign.bind(me));
                me.toolbar.btnImgAlign.menu.on('show:before', me.onBeforeShapeAlign.bind(me));
                me.toolbar.btnImgWrapping.menu.on('item:click', me.onClickMenuWrapping.bind(me));
                me.toolbar.btnImgGroup.menu.on('item:click', me.onClickMenuGroup.bind(me));
                me.toolbar.btnImgForward.menu.on('item:click', me.onClickMenuForward.bind(me));
                me.toolbar.btnImgBackward.menu.on('item:click', me.onClickMenuForward.bind(me));

                me.toolbar.btnImgForward.on('click', me.onClickMenuForward.bind(me, 'forward'));
                me.toolbar.btnImgBackward.on('click', me.onClickMenuForward.bind(me, 'backward'));

                me.toolbar.btnsPageBreak.forEach( function(btn) {
                    var _menu_section_break = btn.menu.items[2].menu;
                    _menu_section_break.on('item:click', function (menu, item, e) {
                        me.toolbar.fireEvent('insert:break', [item.value]);
                    });

                    btn.menu.on('item:click', function (menu, item, e) {
                        if ( !(item.value == 'section') )
                            me.toolbar.fireEvent('insert:break', [item.value]);
                    });

                    btn.on('click', function(e) {
                        me.toolbar.fireEvent('insert:break', ['page']);
                    });
                });

            },

            setApi: function (api) {
                this.api = api;

                this.api.asc_registerCallback('asc_onImgWrapStyleChanged', this.onApiWrappingStyleChanged.bind(this));
                this.api.asc_registerCallback('asc_onCoAuthoringDisconnect', this.onApiCoAuthoringDisconnect.bind(this));
                this.api.asc_registerCallback('asc_onFocusObject', this.onApiFocusObject.bind(this));
                return this;
            },

            onApiWrappingStyleChanged: function (type) {
                var menu = this.toolbar.btnImgWrapping.menu;

                switch ( type ) {
                case Asc.c_oAscWrapStyle2.Inline:       menu.items[0].setChecked(true); break;
                case Asc.c_oAscWrapStyle2.Square:       menu.items[2].setChecked(true); break;
                case Asc.c_oAscWrapStyle2.Tight:        menu.items[3].setChecked(true); break;
                case Asc.c_oAscWrapStyle2.Through:      menu.items[4].setChecked(true); break;
                case Asc.c_oAscWrapStyle2.TopAndBottom: menu.items[5].setChecked(true); break;
                case Asc.c_oAscWrapStyle2.Behind:       menu.items[8].setChecked(true); break;
                case Asc.c_oAscWrapStyle2.InFront:      menu.items[7].setChecked(true); break;
                default:
                    for (var i in menu.items) {
                        menu.items[i].setChecked( false );
                    }
                }
            },

            onApiFocusObject: function(objects) {
                if (!this.editMode) return;

                var me = this;
                var disable = {}, type,
                    islocked = false,
                    shapeProps,
                    canGroupUngroup = false,
                    wrapping,
                    content_locked = false,
                    no_object = true;

                for (var i in objects) {
                    type = objects[i].get_ObjectType();
                    if ( type === Asc.c_oAscTypeSelectElement.Image ) {
                        var props = objects[i].get_ObjectValue(),
                            notflow = !props.get_CanBeFlow();
                        shapeProps = props.get_ShapeProperties();
                        islocked = props.get_Locked();
                        wrapping = props.get_WrappingStyle();
                        no_object = false;
                        me.onApiWrappingStyleChanged(notflow ? -1 : wrapping);

                        _.each(me.toolbar.btnImgWrapping.menu.items, function(item) {
                            item.setDisabled(notflow);
                        });
                        me.toolbar.btnImgWrapping.menu.items[10].setDisabled(!me.api.CanChangeWrapPolygon());

                        var control_props = me.api.asc_IsContentControl() ? this.api.asc_GetContentControlProperties() : null,
                            lock_type = (control_props) ? control_props.get_Lock() : Asc.c_oAscSdtLockType.Unlocked;

                        content_locked = lock_type==Asc.c_oAscSdtLockType.SdtContentLocked || lock_type==Asc.c_oAscSdtLockType.ContentLocked;
                        disable.arrange     = (wrapping == Asc.c_oAscWrapStyle2.Inline) && !props.get_FromGroup();
                        disable.wrapping    = props.get_FromGroup() || (notflow && !me.api.CanChangeWrapPolygon()) ||
                                            (!!control_props && control_props.get_SpecificType()==Asc.c_oAscContentControlSpecificType.Picture && !control_props.get_FormPr());
                        disable.group   = islocked || wrapping == Asc.c_oAscWrapStyle2.Inline || content_locked;
                        canGroupUngroup = me.api.CanGroup() || me.api.CanUnGroup();
                        if (!disable.group && canGroupUngroup) {
                            me.toolbar.btnImgGroup.menu.items[0].setDisabled(!me.api.CanGroup());
                            me.toolbar.btnImgGroup.menu.items[1].setDisabled(!me.api.CanUnGroup());
                        }

                        _imgOriginalProps = props;
                        break;
                    }
                }
                me.toolbar.lockToolbar(Common.enumLock.noObjectSelected, no_object, {array: [me.toolbar.btnImgAlign, me.toolbar.btnImgGroup, me.toolbar.btnImgWrapping, me.toolbar.btnImgForward, me.toolbar.btnImgBackward]});
                me.toolbar.lockToolbar(Common.enumLock.imageLock, islocked, {array: [me.toolbar.btnImgAlign, me.toolbar.btnImgGroup, me.toolbar.btnImgWrapping]});
                me.toolbar.lockToolbar(Common.enumLock.contentLock, content_locked, {array: [me.toolbar.btnImgAlign, me.toolbar.btnImgGroup, me.toolbar.btnImgWrapping, me.toolbar.btnImgForward, me.toolbar.btnImgBackward]});
                me.toolbar.lockToolbar(Common.enumLock.inImageInline, wrapping == Asc.c_oAscWrapStyle2.Inline, {array: [me.toolbar.btnImgAlign, me.toolbar.btnImgGroup]});
                me.toolbar.lockToolbar(Common.enumLock.inSmartartInternal, shapeProps && shapeProps.asc_getFromSmartArtInternal(), {array: [me.toolbar.btnImgForward, me.toolbar.btnImgBackward]});
                me.toolbar.lockToolbar(Common.enumLock.cantGroup, !canGroupUngroup, {array: [me.toolbar.btnImgGroup]});
                me.toolbar.lockToolbar(Common.enumLock.cantWrap, disable.wrapping, {array: [me.toolbar.btnImgWrapping]});
                me.toolbar.lockToolbar(Common.enumLock.cantArrange, disable.arrange, {array: [me.toolbar.btnImgForward, me.toolbar.btnImgBackward]});
            },

            onApiCoAuthoringDisconnect: function() {
                this.editMode = false;
            },

            onBeforeShapeAlign: function() {
                var value = this.api.asc_getSelectedDrawingObjectsCount(),
                    alignto = Common.Utils.InternalSettings.get("de-img-align-to");
                this.toolbar.mniAlignObjects.setDisabled(value<2);
                this.toolbar.mniAlignObjects.setChecked(value>1 && (!alignto || alignto==3), true);
                this.toolbar.mniAlignToMargin.setChecked((value<2 && !alignto || alignto==2), true);
                this.toolbar.mniAlignToPage.setChecked(alignto==1, true);
                this.toolbar.mniDistribHor.setDisabled(value<3 && this.toolbar.mniAlignObjects.isChecked());
                this.toolbar.mniDistribVert.setDisabled(value<3 && this.toolbar.mniAlignObjects.isChecked());
            },

            onClickMenuAlign: function (menu, item, e) {
                var value = this.toolbar.mniAlignToPage.isChecked() ? Asc.c_oAscObjectsAlignType.Page : (this.toolbar.mniAlignToMargin.isChecked() ? Asc.c_oAscObjectsAlignType.Margin : Asc.c_oAscObjectsAlignType.Selected);
                if (item.value>-1 && item.value < 6) {
                    this.api.put_ShapesAlign(item.value, value);
                    Common.component.Analytics.trackEvent('ToolBar', 'Shape Align');
                } else if (item.value == 6) {
                    this.api.DistributeHorizontally(value);
                    Common.component.Analytics.trackEvent('ToolBar', 'Distribute');
                } else if (item.value == 7){
                    this.api.DistributeVertically(value);
                    Common.component.Analytics.trackEvent('ToolBar', 'Distribute');
                }
                this.toolbar.fireEvent('editcomplete', this.toolbar);
            },

            onClickMenuWrapping: function (menu, item, e) {
                if (item.options.wrapType=='edit') {
                    this.api.StartChangeWrapPolygon();
                    this.toolbar.fireEvent('editcomplete', this.toolbar);
                    return;
                }

                var props = new Asc.asc_CImgProperty();
                props.put_WrappingStyle(item.options.wrapType);

                if ( _imgOriginalProps.get_WrappingStyle() === Asc.c_oAscWrapStyle2.Inline && item.options.wrapType !== Asc.c_oAscWrapStyle2.Inline ) {
                    props.put_PositionH(new Asc.CImagePositionH());
                    props.get_PositionH().put_UseAlign(false);
                    props.get_PositionH().put_RelativeFrom(Asc.c_oAscRelativeFromH.Column);

                    var val = _imgOriginalProps.get_Value_X(Asc.c_oAscRelativeFromH.Column);
                    props.get_PositionH().put_Value(val);

                    props.put_PositionV(new Asc.CImagePositionV());
                    props.get_PositionV().put_UseAlign(false);
                    props.get_PositionV().put_RelativeFrom(Asc.c_oAscRelativeFromV.Paragraph);

                    val = _imgOriginalProps.get_Value_Y(Asc.c_oAscRelativeFromV.Paragraph);
                    props.get_PositionV().put_Value(val);
                }

                this.api.ImgApply(props);
                this.toolbar.fireEvent('editcomplete', this.toolbar);
            },

            onClickMenuGroup: function (menu, item, e) {
                var props = new Asc.asc_CImgProperty();
                props.put_Group(item.options.groupval);

                this.api.ImgApply(props);
                this.toolbar.fireEvent('editcomplete', this.toolbar);
            },

            onClickMenuForward: function (menu, item, e) {
                var props = new Asc.asc_CImgProperty();

                if ( menu == 'forward' )
                    props.put_ChangeLevel(Asc.c_oAscChangeLevel.BringForward); else
                if ( menu == 'backward' )
                    props.put_ChangeLevel(Asc.c_oAscChangeLevel.BringBackward); else
                    props.put_ChangeLevel(item.options.valign);

                this.api.ImgApply(props);
                this.toolbar.fireEvent('editcomplete', this.toolbar);
            }
        }
    })());
});
