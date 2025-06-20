import React from 'react';
import {Link} from 'framework7-react';
import { Device } from "../../../../common/mobile/utils/device";
import SvgIcon from '@common/lib/component/SvgIcon'
import IconRedoForAndroid from '@common-android-icons/icon-redo.svg';
import IconUndoForAndroid from '@common-android-icons/icon-undo.svg';
import IconEditSettingsForAndroid from '@common-android-icons/icon-edit-settings.svg';
import IconPlusForAndroid from '@common-android-icons/icon-plus.svg';
import IconRedoForIos from '@common-ios-icons/icon-redo.svg?ios';
import IconUndoForIos from '@common-ios-icons/icon-undo.svg?ios';
import IconEditSettingsForIos from '@common-ios-icons/icon-edit-settings.svg?ios';
import IconPlusForIos from '@common-ios-icons/icon-plus.svg?ios';
import IconCopy from '@common-icons/icon-copy.svg';
import IconCut from '@common-icons/icon-cut.svg';
import IconPaste from '@common-icons/icon-paste.svg';
const EditorUIController = () => {
    return null
};

EditorUIController.isSupportEditFeature = () => {
    return true
};

EditorUIController.toolbarOptions = {
    getUndoRedo: function (props) {
        const {disabledUndo, disabledRedo, onUndoClick, onRedoClick} = props;
        return (
            <React.Fragment>
                {/* 撤销按钮 */}
                <Link
                    iconOnly={true}
                    className={disabledUndo && "disabled"}
                    onClick={onUndoClick}
                >
                    {Device.ios ? (
                        <SvgIcon
                            symbolId={IconRedoForIos.id}
                            className="icon icon-svg"
                        />
                    ) : (
                        <SvgIcon
                            symbolId={IconRedoForAndroid.id}
                            className="icon icon-svg"
                        />
                    )}
                </Link>

                {/* 重做按钮 */}
                <Link
                    iconOnly={true}
                    className={disabledRedo && "disabled"}
                    onClick={onRedoClick}
                >
                    {Device.ios ? (
                        <SvgIcon
                            symbolId={IconUndoForIos.id}
                            className="icon icon-svg"
                        />
                    ) : (
                        <SvgIcon
                            symbolId={IconUndoForAndroid.id}
                            className="icon icon-svg"
                        />
                    )}
                </Link>
            </React.Fragment>
        );
    },
    getEditOptions: function (props) {
        const {disabledEdit, disabledAdd, onEditClick, onAddClick} = props;
        return (
            <React.Fragment>
                <Link
                    iconOnly={true}
                    className={disabledEdit && "disabled"}
                    id="btn-edit"
                    icon="icon-edit-settings"
                    href={false}
                    onClick={onEditClick}
                >
                    {Device.ios ? (
                        <SvgIcon symbolId={IconEditSettingsForIos.id} className="icon icon-svg" />
                    ) : (
                        <SvgIcon symbolId={IconEditSettingsForAndroid.id} className="icon icon-svg" />
                    )}
                </Link>
                
                <Link
                    id="btn-add"
                    className={disabledAdd && "disabled"}
                    icon="icon-plus"
                    href={false}
                    onClick={onAddClick}
                >
                    {Device.ios ? (
                        <SvgIcon symbolId={IconPlusForIos.id} className="icon icon-svg" />
                    ) : (
                        <SvgIcon symbolId={IconPlusForAndroid.id} className="icon icon-svg" />
                    )}
                </Link>
            </React.Fragment>
        )
    }
};

EditorUIController.initThemeColors = function () {
    Common.EditorApi.get().asc_registerCallback("asc_onSendThemeColors", (function (e, t) {
        Common.Utils.ThemeColor.setColors(e, t)
    }))
};

EditorUIController.initCellInfo = function (e) {
    const t = Common.EditorApi.get(), n = e.storeFocusObjects, r = e.storeCellSettings, o = e.storeTextSettings,
        a = e.storeChartSettings;
    t.asc_registerCallback("asc_onSelectionChanged", (e => {
        n.resetCellInfo(e), n.setIsLocked(e), r.initCellSettings(e), o.initTextSettings(e);
        let s = t.asc_getGraphicObjectProps();
        s.length > 0 ? (n.resetFocusObjects(s), "obj" !== n.focusOn && n.changeFocus(!0), n.chartObject && a.updateChartStyles(t.asc_getChartPreviews(n.chartObject.get_ChartProperties().getType()))) : "cell" !== n.focusOn && n.changeFocus(!1)
    }));
    let s = n;
    s.intf = {}, s.intf.getSelections = () => {
        const e = [];
        let t, n, r, o, a, i, l, c, d, u = !1;
        switch (s._cellInfo.asc_getSelectionType()) {
            case Asc.c_oAscSelectionType.RangeCells:
                t = !0;
                break;
            case Asc.c_oAscSelectionType.RangeRow:
                n = !0;
                break;
            case Asc.c_oAscSelectionType.RangeCol:
                r = !0;
                break;
            case Asc.c_oAscSelectionType.RangeMax:
                o = !0;
                break;
            case Asc.c_oAscSelectionType.RangeImage:
                i = !0;
                break;
            case Asc.c_oAscSelectionType.RangeShape:
                c = !0;
                break;
            case Asc.c_oAscSelectionType.RangeChart:
                a = !0;
                break;
            case Asc.c_oAscSelectionType.RangeChartText:
                d = !0;
                break;
            case Asc.c_oAscSelectionType.RangeShapeText:
                l = !0
        }
        if (i || c || a) {
            i = c = a = !1;
            let e = !1, t = Common.EditorApi.get().asc_getGraphicObjectProps();
            for (let n = 0; n < t.length; n++) if (t[n].asc_getObjectType() == Asc.c_oAscTypeSelectElement.Image) {
                const r = t[n].asc_getObjectValue();
                u = u || r.asc_getLocked();
                const o = r.asc_getShapeProperties();
                o ? o.asc_getFromChart() ? a = !0 : c = !0 : r.asc_getChartProperties() ? (a = !0, e = !0) : i = !0
            }
        } else if (l || d) {
            const e = Common.EditorApi.get().asc_getGraphicObjectProps();
            let t = !1;
            for (var p = 0; p < e.length; p++) {
                const n = e[p].asc_getObjectType();
                if (n == Asc.c_oAscTypeSelectElement.Image) {
                    const t = e[p].asc_getObjectValue();
                    u = u || t.asc_getLocked()
                } else n == Asc.c_oAscTypeSelectElement.Paragraph || n == Asc.c_oAscTypeSelectElement.Math && (t = !0)
            }
        }
        return a || d ? (e.push("chart"), d && e.push("text")) : !c && !l || i ? i ? (e.push("image"), c && e.push("shape")) : (e.push("cell"), s._cellInfo.asc_getHyperlink() && e.push("hyperlink")) : (e.push("shape"), l && e.push("text")), e
    }, s.intf.getShapeObject = () => {
        const e = [];
        for (let t of s._focusObjects) t.get_ObjectType() === Asc.c_oAscTypeSelectElement.Image && t.get_ObjectValue() && t.get_ObjectValue().get_ShapeProperties() && e.push(t);
        if (e.length > 0) {
            return e[e.length - 1].get_ObjectValue()
        }
    }, s.intf.getImageObject = () => {
        const e = [];
        for (let t of s._focusObjects) t.get_ObjectType() === Asc.c_oAscTypeSelectElement.Image && e.push(t);
        if (e.length > 0) {
            return e[e.length - 1].get_ObjectValue()
        }
    }, s.intf.getChartObject = () => {
        const e = [];
        for (let t of s._focusObjects) t.get_ObjectType() === Asc.c_oAscTypeSelectElement.Image && t.get_ObjectValue() && t.get_ObjectValue().get_ChartProperties() && e.push(t);
        if (e.length > 0) {
            return e[e.length - 1].get_ObjectValue()
        }
    }
};

EditorUIController.initEditorStyles = function (e) {
    Common.EditorApi.get().asc_registerCallback("asc_onInitEditorStyles", (function (t) {
        e.initCellStyles(t)
    }))
};

EditorUIController.initFonts = function (e) {
    var t = Common.EditorApi.get(), n = e.storeCellSettings, r = e.storeTextSettings;
    t.asc_registerCallback("asc_onInitEditorFonts", (function (e, t) {
        n.initEditorFonts(e, t), r.initEditorFonts(e, t)
    })), t.asc_registerCallback("asc_onEditorSelectionChanged", (function (e) {
        n.initFontSettings(e), r.initFontSettings(e)
    }))
};

EditorUIController.ContextMenu = {
    mapMenuItems: e => {
        const {t: t} = e.props, n = t("ContextMenu", {returnObjects: !0}), {
                canViewComments: r,
                isDisconnected: o,
                wsProps: a,
                wsLock: s,
                isResolvedComments: i,
                isVersionHistoryMode: l
            } = e.props, c = Common.EditorApi.get(), d = c.asc_getCellInfo(), u = !!d.asc_getPivotTableInfo(),
            p = c.asc_canFillHandle(), m = [], h = [];
        let g, f, v, b, y, C, w, E, k, x = d.asc_getLocked();
        const S = d.asc_getSelectionType(), _ = d.asc_getXfs(), A = d.asc_getComments(),
            T = A[0] && A[0].asc_getSolved();
        switch (S) {
            case Asc.c_oAscSelectionType.RangeCells:
                g = !0;
                break;
            case Asc.c_oAscSelectionType.RangeRow:
                f = !0;
                break;
            case Asc.c_oAscSelectionType.RangeCol:
                v = !0;
                break;
            case Asc.c_oAscSelectionType.RangeMax:
                b = !0;
                break;
            case Asc.c_oAscSelectionType.RangeImage:
                C = !0;
                break;
            case Asc.c_oAscSelectionType.RangeShape:
                E = !0;
                break;
            case Asc.c_oAscSelectionType.RangeChart:
                y = !0;
                break;
            case Asc.c_oAscSelectionType.RangeChartText:
                k = !0;
                break;
            case Asc.c_oAscSelectionType.RangeShapeText:
                w = !0
        }
        return (C || E || y || w || k) && a.Objects ? [] : (!x && (C || E || y || w || k) && c.asc_getGraphicObjectProps().every((e => (e.asc_getObjectType() == Asc.c_oAscTypeSelectElement.Image && (x = e.asc_getObjectValue().asc_getLocked()), !x))), x || c.isCellEdited || o ? m.push({
            event: "copy",
            icon: IconCopy,id
        }) : l || (m.push({event: "cut", icon: IconCut.id}), m.push({
            event: "copy",
            icon: IconCopy.id
        }), m.push({event: "paste", icon: IconPaste.id}), C || E || y || w || k ? h.push({
            caption: n.menuEdit,
            event: "edit"
        }) : (v ? a.FormatColumns || (h.push({caption: n.menuHide, event: "hide"}), h.push({
            caption: n.menuShow,
            event: "show"
        })) : f ? a.FormatRows || (h.push({caption: n.menuHide, event: "hide"}), h.push({
            caption: n.menuShow,
            event: "show"
        })) : g && (x || h.push({
            caption: n.menuCell,
            event: "edit"
        }), d.asc_getMerge() != Asc.c_oAscMergeOptions.None || a.FormatCells || h.push({
            caption: n.menuMerge,
            event: "merge"
        }), d.asc_getMerge() != Asc.c_oAscMergeOptions.Merge || a.FormatCells || h.push({
            caption: n.menuUnmerge,
            event: "unmerge"
        }), a.FormatCells || h.push(_.asc_getWrapText() ? {
            caption: n.menuUnwrap,
            event: "unwrap"
        } : {
            caption: n.menuWrap,
            event: "wrap"
        })), h.push({
            caption: c.asc_getSheetViewSettings().asc_getIsFreezePane() ? n.menuUnfreezePanes : n.menuFreezePanes,
            event: "freezePanes"
        })), u || s || h.push({
            caption: n.menuDelete,
            event: "del"
        }), r && (A && A.length && (!T && !i || i) && h.push({
            caption: n.menuViewComment,
            event: "viewcomment"
        }), g && A && !A.length && !a.Objects && h.push({
            caption: n.menuAddComment,
            event: "addcomment"
        }))), d.asc_getHyperlink() && !d.asc_getMultiselect() ? (l || h.push({
            caption: t("ContextMenu.menuEditLink"),
            event: "editlink"
        }), h.push({
            caption: n.menuOpenLink,
            event: "openlink"
        })) : d.asc_getHyperlink() || d.asc_getMultiselect() || u || a.InsertHyperlinks || l || h.push({
            caption: n.menuAddLink,
            event: "addlink"
        }), w && c.asc_canAddShapeHyperlink() && (d.asc_getHyperlink() || a.InsertHyperlinks ? (l || h.push({
            caption: t("ContextMenu.menuEditLink"),
            event: "editlink"
        }), h.push({caption: n.menuOpenLink, event: "openlink"})) : l || h.push({
            caption: n.menuAddLink,
            event: "addlink"
        })), p && h.push({
            caption: t("ContextMenu.menuAutofill"),
            event: "autofillCells"
        }), Device.phone && h.length > 2 ? e.extraItems = h.splice(2, h.length, {
            caption: n.menuMore,
            event: "showActionSheet"
        }) : h.length > 4 && (e.extraItems = h.splice(3, h.length, {
            caption: n.menuMore,
            event: "showActionSheet"
        })), m.concat(h))
    }, handleMenuItemClick: (e, t) => {
        const n = Common.EditorApi.get();
        let r = n.asc_getCellInfo();
        switch (t) {
            case"cut":
                return n.asc_Cut();
            case"paste":
                return n.asc_Paste();
            case"addcomment":
                Common.Notifications.trigger("addcomment");
                break;
            case"del":
                if (n) switch (n.asc_getCellInfo().asc_getSelectionType()) {
                    case Asc.c_oAscSelectionType.RangeRow:
                        n.asc_deleteCells(Asc.c_oAscDeleteOptions.DeleteRows);
                        break;
                    case Asc.c_oAscSelectionType.RangeCol:
                        n.asc_deleteCells(Asc.c_oAscDeleteOptions.DeleteColumns);
                        break;
                    default:
                        n.asc_emptyCells(Asc.c_oAscCleanOptions.All)
                }
            case"wrap":
                n.asc_setCellTextWrap(!0);
                break;
            case"unwrap":
                n.asc_setCellTextWrap(!1);
                break;
            case"edit":
                setTimeout((() => {
                    e.props.openOptions("edit")
                }), 400);
                break;
            case"merge":
                e.onMergeCells();
                break;
            case"unmerge":
                n.asc_mergeCells(Asc.c_oAscMergeOptions.None);
                break;
            case"hide":
                n[r.asc_getSelectionType() == Asc.c_oAscSelectionType.RangeRow ? "asc_hideRows" : "asc_hideColumns"]();
                break;
            case"show":
                n[r.asc_getSelectionType() == Asc.c_oAscSelectionType.RangeRow ? "asc_showRows" : "asc_showColumns"]();
                break;
            case"addlink":
                setTimeout((() => {
                    e.props.openOptions("add-link")
                }), 400);
                break;
            case"editlink":
                setTimeout((() => {
                    e.props.openOptions("edit-link")
                }), 400);
                break;
            case"freezePanes":
                n.asc_freezePane();
                break;
            case"autofillCells":
                n.asc_fillHandleDone();
                break;
            default:
                return !1
        }
        return !0
    }
};

export default EditorUIController;
