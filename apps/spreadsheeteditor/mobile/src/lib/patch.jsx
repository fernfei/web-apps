import React from 'react';
import {Link, Icon, f7} from 'framework7-react';
import {
    AddCommentController,
    EditCommentController,
} from "../../../../common/mobile/lib/controller/collaboration/Comments";
import { Device } from "../../../../common/mobile/utils/device";

function ny(e, t) {
    var n = "undefined" != typeof Symbol && e[Symbol.iterator] || e["@@iterator"];
    if (!n) {
        if (Array.isArray(e) || (n = function (e, t) {
            if (!e) return;
            if ("string" == typeof e) return ry(e, t);
            var n = Object.prototype.toString.call(e).slice(8, -1);
            "Object" === n && e.constructor && (n = e.constructor.name);
            if ("Map" === n || "Set" === n) return Array.from(e);
            if ("Arguments" === n || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return ry(e, t)
        }(e)) || t && e && "number" == typeof e.length) {
            n && (e = n);
            var r = 0, a = function () {
            };
            return {
                s: a, n: function () {
                    return r >= e.length ? {done: !0} : {done: !1, value: e[r++]}
                }, e: function (e) {
                    throw e
                }, f: a
            }
        }
        throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")
    }
    var o, i = !0, s = !1;
    return {
        s: function () {
            n = n.call(e)
        }, n: function () {
            var e = n.next();
            return i = e.done, e
        }, e: function (e) {
            s = !0, o = e
        }, f: function () {
            try {
                i || null == n.return || n.return()
            } finally {
                if (s) throw o
            }
        }
    }
};

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
                <Link
                    className={disabledUndo ? "disabled" : ""}
                    icon="icon-undo"
                    onClick={onUndoClick}
                />
                <Link
                    className={disabledRedo ? "disabled" : ""}
                    icon="icon-redo"
                    onClick={onRedoClick}
                />
            </React.Fragment>
        );
    },
    getEditOptions: function (props) {
        return (
            <React.Fragment>
                <Link
                    className={(props.disabled || "obj" === props.focusOn && props.wsProps.Objects && props.isShapeLocked) && "disabled"}
                    id="btn-edit"
                    icon="icon-edit-settings"
                    href={false}
                    onClick={props.onEditClick}
                />
                <Link
                    className={(props.disabled || "obj" === props.focusOn && props.wsProps.Objects || "cell" === props.focusOn && props.wsProps.InsertHyperlinks && props.wsProps.Objects && props.wsProps.Sort) && "disabled"}
                    id="btn-add"
                    icon="icon-plus"
                    href={false}
                    onClick={props.onAddClick}
                />
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
    var t = Common.EditorApi.get(), n = e.storeFocusObjects, r = e.storeCellSettings, a = e.storeTextSettings,
        o = e.storeChartSettings;
    t.asc_registerCallback("asc_onSelectionChanged", (function (e) {
        n.resetCellInfo(e), n.setIsLocked(e), r.initCellSettings(e), a.initTextSettings(e);
        var i = t.asc_getGraphicObjectProps();
        i.length > 0 ? (n.resetFocusObjects(i), "obj" !== n.focusOn && n.changeFocus(!0), n.chartObject && o.updateChartStyles(t.asc_getChartPreviews(n.chartObject.get_ChartProperties().getType()))) : "cell" !== n.focusOn && n.changeFocus(!1)
    }));
    var i = n;
    i.intf = {}, i.intf.getSelections = function () {
        var e, t, n, r, a, o = [], s = !1;
        switch (i._cellInfo.asc_getSelectionType()) {
            case Asc.c_oAscSelectionType.RangeCells:
                !0;
                break;
            case Asc.c_oAscSelectionType.RangeRow:
                !0;
                break;
            case Asc.c_oAscSelectionType.RangeCol:
                !0;
                break;
            case Asc.c_oAscSelectionType.RangeMax:
                !0;
                break;
            case Asc.c_oAscSelectionType.RangeImage:
                t = !0;
                break;
            case Asc.c_oAscSelectionType.RangeShape:
                r = !0;
                break;
            case Asc.c_oAscSelectionType.RangeChart:
                e = !0;
                break;
            case Asc.c_oAscSelectionType.RangeChartText:
                a = !0;
                break;
            case Asc.c_oAscSelectionType.RangeShapeText:
                n = !0
        }
        if (t || r || e) {
            t = r = e = !1;
            for (var l = Common.EditorApi.get().asc_getGraphicObjectProps(), c = 0; c < l.length; c++) if (l[c].asc_getObjectType() == Asc.c_oAscTypeSelectElement.Image) {
                var u = l[c].asc_getObjectValue();
                s = s || u.asc_getLocked();
                var p = u.asc_getShapeProperties();
                p ? p.asc_getFromChart() ? e = !0 : r = !0 : u.asc_getChartProperties() ? (e = !0, !0) : t = !0
            }
        } else if (n || a) for (var d = Common.EditorApi.get().asc_getGraphicObjectProps(), f = 0; f < d.length; f++) {
            var h = d[f].asc_getObjectType();
            if (h == Asc.c_oAscTypeSelectElement.Image) {
                var m = d[f].asc_getObjectValue();
                s = s || m.asc_getLocked()
            } else h == Asc.c_oAscTypeSelectElement.Paragraph || h == Asc.c_oAscTypeSelectElement.Math && !0
        }
        return e || a ? (o.push("chart"), a && o.push("text")) : !r && !n || t ? t ? (o.push("image"), r && o.push("shape")) : (o.push("cell"), i._cellInfo.asc_getHyperlink() && o.push("hyperlink")) : (o.push("shape"), n && o.push("text")), o
    }, i.intf.getShapeObject = function () {
        var e, t = [], n = ny(i._focusObjects);
        try {
            for (n.s(); !(e = n.n()).done;) {
                var r = e.value;
                r.get_ObjectType() === Asc.c_oAscTypeSelectElement.Image && r.get_ObjectValue() && r.get_ObjectValue().get_ShapeProperties() && t.push(r)
            }
        } catch (e) {
            n.e(e)
        } finally {
            n.f()
        }
        return t.length > 0 ? t[t.length - 1].get_ObjectValue() : void 0
    }, i.intf.getImageObject = function () {
        var e, t = [], n = ny(i._focusObjects);
        try {
            for (n.s(); !(e = n.n()).done;) {
                var r = e.value;
                r.get_ObjectType() === Asc.c_oAscTypeSelectElement.Image && t.push(r)
            }
        } catch (e) {
            n.e(e)
        } finally {
            n.f()
        }
        return t.length > 0 ? t[t.length - 1].get_ObjectValue() : void 0
    }, i.intf.getChartObject = function () {
        var e, t = [], n = ny(i._focusObjects);
        try {
            for (n.s(); !(e = n.n()).done;) {
                var r = e.value;
                r.get_ObjectType() === Asc.c_oAscTypeSelectElement.Image && r.get_ObjectValue() && r.get_ObjectValue().get_ChartProperties() && t.push(r)
            }
        } catch (e) {
            n.e(e)
        } finally {
            n.f()
        }
        return t.length > 0 ? t[t.length - 1].get_ObjectValue() : void 0
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
    mapMenuItems: function (e) {
        var t, n, r, a, o, i, s, l, c = e.props.t, u = c("ContextMenu", {returnObjects: !0}), p = e.props,
            d = p.canViewComments, f = p.isDisconnected, h = p.wsProps, m = p.wsLock, g = p.isResolvedComments,
            v = p.isVersionHistoryMode, b = Common.EditorApi.get(), y = b.asc_getCellInfo(),
            C = !!y.asc_getPivotTableInfo(), k = [], w = [], E = y.asc_getLocked(),
            x = y.asc_getSelectionType(), S = y.asc_getXfs(), _ = y.asc_getComments(),
            A = _[0] && _[0].asc_getSolved();
        switch (x) {
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
                !0;
                break;
            case Asc.c_oAscSelectionType.RangeImage:
                o = !0;
                break;
            case Asc.c_oAscSelectionType.RangeShape:
                s = !0;
                break;
            case Asc.c_oAscSelectionType.RangeChart:
                a = !0;
                break;
            case Asc.c_oAscSelectionType.RangeChartText:
                l = !0;
                break;
            case Asc.c_oAscSelectionType.RangeShapeText:
                i = !0
        }
        return (o || s || a || i || l) && h.Objects ? [] : (!E && (o || s || a || i || l) && b.asc_getGraphicObjectProps().every((function (e) {
            return e.asc_getObjectType() == Asc.c_oAscTypeSelectElement.Image && (E = e.asc_getObjectValue().asc_getLocked()), !E
        })), E || b.isCellEdited || f ? k.push({event: "copy", icon: "icon-copy"}) : v || (k.push({
            event: "cut",
            icon: "icon-cut"
        }), k.push({event: "copy", icon: "icon-copy"}), k.push({
            event: "paste",
            icon: "icon-paste"
        }), o || s || a || i || l ? w.push({
            caption: u.menuEdit,
            event: "edit"
        }) : (r ? h.FormatColumns || (w.push({caption: u.menuHide, event: "hide"}), w.push({
            caption: u.menuShow,
            event: "show"
        })) : n ? h.FormatRows || (w.push({caption: u.menuHide, event: "hide"}), w.push({
            caption: u.menuShow,
            event: "show"
        })) : t && (E || w.push({
            caption: u.menuCell,
            event: "edit"
        }), y.asc_getMerge() != Asc.c_oAscMergeOptions.None || h.FormatCells || w.push({
            caption: u.menuMerge,
            event: "merge"
        }), y.asc_getMerge() != Asc.c_oAscMergeOptions.Merge || h.FormatCells || w.push({
            caption: u.menuUnmerge,
            event: "unmerge"
        }), h.FormatCells || w.push(S.asc_getWrapText() ? {
            caption: u.menuUnwrap,
            event: "unwrap"
        } : {
            caption: u.menuWrap,
            event: "wrap"
        })), w.push({
            caption: b.asc_getSheetViewSettings().asc_getIsFreezePane() ? u.menuUnfreezePanes : u.menuFreezePanes,
            event: "freezePanes"
        })), C || m || w.push({
            caption: u.menuDelete,
            event: "del"
        }), d && (_ && _.length && (!A && !g || g) ? w.push({
            caption: u.menuViewComment,
            event: "viewcomment"
        }) : t && _ && !_.length && !h.Objects && w.push({
            caption: u.menuAddComment,
            event: "addcomment"
        }))), y.asc_getHyperlink() && !y.asc_getMultiselect() ? (v || w.push({
            caption: c("ContextMenu.menuEditLink"),
            event: "editlink"
        }), w.push({
            caption: u.menuOpenLink,
            event: "openlink"
        })) : y.asc_getHyperlink() || y.asc_getMultiselect() || C || h.InsertHyperlinks || v || w.push({
            caption: u.menuAddLink,
            event: "addlink"
        }), i && b.asc_canAddShapeHyperlink() && (y.asc_getHyperlink() || h.InsertHyperlinks ? (v || w.push({
            caption: c("ContextMenu.menuEditLink"),
            event: "editlink"
        }), w.push({caption: u.menuOpenLink, event: "openlink"})) : v || w.push({
            caption: u.menuAddLink,
            event: "addlink"
        })), Device.phone && w.length > 2 ? e.extraItems = w.splice(2, w.length, {
            caption: u.menuMore,
            event: "showActionSheet"
        }) : w.length > 4 && (e.extraItems = w.splice(3, w.length, {
            caption: u.menuMore,
            event: "showActionSheet"
        })), k.concat(w))
    }, handleMenuItemClick: function (e, t) {
        var n = Common.EditorApi.get(), r = n.asc_getCellInfo();
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
                setTimeout((function () {
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
                setTimeout((function () {
                    e.props.openOptions("add-link")
                }), 400);
                break;
            case"editlink":
                setTimeout((function () {
                    e.props.openOptions("edit-link")
                }), 400);
                break;
            case"freezePanes":
                n.asc_freezePane();
                break;
            default:
                return !1
        }
        return !0
    }
};

export default EditorUIController;