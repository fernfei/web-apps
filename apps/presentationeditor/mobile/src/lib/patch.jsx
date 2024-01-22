import React from 'react';
import {Link, Icon, f7} from 'framework7-react';
import {
    AddCommentController,
    EditCommentController
} from "../../../../common/mobile/lib/controller/collaboration/Comments";
import { Device } from "../../../../common/mobile/utils/device";

function ab(e, t) {
    var n = "undefined" != typeof Symbol && e[Symbol.iterator] || e["@@iterator"];
    if (!n) {
        if (Array.isArray(e) || (n = function (e, t) {
            if (!e) return;
            if ("string" == typeof e) return ob(e, t);
            var n = Object.prototype.toString.call(e).slice(8, -1);
            "Object" === n && e.constructor && (n = e.constructor.name);
            if ("Map" === n || "Set" === n) return Array.from(e);
            if ("Arguments" === n || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return ob(e, t)
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
}
const EditorUIController = () => null;
EditorUIController.isSupportEditFeature = () => true;

EditorUIController.getUndoRedo = function (props) {
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
};

EditorUIController.getToolbarOptions = function (props) {
    const disableEditBtn = props.disabled;
    return (
        <React.Fragment>
            <Link
                className={disableEditBtn ? "disabled" : ""}
                id="btn-edit"
                icon="icon-edit-settings"
                href={false}
                onClick={e => props.onEditClick(e)}
            />
            <Link
                className={disableEditBtn ? "disabled" : ""}
                id="btn-add"
                icon="icon-plus"
                href={false}
                onClick={e => props.onAddClick(e)}
            />
        </React.Fragment>
    );
};

EditorUIController.initFocusObjects = function (e) {
    Common.EditorApi.get().asc_registerCallback("asc_onFocusObject", (function (t) {
        e.resetFocusObjects(t)
    })), e.intf = {}, e.intf.filterFocusObjects = function () {
        var t, n = [], r = !0, a = ab(e._focusObjects);
        try {
            for (a.s(); !(t = a.n()).done;) {
                var o = t.value, i = o.get_ObjectType(), s = o.get_ObjectValue();
                Asc.c_oAscTypeSelectElement.Paragraph == i ? s.get_Locked() || (r = !1) : Asc.c_oAscTypeSelectElement.Table == i ? s.get_Locked() || (n.push("table"), r = !1) : Asc.c_oAscTypeSelectElement.Slide == i ? s.get_LockLayout() || s.get_LockBackground() || s.get_LockTransition() || s.get_LockTiming() || n.push("slide") : Asc.c_oAscTypeSelectElement.Image == i ? s.get_Locked() || n.push("image") : Asc.c_oAscTypeSelectElement.Chart == i ? s.get_Locked() || n.push("chart") : Asc.c_oAscTypeSelectElement.Shape != i || s.get_FromChart() ? Asc.c_oAscTypeSelectElement.Hyperlink == i && n.push("hyperlink") : s.get_Locked() || (n.push("shape"), r = !1)
            }
        } catch (e) {
            a.e(e)
        } finally {
            a.f()
        }
        !r && n.indexOf("image") < 0 && n.unshift("text");
        var l = n.filter((function (e, t, n) {
            return n.indexOf(e) === t
        }));
        return l.indexOf("hyperlink") > -1 && l.indexOf("text") < 0 && l.splice(l.indexOf("hyperlink"), 1), l.indexOf("chart") > -1 && l.indexOf("shape") > -1 && l.splice(l.indexOf("shape"), 1), l
    }, e.intf.getSlideObject = function () {
        var t, n = [], r = ab(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                a.get_ObjectType() === Asc.c_oAscTypeSelectElement.Slide && n.push(a)
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }, e.intf.getParagraphObject = function () {
        var t, n = [], r = ab(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                a.get_ObjectType() === Asc.c_oAscTypeSelectElement.Paragraph && n.push(a)
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }, e.intf.getShapeObject = function () {
        var t, n = [], r = ab(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                a.get_ObjectType() === Asc.c_oAscTypeSelectElement.Shape && n.push(a)
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }, e.intf.getImageObject = function () {
        var t, n = [], r = ab(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                a.get_ObjectType() == Asc.c_oAscTypeSelectElement.Image && a.get_ObjectValue() && n.push(a)
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }, e.intf.getTableObject = function () {
        var t, n = [], r = ab(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                a.get_ObjectType() == Asc.c_oAscTypeSelectElement.Table && n.push(a)
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }, e.intf.getChartObject = function () {
        var t, n = [], r = ab(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                a.get_ObjectType() == Asc.c_oAscTypeSelectElement.Chart && n.push(a)
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }, e.intf.getLinkObject = function () {
        var t, n = [], r = ab(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                a.get_ObjectType() == Asc.c_oAscTypeSelectElement.Hyperlink && n.push(a)
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }
};

EditorUIController.initTableTemplates = function (e) {
    Common.EditorApi.get().asc_registerCallback("asc_onInitTableTemplates", (function (t) {
        return e.initTableTemplates(t)
    }))
};

EditorUIController.updateChartStyles = function (e, t) {
    var n = Common.EditorApi.get();
    n.asc_registerCallback("asc_onUpdateChartStyles", (function () {
        t.chartObject && e.updateChartStyles(n.asc_getChartPreviews(t.chartObject.getType()))
    }))
};

EditorUIController.getEditCommentControllers = function () {
    return (
        <React.Fragment>
            <AddCommentController />
            <EditCommentController />
        </React.Fragment>
    );
};

EditorUIController.ContextMenu = {
    mapMenuItems: function (e) {
        var t, n, r, a, o = e.props.t, i = o("ContextMenu", {returnObjects: !0}), s = e.props,
            l = s.canViewComments, c = s.isDisconnected, u = s.isVersionHistoryMode, p = Common.EditorApi.get(),
            d = p.getSelectedElements(), f = p.can_CopyCut(), h = [], m = [], v = !1, g = !1, b = !1, y = !1,
            w = !1, k = !1;
        if (d.forEach((function (e) {
            var t = e.get_ObjectType();
            e.get_ObjectValue();
            t == Asc.c_oAscTypeSelectElement.Paragraph ? v = !0 : t == Asc.c_oAscTypeSelectElement.Image ? b = !0 : t == Asc.c_oAscTypeSelectElement.Chart ? y = !0 : t == Asc.c_oAscTypeSelectElement.Shape ? w = !0 : t == Asc.c_oAscTypeSelectElement.Table ? g = !0 : t == Asc.c_oAscTypeSelectElement.Hyperlink ? k = !0 : t == Asc.c_oAscTypeSelectElement.Slide && !0
        })), t = v || b || y || w || g, f && t && h.push({event: "copy", icon: "icon-copy"}), d.length > 0) {
            var C = d[d.length - 1], E = (C.get_ObjectType(), C.get_ObjectValue()),
                x = "function" == typeof E.get_Locked && E.get_Locked();
            !x && (x = "function" == typeof E.get_LockDelete && E.get_LockDelete());
            if (!x && !c && !u) f && t && (h.push({
                event: "cut",
                icon: "icon-cut"
            }), r = 0, (n = h)[a = 1] = n.splice(r, 1, n[a])[0]), h.push({
                event: "paste",
                icon: "icon-paste"
            }), g && p.CheckBeforeMergeCells() && m.push({
                caption: i.menuMerge,
                event: "merge"
            }), g && p.CheckBeforeSplitCells() && m.push({
                caption: i.menuSplit,
                event: "split"
            }), t && m.push({caption: i.menuDelete, event: "delete"}), g && m.push({
                caption: i.menuDeleteTable,
                event: "deletetable"
            }), m.push({
                caption: i.menuEdit,
                event: "edit"
            }), k || !1 === p.can_AddHyperlink() || m.push({
                caption: i.menuAddLink,
                event: "addlink"
            }), e.isComments && l && m.push({
                caption: i.menuViewComment,
                event: "viewcomment"
            }), v && y || !1 === p.can_AddQuotedComment() || !l || m.push({
                caption: i.menuAddComment,
                event: "addcomment"
            }), k && m.push({caption: o("ContextMenu.menuEditLink"), event: "editlink"});
            k && m.push({caption: i.menuOpenLink, event: "openlink"})
        }
        return Device.phone && m.length > 2 ? e.extraItems = m.splice(2, m.length, {
            caption: i.menuMore,
            event: "showActionSheet"
        }) : m.length > 4 && (e.extraItems = m.splice(3, m.length, {
            caption: i.menuMore,
            event: "showActionSheet"
        })), h.concat(m)
    }, handleMenuItemClick: function (e, t) {
        var n = Common.EditorApi.get();
        switch (t) {
            case"cut":
                return n.Cut();
            case"paste":
                return n.Paste();
            case"addcomment":
                Common.Notifications.trigger("addcomment");
                break;
            case"merge":
                n.MergeCells();
                break;
            case"delete":
                n.asc_Remove();
                break;
            case"deletetable":
                n.remTable();
                break;
            case"split":
                e.showSplitModal();
                break;
            case"edit":
                setTimeout((function () {
                    e.props.openOptions("edit")
                }), 400);
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
            case"openlink":
                var r;
                n.getSelectedElements().forEach((function (e) {
                    e.get_ObjectType() == Asc.c_oAscTypeSelectElement.Hyperlink && (r = e.get_ObjectValue().get_Value())
                })), r && e.openLink(r);
                break;
            default:
                return !1
        }
        return !0
    }
};

EditorUIController.initThemeColors = function () {
    Common.EditorApi.get().asc_registerCallback("asc_onSendThemeColors", (function (e, t) {
        Common.Utils.ThemeColor.setColors(e, t)
    }))
};

EditorUIController.initFonts = function (e) {
    var t = Common.EditorApi.get();
    t.asc_registerCallback("asc_onInitEditorFonts", (function (t, n) {
        e.initEditorFonts(t, n)
    })), t.asc_registerCallback("asc_onFontFamily", (function (t) {
        e.resetFontName(t)
    })), t.asc_registerCallback("asc_onFontSize", (function (t) {
        e.resetFontSize(t)
    })), t.asc_registerCallback("asc_onBold", (function (t) {
        e.resetIsBold(t)
    })), t.asc_registerCallback("asc_onItalic", (function (t) {
        e.resetIsItalic(t)
    })), t.asc_registerCallback("asc_onUnderline", (function (t) {
        e.resetIsUnderline(t)
    })), t.asc_registerCallback("asc_onStrikeout", (function (t) {
        e.resetIsStrikeout(t)
    }))
};

EditorUIController.initEditorStyles = function (e) {
    var t = Common.EditorApi.get();
    t.asc_registerCallback("asc_onInitEditorStyles", (function (t) {
        var n = t[0] || [], r = t[1] || [], a = [];
        n.forEach((function (e, t) {
            a.push({themeId: e.get_Index(), offsety: 40 * t})
        })), r.forEach((function (e) {
            a.push({imageUrl: e.get_Image(), themeId: e.get_Index(), offsety: 0})
        })), e.addArrayThemes(a)
    })), t.asc_registerCallback("asc_onUpdateThemeIndex", (function (t) {
        e.changeSlideThemeIndex(t)
    })), t.asc_registerCallback("asc_onUpdateLayout", (function (t) {
        e.addArrayLayouts(t)
    }))
}
export default EditorUIController;