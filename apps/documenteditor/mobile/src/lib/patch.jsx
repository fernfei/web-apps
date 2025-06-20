import React from 'react';
import {Link, Icon, f7} from 'framework7-react';
import {
    AddCommentController,
    EditCommentController,
} from "../../../../common/mobile/lib/controller/collaboration/Comments";
import { Device } from "../../../../common/mobile/utils/device";
function zw(e, t) {
    var n = "undefined" != typeof Symbol && e[Symbol.iterator] || e["@@iterator"];
    if (!n) {
        if (Array.isArray(e) || (n = function (e, t) {
            if (!e) return;
            if ("string" == typeof e) return Hw(e, t);
            var n = Object.prototype.toString.call(e).slice(8, -1);
            "Object" === n && e.constructor && (n = e.constructor.name);
            if ("Map" === n || "Set" === n) return Array.from(e);
            if ("Arguments" === n || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return Hw(e, t)
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
const EditorUIController = () => {
    return null
};

EditorUIController.isSupportEditFeature = () => {
    return true
};

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
    t.asc_setParagraphStylesSizes(330, 38), t.asc_registerCallback("asc_onInitEditorStyles", (function (t) {
        e.initEditorStyles(t)
    })), t.asc_registerCallback("asc_onParaStyleName", (function (t) {
        e.changeParaStyleName(t)
    }))
};
EditorUIController.initFocusObjects = function (e) {
    Common.EditorApi.get().asc_registerCallback("asc_onFocusObject", (function (t) {
        e.resetFocusObjects(t)
    })), e.intf = {}, e.intf.filterFocusObjects = function () {
        var t, n = [], r = zw(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value, o = a.get_ObjectType();
                if (Asc.c_oAscTypeSelectElement.Paragraph === o) n.push("text", "paragraph"); else if (Asc.c_oAscTypeSelectElement.Table === o) n.push("table"); else if (Asc.c_oAscTypeSelectElement.Image === o) if (a.get_ObjectValue().get_ChartProperties()) {
                    var i = n.indexOf("shape");
                    i < 0 ? n.push("chart") : n.splice(i, 1, "chart")
                } else a.get_ObjectValue().get_ShapeProperties() && !n.includes("chart") ? n.push("shape") : n.push("image"); else Asc.c_oAscTypeSelectElement.Hyperlink === o ? n.push("hyperlink") : Asc.c_oAscTypeSelectElement.Header === o && n.push("header")
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.filter((function (e, t, n) {
            return n.indexOf(e) === t
        }))
    }, e.intf.getHeaderObject = function () {
        var t, n = [], r = zw(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                a.get_ObjectType() == Asc.c_oAscTypeSelectElement.Header && n.push(a)
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }, e.intf.getParagraphObject = function () {
        var t, n = [], r = zw(e._focusObjects);
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
        var t, n = [], r = zw(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                a.get_ObjectType() === Asc.c_oAscTypeSelectElement.Image && a.get_ObjectValue() && a.get_ObjectValue().get_ShapeProperties() && n.push(a)
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }, e.intf.getImageObject = function () {
        var t, n = [], r = zw(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                if (a.get_ObjectType() == Asc.c_oAscTypeSelectElement.Image) {
                    var o = a.get_ObjectValue();
                    o && null === o.get_ShapeProperties() && null === o.get_ChartProperties() && n.push(a)
                }
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }, e.intf.getTableObject = function () {
        var t, n = [], r = zw(e._focusObjects);
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
        var t, n = [], r = zw(e._focusObjects);
        try {
            for (r.s(); !(t = r.n()).done;) {
                var a = t.value;
                a.get_ObjectType() == Asc.c_oAscTypeSelectElement.Image && a.get_ObjectValue() && a.get_ObjectValue().get_ChartProperties() && n.push(a)
            }
        } catch (e) {
            r.e(e)
        } finally {
            r.f()
        }
        return n.length > 0 ? n[n.length - 1].get_ObjectValue() : void 0
    }, e.intf.getLinkObject = function () {
        var t, n = [], r = zw(e._focusObjects);
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
        e.chartObject && e.chartObject.get_ChartProperties() && e.updateChartStyles(n.asc_getChartPreviews(t.chartObject.get_ChartProperties().getType()))
    }))
};
EditorUIController.ContextMenu = {
    mapMenuItems: function (e) {
        var t = e.props.t, n = t("ContextMenu", {returnObjects: !0}), r = e.props, a = r.isEdit,
            o = r.canViewComments, i = r.canEditComments, s = r.canReview, l = r.isDisconnected,
            c = r.displayMode, u = r.dataDoc, p = r.isViewer, d = r.isProtected, f = r.typeProtection,
            m = Common.EditorApi.get(), h = m.asc_GetTableOfContentsPr(!0), v = m.getSelectedElements(),
            g = m.can_CopyCut(), b = "markup" !== c,
            y = m.asc_IsContentControl() ? m.asc_GetContentControlProperties() : null,
            C = y ? y.get_Lock() : Asc.c_oAscSdtLockType.Unlocked,
            w = C == Asc.c_oAscSdtLockType.SdtContentLocked || C == Asc.c_oAscSdtLockType.SdtLocked,
            k = f === Asc.c_oAscEDocProtect.Forms, E = [], x = [];
        !g || d && k || E.push({event: "copy", icon: "icon-copy"});
        var S, _, T, A = !1, O = !1, P = !1, R = !1, M = !1, I = !1, N = !1, L = !1, D = !1, B = !1;
        if (v.forEach((function (e) {
            var t = e.get_ObjectType(), n = e.get_ObjectValue();
            t == Asc.c_oAscTypeSelectElement.Header ? B = n.get_Locked() : t == Asc.c_oAscTypeSelectElement.Paragraph ? (N = n.get_Locked(), A = !0) : t == Asc.c_oAscTypeSelectElement.Image ? (D = n.get_Locked(), n && n.get_ChartProperties() ? R = !0 : n && n.get_ShapeProperties() ? M = !0 : P = !0) : t == Asc.c_oAscTypeSelectElement.Table ? (L = n.get_Locked(), O = !0) : t == Asc.c_oAscTypeSelectElement.Hyperlink && (I = !0)
        })), v.length > 0) {
            if (a && !l) {
                N || L || D || B || !g || b || p && "oform" !== u.fileType || (E.push({
                    event: "cut",
                    icon: "icon-cut"
                }), _ = 0, (S = E)[T = 1] = S.splice(_, 1, S[T])[0]), N || L || D || B || "oform" === u.fileType || b || p && "oform" !== u.fileType || E.push({
                    event: "paste",
                    icon: "icon-paste"
                }), !O || !m.CheckBeforeMergeCells() || L || B || b || p || x.push({
                    caption: n.menuMerge,
                    event: "merge"
                }), !O || !m.CheckBeforeSplitCells() || L || B || b || p || x.push({
                    caption: n.menuSplit,
                    event: "split"
                }), N || L || D || B || b || w || p || x.push({
                    caption: n.menuDelete,
                    event: "delete"
                }), !O || L || N || B || b || p || x.push({
                    caption: n.menuDeleteTable,
                    event: "deletetable"
                }), N || L || D || B || b || p || x.push({
                    caption: n.menuEdit,
                    event: "edit"
                }), !m.can_AddHyperlink() || B || b || p || x.push({
                    caption: n.menuAddLink,
                    event: "addlink"
                }), !s || b || p || (e.inRevisionChange ? x.push({
                    caption: n.menuReviewChange,
                    event: "reviewchange"
                }) : x.push({
                    caption: n.menuReview,
                    event: "review"
                })), e.isComments && o && !b && x.push({caption: n.menuViewComment, event: "viewcomment"});
                var $ = M || R || P || O;
                !o || !i || !1 === m.can_AddQuotedComment() || N || L || D || B || !A && $ || b || p && !i || x.push({
                    caption: n.menuAddComment,
                    event: "addcomment"
                });
                var F = m.asc_GetCurrentNumberingId();
                if (null !== F && !p) {
                    var j = m.asc_GetNumberingPr(F).get_Lvl(m.asc_GetCurrentNumberingLvl()), V = j.get_Format(),
                        z = m.asc_GetCalculatedNumberingValue();
                    j.get_Start() != z && x.push({
                        caption: V == Asc.c_oAscNumberingFormat.Bullet ? n.menuSeparateList : n.menuStartNewList,
                        event: "startNumbering"
                    }), V != Asc.c_oAscNumberingFormat.Bullet && x.push({
                        caption: n.menuStartNumberingFrom,
                        event: "startNumberingFrom"
                    }), x.push({
                        caption: V == Asc.c_oAscNumberingFormat.Bullet ? n.menuJoinList : n.menuContinueNumbering,
                        event: "continueNumbering"
                    })
                }
                I && (x.push({
                    caption: n.menuOpenLink,
                    event: "openlink"
                }), p || x.push({caption: t("ContextMenu.menuEditLink"), event: "editlink"}))
            }
        }
        return h && a && !p && (x.push({
            caption: t("ContextMenu.textRefreshEntireTable"),
            event: "refreshEntireTable"
        }), x.push({
            caption: t("ContextMenu.textRefreshPageNumbersOnly"),
            event: "refreshPageNumbers"
        })), Device.phone && x.length > 2 ? e.extraItems = x.splice(2, x.length, {
            caption: n.menuMore,
            event: "showActionSheet"
        }) : x.length > 4 && (e.extraItems = x.splice(3, x.length, {
            caption: n.menuMore,
            event: "showActionSheet"
        })), E.concat(x)
    }, handleMenuItemClick: function (e, t) {
        var n, r, a = Common.EditorApi.get(), o = (0, e.props.t)("ContextMenu", {returnObjects: !0}),
            i = function (e, t) {
                var n = a.asc_GetTableOfContentsPr(t);
                n && (t && n && (t = n.get_InternalClass()), a.asc_UpdateTableOfContents("pages" == e, t))
            }, s = a.asc_GetCurrentNumberingId();
        switch (s && (n = a.asc_GetNumberingPr(s).get_Lvl(a.asc_GetCurrentNumberingLvl()), r = n.get_Format()), t) {
            case"cut":
                return a.Cut();
            case"paste":
                return a.Paste();
            case"addcomment":
                Common.Notifications.trigger("addcomment");
                break;
            case"merge":
                a.MergeCells();
                break;
            case"delete":
                a.asc_Remove();
                break;
            case"deletetable":
                a.remTable();
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
            case"startNumbering":
                a.asc_RestartNumbering(a.asc_GetNumberingPr(s).get_Lvl(a.asc_GetCurrentNumberingLvl()).get_Start());
                break;
            case"startNumberingFrom":
                var l, c = function (e) {
                    e = parseInt(e);
                    for (var t = Math.ceil(e / 26), n = String.fromCharCode((e - 1) % 26 + "A".charCodeAt(0)), r = "", a = 0; a < t; a++) r += n;
                    return r
                }, u = function (e) {
                    e = parseInt(e);
                    for (var t = "", n = [["M", 1e3], ["CM", 900], ["D", 500], ["CD", 400], ["C", 100], ["XC", 90], ["L", 50], ["XL", 40], ["X", 10], ["IX", 9], ["V", 5], ["IV", 4], ["I", 1]], r = n[0][1], a = Math.floor(e / r), o = 0, i = 0; i < a; i++) t += n[o][0];
                    for (e -= a * r, o++; e > 0;) (a = e - (r = n[o][1])) >= 0 ? (t += n[o][0], e = a) : o++;
                    return t
                }, p = function (e) {
                    switch (r) {
                        case Asc.c_oAscNumberingFormat.UpperRoman:
                            l.$inputEl[0].value = u(e);
                            break;
                        case Asc.c_oAscNumberingFormat.LowerRoman:
                            l.$inputEl[0].value = u(e).toLocaleLowerCase();
                            break;
                        case Asc.c_oAscNumberingFormat.UpperLetter:
                            l.$inputEl[0].value = c(e);
                            break;
                        case Asc.c_oAscNumberingFormat.LowerLetter:
                            l.$inputEl[0].value = c(e).toLocaleLowerCase();
                            break;
                        default:
                            l.$inputEl[0].value = e
                    }
                };
                f7.dialog.create({
                    title: o.textNumberingValue,
                    content: '<div class="content-block stepper-block">\n                        <div class="stepper stepper-large">\n                            <div class="stepper-button-minus">\n                            '.concat(Device.android ? '<i class="icon icon-expand-down"></i>' : "-", '\n                            </div>\n                            <div class="stepper-input-wrap">\n                                <input type="text" readonly />\n                            </div>\n                            <div class="stepper-button-plus">\n                                ').concat(Device.android ? '<i class="icon icon-expand-up"></i>' : "+", "\n                            </div>\n                        </div>\n                    </div>"),
                    buttons: [{
                        text: o.textOk, bold: !0, onClick: function () {
                            a.asc_RestartNumbering(l.value)
                        }
                    }, {text: o.textCancel}],
                    on: {
                        open: function () {
                            (l = f7.stepper.create({
                                el: ".stepper",
                                value: a.asc_GetCalculatedNumberingValue()
                            })).on("change", (function () {
                                return p(l.value)
                            })), p(l.value)
                        }
                    }
                }).open();
                break;
            case"continueNumbering":
                a.asc_ContinueNumbering();
                break;
            case"refreshEntireTable":
                i("all");
                break;
            case"refreshPageNumbers":
                i("pages");
                break;
            default:
                return !1
        }
        return !0
    }
};
EditorUIController.getEditCommentControllers = function () {
    return (
        <React.Fragment>
            <AddCommentController />
            <EditCommentController />
        </React.Fragment>
    );
};

export default EditorUIController;
