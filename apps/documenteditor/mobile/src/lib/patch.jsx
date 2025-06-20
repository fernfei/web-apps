import React from 'react';
import {Link} from 'framework7-react';
import {
    AddCommentController,
    EditCommentController,
} from "../../../../common/mobile/lib/controller/collaboration/Comments";
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

EditorUIController.getUndoRedo = function (props) {
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
};

EditorUIController.getToolbarOptions = function (props) {
    const {disabledEdit, disabledAdd, onEditClick, onAddClick} = props;
    return (
        <React.Fragment>
            <Link
                iconOnly={true}
                className={disabledEdit && "disabled"}
                id="btn-edit"
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
                iconOnly={true}
                className={disabledAdd && "disabled"}
                id="btn-add"
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
    Common.EditorApi.get().asc_registerCallback("asc_onFocusObject", (t => {
        e.resetFocusObjects(t)
    })), e.intf = {}, e.intf.filterFocusObjects = () => {
        const t = [];
        for (let n of e._focusObjects) {
            let e = n.get_ObjectType();
            if (Asc.c_oAscTypeSelectElement.Paragraph === e) t.push("text", "paragraph"); else if (Asc.c_oAscTypeSelectElement.Table === e) t.push("table"); else if (Asc.c_oAscTypeSelectElement.Image === e) if (n.get_ObjectValue().get_ChartProperties()) {
                let e = t.indexOf("shape");
                e < 0 ? t.push("chart") : t.splice(e, 1, "chart")
            } else n.get_ObjectValue().get_ShapeProperties() && !t.includes("chart") ? t.push("shape") : t.push("image"); else Asc.c_oAscTypeSelectElement.Hyperlink === e ? t.push("hyperlink") : Asc.c_oAscTypeSelectElement.Header === e && t.push("header")
        }
        return t.filter(((e, t, n) => n.indexOf(e) === t))
    }, e.intf.getHeaderObject = () => {
        const t = [];
        for (let n of e._focusObjects) n.get_ObjectType() == Asc.c_oAscTypeSelectElement.Header && t.push(n);
        if (t.length > 0) {
            return t[t.length - 1].get_ObjectValue()
        }
    }, e.intf.getParagraphObject = () => {
        const t = [];
        for (let n of e._focusObjects) n.get_ObjectType() === Asc.c_oAscTypeSelectElement.Paragraph && t.push(n);
        if (t.length > 0) {
            return t[t.length - 1].get_ObjectValue()
        }
    }, e.intf.getShapeObject = () => {
        const t = [];
        for (let n of e._focusObjects) n.get_ObjectType() === Asc.c_oAscTypeSelectElement.Image && n.get_ObjectValue() && n.get_ObjectValue().get_ShapeProperties() && t.push(n);
        if (t.length > 0) {
            return t[t.length - 1].get_ObjectValue()
        }
    }, e.intf.getImageObject = () => {
        const t = [];
        for (let n of e._focusObjects) if (n.get_ObjectType() == Asc.c_oAscTypeSelectElement.Image) {
            const e = n.get_ObjectValue();
            e && null === e.get_ShapeProperties() && null === e.get_ChartProperties() && t.push(n)
        }
        if (t.length > 0) {
            return t[t.length - 1].get_ObjectValue()
        }
    }, e.intf.getTableObject = () => {
        const t = [];
        for (let n of e._focusObjects) n.get_ObjectType() == Asc.c_oAscTypeSelectElement.Table && t.push(n);
        if (t.length > 0) {
            return t[t.length - 1].get_ObjectValue()
        }
    }, e.intf.getChartObject = () => {
        const t = [];
        for (let n of e._focusObjects) n.get_ObjectType() == Asc.c_oAscTypeSelectElement.Image && n.get_ObjectValue() && n.get_ObjectValue().get_ChartProperties() && t.push(n);
        if (t.length > 0) {
            return t[t.length - 1].get_ObjectValue()
        }
    }, e.intf.getLinkObject = () => {
        const t = [];
        for (let n of e._focusObjects) n.get_ObjectType() == Asc.c_oAscTypeSelectElement.Hyperlink && t.push(n);
        if (t.length > 0) {
            return t[t.length - 1].get_ObjectValue()
        }
    }
};
EditorUIController.initTableTemplates = function (e) {
    Common.EditorApi.get().asc_registerCallback("asc_onInitTableTemplates", (t => e.initTableTemplates(t)))
};
EditorUIController.updateChartStyles = function (e, t) {
    const n = Common.EditorApi.get();
    n.asc_registerCallback("asc_onUpdateChartStyles", (() => {
        e.chartObject && e.chartObject.get_ChartProperties() && e.updateChartStyles(n.asc_getChartPreviews(t.chartObject.get_ChartProperties().getType()))
    }))
};
EditorUIController.ContextMenu = {
    mapMenuItems: e => {
        const {t: t} = e.props, n = t("ContextMenu", {returnObjects: !0}), {
                isEdit: o,
                canViewComments: r,
                canEditComments: a,
                canReview: s,
                isDisconnected: i,
                displayMode: l,
                isViewer: c,
                isProtected: d,
                typeProtection: p,
                isForm: u
            } = e.props, m = Common.EditorApi.get(), h = m.asc_GetTableOfContentsPr(!0),
            g = m.getSelectedElements(), f = m.can_CopyCut(), v = "markup" !== l,
            b = m.asc_IsContentControl() ? m.asc_GetContentControlProperties() : null,
            y = b ? b.get_Lock() : Asc.c_oAscSdtLockType.Unlocked,
            C = y == Asc.c_oAscSdtLockType.SdtContentLocked || y == Asc.c_oAscSdtLockType.SdtLocked,
            w = p === Asc.c_oAscEDocProtect.Forms;
        let E = [], k = [];
        !f || d && w || E.push({event: "copy", icon: IconCopy.id});
        let x = !1, S = !1, _ = !1, T = !1, A = !1, M = !1, P = !1, O = !1, L = !1, I = !1;
        if (g.forEach((e => {
            const t = e.get_ObjectType(), n = e.get_ObjectValue();
            t == Asc.c_oAscTypeSelectElement.Header ? I = n.get_Locked() : t == Asc.c_oAscTypeSelectElement.Paragraph ? (P = n.get_Locked(), x = !0) : t == Asc.c_oAscTypeSelectElement.Image ? (L = n.get_Locked(), n && n.get_ChartProperties() ? T = !0 : n && n.get_ShapeProperties() ? A = !0 : _ = !0) : t == Asc.c_oAscTypeSelectElement.Table ? (O = n.get_Locked(), S = !0) : t == Asc.c_oAscTypeSelectElement.Hyperlink && (M = !0)
        })), g.length > 0) {
            const l = function (e, t, n) {
                e[n] = e.splice(t, 1, e[n])[0]
            };
            if (o && !i) {
                P || O || L || I || !f || v || c && !u || (E.push({
                    event: "cut", icon: IconCut.id
                }), l(E, 0, 1)), P || O || L || I || u || v || c && !u || E.push({
                    event: "paste", icon: IconPaste.id
                }), !S || !m.CheckBeforeMergeCells() || O || I || v || c || k.push({
                    caption: n.menuMerge, event: "merge"
                }), !S || !m.CheckBeforeSplitCells() || O || I || v || c || k.push({
                    caption: n.menuSplit, event: "split"
                }), P || O || L || I || v || C || c || k.push({
                    caption: n.menuDelete, event: "delete"
                }), !S || O || P || I || v || c || k.push({
                    caption: n.menuDeleteTable, event: "deletetable"
                }), P || O || L || I || v || c || k.push({
                    caption: n.menuEdit, event: "edit"
                }), !m.can_AddHyperlink() || I || v || c || k.push({
                    caption: n.menuAddLink, event: "addlink"
                }), !s || v || c || (e.inRevisionChange ? k.push({
                    caption: n.menuReviewChange, event: "reviewchange"
                }) : k.push({
                    caption: n.menuReview, event: "review"
                })), e.isComments && r && !v && k.push({caption: n.menuViewComment, event: "viewcomment"});
                const o = A || T || _ || S;
                !r || !a || !1 === m.can_AddQuotedComment() || P || O || L || I || !x && o || v || c && !a || k.push({
                    caption: n.menuAddComment, event: "addcomment"
                });
                let i = m.asc_GetCurrentNumberingId();
                if (null !== i && !c) {
                    let e = m.asc_GetNumberingPr(i).get_Lvl(m.asc_GetCurrentNumberingLvl()), t = e.get_Format(),
                        o = m.asc_GetCalculatedNumberingValue();
                    e.get_Start() != o && k.push({
                        caption: t == Asc.c_oAscNumberingFormat.Bullet ? n.menuSeparateList : n.menuStartNewList,
                        event: "startNumbering"
                    }), t != Asc.c_oAscNumberingFormat.Bullet && k.push({
                        caption: n.menuStartNumberingFrom, event: "startNumberingFrom"
                    }), k.push({
                        caption: t == Asc.c_oAscNumberingFormat.Bullet ? n.menuJoinList : n.menuContinueNumbering,
                        event: "continueNumbering"
                    })
                }
                M && (k.push({
                    caption: n.menuOpenLink, event: "openlink"
                }), c || k.push({caption: t("ContextMenu.menuEditLink"), event: "editlink"}))
            }
        }
        return h && o && !c && (k.push({
            caption: t("ContextMenu.textRefreshEntireTable"), event: "refreshEntireTable"
        }), k.push({
            caption: t("ContextMenu.textRefreshPageNumbersOnly"), event: "refreshPageNumbers"
        })), Device.phone && k.length > 2 ? e.extraItems = k.splice(2, k.length, {
            caption: n.menuMore, event: "showActionSheet"
        }) : k.length > 4 && (e.extraItems = k.splice(3, k.length, {
            caption: n.menuMore, event: "showActionSheet"
        })), E.concat(k)
    }, handleMenuItemClick: (e, t) => {
        const n = Common.EditorApi.get(), {t: o} = e.props, r = o("ContextMenu", {returnObjects: !0}),
            a = (e, t) => {
                let o = n.asc_GetTableOfContentsPr(t);
                o && (t && o && (t = o.get_InternalClass()), n.asc_UpdateTableOfContents("pages" == e, t))
            };
        let s, i, l = n.asc_GetCurrentNumberingId();
        switch (l && (s = n.asc_GetNumberingPr(l).get_Lvl(n.asc_GetCurrentNumberingLvl()), i = s.get_Format()), t) {
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
                e.props.openOptions("edit");
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
            case"startNumbering":
                n.asc_RestartNumbering(n.asc_GetNumberingPr(l).get_Lvl(n.asc_GetCurrentNumberingLvl()).get_Start());
                break;
            case"startNumberingFrom":
                let t;
                const o = e => {
                    e = parseInt(e);
                    let t = Math.ceil(e / 26), n = String.fromCharCode((e - 1) % 26 + "A".charCodeAt(0)),
                        o = "";
                    for (let e = 0; e < t; e++) o += n;
                    return o
                }, s = e => {
                    e = parseInt(e);
                    let t = "",
                        n = [["M", 1e3], ["CM", 900], ["D", 500], ["CD", 400], ["C", 100], ["XC", 90], ["L", 50], ["XL", 40], ["X", 10], ["IX", 9], ["V", 5], ["IV", 4], ["I", 1]],
                        o = n[0][1], r = Math.floor(e / o), a = 0;
                    for (let e = 0; e < r; e++) t += n[a][0];
                    for (e -= r * o, a++; e > 0;) o = n[a][1], r = e - o, r >= 0 ? (t += n[a][0], e = r) : a++;
                    return t
                }, c = e => {
                    switch (i) {
                        case Asc.c_oAscNumberingFormat.UpperRoman:
                            t.$inputEl[0].value = s(e);
                            break;
                        case Asc.c_oAscNumberingFormat.LowerRoman:
                            t.$inputEl[0].value = s(e).toLocaleLowerCase();
                            break;
                        case Asc.c_oAscNumberingFormat.UpperLetter:
                            t.$inputEl[0].value = o(e);
                            break;
                        case Asc.c_oAscNumberingFormat.LowerLetter:
                            t.$inputEl[0].value = o(e).toLocaleLowerCase();
                            break;
                        default:
                            t.$inputEl[0].value = e
                    }
                };
                yi.dialog.create({
                    title: r.textNumberingValue,
                    content: '<div class="content-block stepper-block">\n                        <div class="stepper stepper-large">\n                            <div class="stepper-button-minus">\n                            '.concat(Device.android ? '<i class="icon icon-expand-down"></i>' : "-", '\n                            </div>\n                            <div class="stepper-input-wrap">\n                                <input type="text" readonly />\n                            </div>\n                            <div class="stepper-button-plus">\n                                ').concat(Device.android ? '<i class="icon icon-expand-up"></i>' : "+", "\n                            </div>\n                        </div>\n                    </div>"),
                    buttons: [{
                        text: r.textOk, bold: !0, onClick: () => {
                            n.asc_RestartNumbering(t.value)
                        }
                    }, {text: r.textCancel}],
                    on: {
                        open: () => {
                            t = yi.stepper.create({
                                el: ".stepper", value: n.asc_GetCalculatedNumberingValue()
                            }), t.on("change", (() => c(t.value))), c(t.value)
                        }
                    }
                }).open();
                break;
            case"continueNumbering":
                n.asc_ContinueNumbering();
                break;
            case"refreshEntireTable":
                a("all");
                break;
            case"refreshPageNumbers":
                a("pages");
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
