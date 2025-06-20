import React from 'react';
import {Link} from 'framework7-react';
import {
    AddCommentController,
    EditCommentController
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

const EditorUIController = () => null;
EditorUIController.isSupportEditFeature = () => true;

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

EditorUIController.initFocusObjects = function (e) {
    Common.EditorApi.get().asc_registerCallback("asc_onFocusObject", (t => {
            e.resetFocusObjects(t)
        }
    )),
        e.intf = {},
        e.intf.filterFocusObjects = () => {
            const t = [];
            let n = !0;
            for (let r of e._focusObjects) {
                const e = r.get_ObjectType()
                    , a = r.get_ObjectValue();
                Asc.c_oAscTypeSelectElement.Paragraph == e ? a.get_Locked() || (n = !1) : Asc.c_oAscTypeSelectElement.Table == e ? a.get_Locked() || (t.push("table"),
                    n = !1) : Asc.c_oAscTypeSelectElement.Slide == e ? a.get_LockLayout() || a.get_LockBackground() || a.get_LockTransition() || a.get_LockTiming() || t.push("slide") : Asc.c_oAscTypeSelectElement.Image == e ? a.get_Locked() || t.push("image") : Asc.c_oAscTypeSelectElement.Chart == e ? a.get_Locked() || t.push("chart") : Asc.c_oAscTypeSelectElement.Shape != e || a.get_FromChart() ? Asc.c_oAscTypeSelectElement.Hyperlink == e && t.push("hyperlink") : a.get_Locked() || (t.push("shape"),
                    n = !1)
            }
            !n && t.indexOf("image") < 0 && t.unshift("text");
            const r = t.filter(( (e, t, n) => n.indexOf(e) === t));
            return r.indexOf("hyperlink") > -1 && r.indexOf("text") < 0 && r.splice(r.indexOf("hyperlink"), 1),
            r.indexOf("chart") > -1 && r.indexOf("shape") > -1 && r.splice(r.indexOf("shape"), 1),
                r
        }
        ,
        e.intf.getSlideObject = () => {
            const t = [];
            for (let n of e._focusObjects)
                n.get_ObjectType() === Asc.c_oAscTypeSelectElement.Slide && t.push(n);
            if (t.length > 0) {
                return t[t.length - 1].get_ObjectValue()
            }
        }
        ,
        e.intf.getParagraphObject = () => {
            const t = [];
            for (let n of e._focusObjects)
                n.get_ObjectType() === Asc.c_oAscTypeSelectElement.Paragraph && t.push(n);
            if (t.length > 0) {
                return t[t.length - 1].get_ObjectValue()
            }
        }
        ,
        e.intf.getShapeObject = () => {
            const t = [];
            for (let n of e._focusObjects)
                n.get_ObjectType() === Asc.c_oAscTypeSelectElement.Shape && t.push(n);
            if (t.length > 0) {
                return t[t.length - 1].get_ObjectValue()
            }
        }
        ,
        e.intf.getImageObject = () => {
            const t = [];
            for (let n of e._focusObjects)
                n.get_ObjectType() == Asc.c_oAscTypeSelectElement.Image && n.get_ObjectValue() && t.push(n);
            if (t.length > 0) {
                return t[t.length - 1].get_ObjectValue()
            }
        }
        ,
        e.intf.getTableObject = () => {
            const t = [];
            for (let n of e._focusObjects)
                n.get_ObjectType() == Asc.c_oAscTypeSelectElement.Table && t.push(n);
            if (t.length > 0) {
                return t[t.length - 1].get_ObjectValue()
            }
        }
        ,
        e.intf.getChartObject = () => {
            const t = [];
            for (let n of e._focusObjects)
                n.get_ObjectType() == Asc.c_oAscTypeSelectElement.Chart && t.push(n);
            if (t.length > 0) {
                return t[t.length - 1].get_ObjectValue()
            }
        }
        ,
        e.intf.getLinkObject = () => {
            const t = [];
            for (let n of e._focusObjects)
                n.get_ObjectType() == Asc.c_oAscTypeSelectElement.Hyperlink && t.push(n);
            if (t.length > 0) {
                return t[t.length - 1].get_ObjectValue()
            }
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
    mapMenuItems: e => {
        const {t: t} = e.props
            , n = t("ContextMenu", {
            returnObjects: !0
        })
            , {canViewComments: r, isDisconnected: a, isVersionHistoryMode: o} = e.props
            , s = Common.EditorApi.get()
            , i = s.getSelectedElements()
            , l = s.can_CopyCut();
        let c = []
            , d = []
            , u = !1
            , p = !1
            , m = !1
            , h = !1
            , g = !1
            , f = !1
            , v = !1
            , b = !1;
        if (i.forEach((e => {
                const t = e.get_ObjectType();
                e.get_ObjectValue();
                t == Asc.c_oAscTypeSelectElement.Paragraph ? u = !0 : t == Asc.c_oAscTypeSelectElement.Image ? m = !0 : t == Asc.c_oAscTypeSelectElement.Chart ? h = !0 : t == Asc.c_oAscTypeSelectElement.Shape ? g = !0 : t == Asc.c_oAscTypeSelectElement.Table ? p = !0 : t == Asc.c_oAscTypeSelectElement.Hyperlink ? f = !0 : t == Asc.c_oAscTypeSelectElement.Slide && (v = !0)
            }
        )),
            b = u || m || h || g || p,
        l && b && c.push({
            event: "copy",
            icon: IconCopy.id
        }),
        i.length > 0) {
            let m = i[i.length - 1]
                , g = (m.get_ObjectType(),
                m.get_ObjectValue())
                , v = "function" == typeof g.get_Locked && g.get_Locked();
            !v && (v = "function" == typeof g.get_LockDelete && g.get_LockDelete());
            const y = function(e, t, n) {
                e[n] = e.splice(t, 1, e[n])[0]
            };
            if (!v && !a && !o) {
                l && b && (c.push({
                    event: "cut",
                    icon: IconCut.id
                }),
                    y(c, 0, 1)),
                    c.push({
                        event: "paste",
                        icon: IconPaste.id
                    }),
                p && s.CheckBeforeMergeCells() && d.push({
                    caption: n.menuMerge,
                    event: "merge"
                }),
                p && s.CheckBeforeSplitCells() && d.push({
                    caption: n.menuSplit,
                    event: "split"
                }),
                b && d.push({
                    caption: n.menuDelete,
                    event: "delete"
                }),
                p && d.push({
                    caption: n.menuDeleteTable,
                    event: "deletetable"
                }),
                    d.push({
                        caption: n.menuEdit,
                        event: "edit"
                    }),
                f || !1 === s.can_AddHyperlink() || d.push({
                    caption: n.menuAddLink,
                    event: "addlink"
                });
                u && h || !1 === s.can_AddQuotedComment() || !r || d.push({
                    caption: n.menuAddComment,
                    event: "addcomment"
                }),
                f && d.push({
                    caption: t("ContextMenu.menuEditLink"),
                    event: "editlink"
                })
            }
            e.isComments && r && d.push({
                caption: n.menuViewComment,
                event: "viewcomment"
            }),
            f && d.push({
                caption: n.menuOpenLink,
                event: "openlink"
            })
        }
        return Device.phone && d.length > 2 ? e.extraItems = d.splice(2, d.length, {
            caption: n.menuMore,
            event: "showActionSheet"
        }) : d.length > 4 && (e.extraItems = d.splice(3, d.length, {
            caption: n.menuMore,
            event: "showActionSheet"
        })),
            c.concat(d)
    }
    ,
    handleMenuItemClick: (e, t) => {
        const n = Common.EditorApi.get();
        switch (t) {
            case "cut":
                return n.Cut();
            case "paste":
                return n.Paste();
            case "addcomment":
                Common.Notifications.trigger("addcomment");
                break;
            case "merge":
                n.MergeCells();
                break;
            case "delete":
                n.asc_Remove();
                break;
            case "deletetable":
                n.remTable();
                break;
            case "split":
                e.showSplitModal();
                break;
            case "edit":
                setTimeout(( () => {
                        e.props.openOptions("edit")
                    }
                ), 400);
                break;
            case "addlink":
                setTimeout(( () => {
                        e.props.openOptions("add-link")
                    }
                ), 400);
                break;
            case "editlink":
                setTimeout(( () => {
                        e.props.openOptions("edit-link")
                    }
                ), 400);
                break;
            case "openlink":
                let t;
                n.getSelectedElements().forEach((e => {
                        e.get_ObjectType() == Asc.c_oAscTypeSelectElement.Hyperlink && (t = e.get_ObjectValue().get_Value())
                    }
                )),
                t && e.openLink(t);
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
    const t = Common.EditorApi.get();
    t.asc_registerCallback("asc_onInitEditorStyles", (t => {
            let n = t[0] || []
                , r = t[1] || []
                , a = [];
            n.forEach(( (e, t) => {
                    a.push({
                        themeId: e.get_Index(),
                        offsety: 40 * t
                    })
                }
            )),
                r.forEach((e => {
                        a.push({
                            imageUrl: e.get_Image(),
                            themeId: e.get_Index(),
                            offsety: 0
                        })
                    }
                )),
                e.addArrayThemes(a)
        }
    )),
        t.asc_registerCallback("asc_onUpdateThemeIndex", (t => {
                e.changeSlideThemeIndex(t)
            }
        )),
        t.asc_registerCallback("asc_onUpdateLayout", (t => {
                e.addArrayLayouts(t)
            }
        ))
}
export default EditorUIController;
