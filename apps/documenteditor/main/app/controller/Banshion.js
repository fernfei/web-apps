/**
 * create by hufei on 06/17/2023
 * Banshion tech support for document editor
 */

define([
    'core',
], function (core) {
    'use strict';
    DE.Controllers.Banshion = Backbone.Controller.extend(_.assign({
        // Specifying a Banshion model
        models: [],

        // Specifying a collection of out Banshion
        collections: [],

        // Specifying application views
        views: [],
        initialize: function () {
        },
        /**
         * When our application is ready, lets get started. {@link Backbone.Application#launchControllers callback}
         */
        onLaunch: function () {
            this.api = this.getApplication().getController('Viewport').getApi();
            Common.Gateway.on('bindListenerCursorEvent', _.bind(this.bindListenerCursorEvent, this));
            Common.Gateway.on('unbindListenerCursorEvent', _.bind(this.unbindListenerCursorEvent, this));
            Common.Gateway.on('getBookmarkPosition', _.bind(this.getBookmarkPosition, this));
            Common.Gateway.on('switchDisableEditor', _.bind(this.switchDisableEditor, this));
            Common.Gateway.on('gotobookmark', _.bind(this.goToBookmark, this));
            Common.Gateway.on('addBookMark', _.bind(this.addBookMark, this));
            Common.Gateway.on('switchCursor', _.bind(this.switchCursor, this));
            Common.Gateway.on('setWatermark', _.bind(this.setWatermark, this));
            Common.Gateway.on('removeWatermark', _.bind(this.removeWatermark, this));
            Common.Gateway.on('removeBookmark', _.bind(this.removeBookmark, this));
            Common.Gateway.on('getCatalogList', _.bind(this.getCatalogList, this));
            Common.Gateway.on('signature', _.bind(this.signature, this));
            this._state = {
                isHighlightedResults: false,
            };
        },
        calculateImageWH(aImages) {
            function getImageDimensions(url) {
                return new Promise((resolve, reject) => {
                    const img = new Image();

                    img.onload = function () {
                        const width = img.width;
                        const height = img.height;
                        resolve({width, height});
                    };

                    img.onerror = function () {
                        reject(new Error('Failed to load image'));
                    };

                    img.src = url;
                });
            }

            getImageDimensions(aImages).then(dimensions => {
                let nWidth = dimensions.width;
                let nHeight = dimensions.height;
                if (nWidth === 0 || nHeight === 0) {
                    const img_prop = new Asc.asc_CImgProperty();
                    img_prop.asc_putImageUrl(sSrc);
                    const or_sz = img_prop.asc_getOriginSize(window['Asc']['editor'] || window['editor']);
                    nWidth = or_sz.Width;
                    nHeight = or_sz.Height;
                } else {
                    nWidth *= AscCommon.g_dKoef_pix_to_mm;
                    nHeight *= AscCommon.g_dKoef_pix_to_mm;
                }
                nWidth = nWidth || 50;
                nHeight = nHeight || 50;
                const oColumnSize = oThis.api.GetDocument().Document.GetColumnSize();
                if (oColumnSize) {
                    if (nWidth > oColumnSize.W || nHeight > oColumnSize.H) {
                        if (oColumnSize.W > 0 && oColumnSize.H > 0) {
                            const dScaleW = oColumnSize.W / nWidth;
                            const dScaleH = oColumnSize.H / nHeight;
                            const dScale = Math.min(dScaleW, dScaleH);
                            nWidth *= dScale;
                            nHeight *= dScale;
                        }
                    }
                }
                console.log(`Width: ${nWidth}px, Height: ${nHeight}px`);
            }).catch(error => {
                console.error(error);
            });
        },

        uploadImage(aImagesToDownload, callback) {
            const oThis = this;
            AscCommon.sendImgUrls(this.api, aImagesToDownload, function (data) {
                const urls = data.map(item => item.url);
                callback && callback(urls);
            });
        },

        goToBookmark(name, hidden = true) {
            this.api.asc_selectSearchingResults(false);
            this._state.isHighlightedResults = false;
            const oDocument = this.api.GetDocument().Document;
            const oParagraph = this.api.asc_GetDocumentOutlineManager().Elements[parseInt(name)];
            oDocument.Search2(oParagraph.Paragraph);
            this.api.asc_GetDocumentOutlineManager().goto(parseInt(name));
            this.api.asc_selectSearchingResults(true);
            this._state.isHighlightedResults = true;
        },
        /**
         * generate catalog number
         * @returns {*[]} catalog array
         * @param pCatalogArray
         */
        generateCatalogNumber(pCatalogArray = []) {
            const MAX_TITLE_LEVELS = 9;
            // 初始化数组，用于保存当前每个层级的编号
            let vCurrentLevels = Array(MAX_TITLE_LEVELS).fill(0);
            // 结果数组，用于保存每个标签的编号
            let vResults = [];
            // 当前的最大级别
            let vMaxLevel = -1;
            for (let i = 0; i < pCatalogArray.length; i++) {
                const vTitleLevel = pCatalogArray[i].titleLevel;
                let vLevel = parseInt(vTitleLevel[vTitleLevel.length - 1]) - 1;
                // 重置该级别以上所有级别的编号
                for (let j = vLevel + 1; j <= vMaxLevel; j++) {
                    vCurrentLevels[j] = 0;
                }
                // 更新当前的最大级别
                vMaxLevel = Math.max(vMaxLevel, vLevel);
                // 增加当前级别的编号
                vCurrentLevels[vLevel]++;
                // 构造当前编号字符串
                let vResStr = "";
                for (let k = 0; k <= vMaxLevel; k++) {
                    if (vCurrentLevels[k] !== 0) {
                        vResStr += vCurrentLevels[k] + '.';
                    }
                }
                // 移除最后的 '.'
                vResStr = vResStr.slice(0, -1);
                // 添加书签
                let vBookmark = this.addBookMark(pCatalogArray[i].Paragraph, vResStr + pCatalogArray[i].titleText);
                // 将当前编号字符串添加到结果数组
                vResults.push({
                    no: vResStr,
                    level: vLevel,
                    text: pCatalogArray[i].titleText,
                    bookmark: vBookmark
                });
            }
            return vResults;

        },
        addBookMark(pParagraph, pName, hidden = true) {
            // 判读pParagraph是否为字符串
            if (typeof pParagraph === 'string') {
                const oCurrentParagraph = this.api.GetDocument().Document.GetCurrentParagraph(true);
                pName = pParagraph;
                pParagraph = oCurrentParagraph;
            }
            if (!pName) {
                return;
            }
            pName = hidden ? "_" + pName : pName;
            let vDocument = pParagraph.Parent;
            if (!(vDocument instanceof CDocument)) {
                vDocument = pParagraph.LogicDocument;
            }
            if (!vDocument.BookmarksManager.HaveBookmark(pName)) {
                // vDocument.BookmarksManager.RemoveBookmark(pName);

                vDocument.StartAction(AscDFH.historydescription_Document_AddBookmark);
                const vBookmarkId = vDocument.BookmarksManager.GetNewBookmarkId();
                pParagraph.AddBookmarkChar(new CParagraphBookmark(false, vBookmarkId, pName), false);
                pParagraph.AddBookmarkChar(new CParagraphBookmark(true, vBookmarkId, pName), false);
                vDocument.Recalculate();
                vDocument.FinalizeAction();

            }
            const vBookmark = vDocument.BookmarksManager.GetBookmarkByName(pName);
            return {
                bookmarkId: vBookmark[0].BookmarkId ? vBookmark[0].BookmarkId : -1,
                bookmarkName: pName
            }
        },
        removeBookmark(pName, hidden = true) {
            const vDocument = this.api.GetDocument().Document;
            pName = hidden ? "_" + pName : pName;
            if (vDocument.BookmarksManager.HaveBookmark(pName)) {
                vDocument.BookmarksManager.RemoveBookmark(pName);
            }
        },
        getParagraphText(pParagraph) {
            // 字符序列
            let catalogString = "";
            let contentArray = pParagraph.Content;
            contentArray.forEach(item => {
                if (item instanceof ParaRun) {
                    // 字符编码数组
                    let charCodeArr = item.Content
                        .filter(item2 => item2.GetType() === para_Text)
                        .map(item2 => item2.Value);
                    if (String.fromCharCode(...charCodeArr)) {
                        catalogString += String.fromCharCode(...charCodeArr);
                    }
                }
            });
            return catalogString;
        },


        getDocumentOutlineList() {
            // 加载文档目录
            this.api.asc_ShowDocumentOutline();
            let count = this.api.asc_GetDocumentOutlineManager().get_ElementsCount(),
                header_level = -1,
                vCatalogList = [];
            const vCDocumentOutline = this.api.asc_GetDocumentOutlineManager();
            const vElements = vCDocumentOutline.Elements;
            for (let i = 0; i < count; i++) {
                let level = this.api.asc_GetDocumentOutlineManager().get_Level(i),
                    hasParent = true;
                if (level >= 3) {
                    continue;
                }
                if (header_level < 0 || level <= header_level) {
                    if (i > 0) {
                        header_level = level;
                    }
                    hasParent = false;
                }
                const vParagraph = vElements[i].Paragraph;
                if (vParagraph) {
                    vCatalogList.push({
                        no: vParagraph.GetNumberingText(),
                        level: level,
                        text: this.api.asc_GetDocumentOutlineManager().get_Text(i),
                        hasParent: hasParent,
                        index: i
                    });
                }
            }
            this.api.asc_HideDocumentOutline();
            return vCatalogList;
        },

        executeCommand(script, callbackString) {
            let callback = new Function('...args', `(function (...args) { ${callbackString} })(...args)`);
            eval(script);
        },

        callCommand(script, callback) {
            function getFunctionBody(func) {
                const functionString = func.toString();
                const bodyStartIndex = functionString.indexOf('{') + 1;
                const bodyEndIndex = functionString.lastIndexOf('}');
                return functionString.slice(bodyStartIndex, bodyEndIndex).trim();
            }

            this.executeCommand(script, getFunctionBody(callback));
        },

        addNewParagraph(bMultiPage, nContentPos) {
            const oDocument = this.api.GetDocument();
            oDocument.Document.AddToContent(nContentPos + 1, this.api.CreateParagraph().Paragraph);
            oDocument.Document.CurPos.ContentPos = nContentPos + 1;
            oDocument.Document.Content[oDocument.Document.CurPos.ContentPos].MoveCursorToStartPos();
        },

        findFirstElementByPage(targetPage) {
            const result = [];
            // 找出目标页的最后一个元素
            const oDocument = this.api.GetDocument();
            const nCount = oDocument.GetElementsCount();
            let nLastIndex = 0;
            for (let i = 0; i < nCount; i++) {
                const Element = oDocument.Document.Content[i];
                if (type_Paragraph === Element.GetType() && undefined !== Element.Get_SectionPr())
                    result.push(Element);
            }
            return result;
        },

        private_Twips2MM(twips) {
            return 25.4 / 72.0 / 20 * twips;
        },
        private_MM2Twips(mm) {
            return mm / (25.4 / 72.0 / 20);
        },

        /**
         * 添加新空页
         */
        addBlankPage(targetPage = 0) {
            // 找出目标页的最后一个元素
            const oDocument = this.api.GetDocument();
            const nCount = oDocument.GetElementsCount();
            let nLastIndex = 0;
            for (let i = 0; i < nCount; i++) {
                const oApiElements = oDocument.GetElement(i);
                if (oApiElements.GetClassType() === 'paragraph' && oApiElements.Paragraph.PageNum === targetPage) {
                    nLastIndex = i;
                }
                if (oApiElements.GetClassType() === 'table' && oApiElements.Table.PageNum === targetPage) {
                    nLastIndex = i;
                }
            }
            // 是否是有多个页面
            const bMultiPage = this.api.GetDocument().Document.FullRecalc.PageIndex;
            // TODO nLastIndex原本应该是第一页的最后一个元素，未知原因导致nLastIndex不是第一页的最后一个元素
            let oLastParagraph = bMultiPage ? oDocument.GetElement(nLastIndex - 1 < 0 ? 0 : nLastIndex - 1) : oDocument.GetElement(nLastIndex);
            oLastParagraph = oLastParagraph.Paragraph;

            // 开始活动，通知文档开始修改
            // oDocument.Document.StartAction(AscDFH.historydescription_Document_AddBlankPage);

            let nContentPos = oLastParagraph.Index


            oLastParagraph.Content[oLastParagraph.CurPos.ContentPos].State.ContentPos = oLastParagraph.GetText().length;

            if (oLastParagraph.Next) {
                oDocument.Document.CurPos.ContentPos = nContentPos + 1;
                oDocument.Document.Content[nContentPos + 1].MoveCursorToStartPos();
            }
            oDocument.Document.AddNewParagraph(undefined, true);
            oDocument.Document.CurPos.ContentPos = nContentPos + 1;
            oDocument.Document.Content[nContentPos + 1].MoveCursorToStartPos();
            oDocument.Document.AddNewParagraph(undefined, true);

            if (oLastParagraph.Next && oLastParagraph.Next.IsParagraph()) {
                oLastParagraph.Next.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));
                // 清空样式
                oLastParagraph.Next.Clear_Formatting();
                oDocument.Document.CurPos.ContentPos = nContentPos + 1;
                oDocument.Document.Content[nContentPos + 1].MoveCursorToStartPos();
                // oDocument.Document.Content[oDocument.Document.CurPos.ContentPos].Clear_Formatting();
            }

            // oDocument.Document.Content[oDocument.Document.CurPos.ContentPos].Content[oLastParagraph.CurPos.ContentPos].State.ContentPos = oLastParagraph.GetText().length;
            // 结束活动，通知文档修改结束
            oDocument.Document.Recalculate();
            oDocument.Document.UpdateInterface();
            oDocument.Document.UpdateSelection();
            oDocument.Document.FinalizeAction();
            // 多页面返回被添加break_Page的段落，单页面返回被添加break_Page的段落的下一个段落，目的就是始终返回的是空白页的第一个段落
            if (!bMultiPage) {
                return oLastParagraph.Next ? oLastParagraph.Next.Next : oLastParagraph;
            }
            return oLastParagraph.Next;
        },
        signature: function (oData = {
            review: [
                {
                    text: '核    定：',
                    textFont: '黑体',
                    textFontSize: 32,
                    peoples: ['熊  晶'],
                    peopleFont: '楷体_GB2312',
                    peopleFontSize: 32,
                    peopleSplit: '   ',
                    photos: ['http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg'],
                    imageWidth: 1.6,
                    imageHeight: 0.93,
                },
                {
                    text: '审    定：',
                    textFont: '黑体',
                    textFontSize: 32,
                    peoples: ['王  瑾'],
                    peopleFont: '楷体_GB2312',
                    peopleFontSize: 32,
                    peopleSplit: '   ',
                    photos: ['http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg'],
                    imageWidth: 1.6,
                    imageHeight: 0.93,
                },
                {
                    text: '审    查：',
                    textFont: '黑体',
                    textFontSize: 32,
                    peoples: ['王瑞锋', '胡  琴', '程  强', '牟松芳', '刘详刚', '王瑞锋', '胡  琴', '程  强', '牟松芳', '刘详刚'],
                    peopleFont: '楷体_GB2312',
                    peopleFontSize: 32,
                    peopleSplit: '   ',
                    photos: [
                        'http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg',
                        'http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg',
                        'http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg',
                        'http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg',
                        'http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg',
                        'http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg',
                        'http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg',
                    ],
                    imageWidth: 1.6,
                    imageHeight: 0.93,
                },
                {
                    text: '校    核：',
                    textFont: '黑体',
                    textFontSize: 32,
                    peoples: ['吕全伟', '王  洋', '唐  榕', '冯  旭', '徐睿志', '常  强'],
                    peopleFont: '楷体_GB2312',
                    peopleFontSize: 32,
                    peopleSplit: '   ',
                    photos: ['http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg'],
                    imageWidth: 1.6,
                    imageHeight: 0.93,
                },
                {
                    text: '编    写：',
                    textFont: '黑体',
                    textFontSize: 32,
                    peoples: ['郭应文', '付  华', '卢  俊', '朱继辉', '谢杰君', '穆  远'],
                    peopleFont: '楷体_GB2312',
                    peopleFontSize: 32,
                    peopleSplit: '   ',
                    photos: ['http://image.hi-hufei.com/eac4b74543a98226d07efeeb8a82b9014b90ebbb.jpg'],
                    imageWidth: 1.6,
                    imageHeight: 0.93,
                },
                {
                    text: '工作人员：',
                    textFont: '黑体',
                    textFontSize: 32,
                    peoples: ['刘小平'],
                    peopleFont: '楷体_GB2312',
                    peopleFontSize: 32,
                    peopleSplit: '   ',
                    photos: [],
                    imageWidth: 1.6,
                    imageHeight: 0.93,
                },
                {
                    text: '商务联系电话：',
                    textFont: '黑体',
                    textFontSize: 32,
                    peoples: ['0851-85388459'],
                    peopleFont: '楷体_GB2312',
                    peopleFontSize: 32,
                    peopleSplit: '   ',
                    photos: [],
                    imageWidth: 1.6,
                    imageHeight: 0.93,
                },
            ]

        }) {
            /**
             * 计算twip
             * @param oRun 文本对象
             * @returns {undefined} twip 表示点的二十分之一（1/1440）的宽度
             */
            function calculateIndent(oRun) {
                // 毫米
                let nMM = 0;
                oRun.Run.Content.forEach(oContent => {
                    nMM += oContent.GetWidth()
                });
                return nMM / 10 * 1440 / 2.54;
            }

            /**
             * 将数组分割成指定大小的块
             * @param array 数组
             * @param chunkSize 块大小
             * @returns {*[]} 分割后的数组
             */
            function splitArrayIntoChunks(array, chunkSize) {
                const result = [];
                for (let i = 0; i < array.length; i += chunkSize) {
                    result.push(array.slice(i, i + chunkSize));
                }
                return result;
            }

            function createNextParagraph(oDocument, nIndex) {
                oDocument.Document.CurPos.ContentPos = nIndex;
                oDocument.Document.Content[oDocument.Document.CurPos.ContentPos].MoveCursorToStartPos();
                oDocument.Document.AddNewParagraph(undefined, true);
                oDocument.Document.CurPos.ContentPos = nIndex + 1;
                oDocument.Document.Content[oDocument.Document.CurPos.ContentPos].MoveCursorToStartPos();
                return oDocument.Document.CurPos.ContentPos;
            }

            function createTable(audit, oLastParagraph) {

                // 插入位置
                let nIndex = oLastParagraph.Index + 1;
                // 创建下一个段落
                nIndex = createNextParagraph(oDocument, nIndex);
                const oTable = this.api.CreateTable(2, audit.length);
                for (let i = 0; i < audit.length; i++) {
                    const oObj = audit[i];
                    // 设置text字体
                    const oTextPr = this.api.CreateTextPr();
                    oTextPr.SetFontSize(oObj.textFontSize);
                    oTextPr.SetFontFamily(oObj.textFont);

                    const oRun = oTable.GetCell(i, 0).GetContent().GetElement(0).AddText(oObj.text);
                    todoTableRun[i] = oRun;
                    if (!todoTableRun[-1]) {
                        todoTableRun[-1] = oTable;
                    }
                    oTable.GetCell(i, 0).GetContent().GetElement(0).SetTextPr(oTextPr);
                    oObj.photos.forEach((photo, index) => {
                        const oRunDrawing = this.api.CreateRun();
                        // 1.6cm * 0.93cm
                        const oDrawing = this.api.CreateImage(photo, oObj.imageWidth * 10 * 36000, oObj.imageHeight * 10 * 36000);
                        oRunDrawing.AddDrawing(oDrawing);
                        oRunDrawing.AddText(oObj.peopleSplit);
                        oRunDrawing.SetTextPr(oTextPr);
                        oTable.GetRow(i).GetCell(1).GetContent().GetElement(0).AddElement(oRunDrawing);
                        // 最后一行换行
                        if (index === oObj.photos.length - 1) {
                            oTable.GetRow(i).GetCell(1).GetContent().GetElement(0).AddLineBreak();
                        }
                    });
                }
                const oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
                const oTablePr = oTableStyle.GetTablePr();
                oTablePr.SetTableBorderBottom("single", 0, 0, 255, 255, 255);
                oTablePr.SetTableBorderLeft("single", 0, 0, 255, 255, 255);
                oTablePr.SetTableBorderRight("single", 0, 0, 255, 255, 255);
                oTablePr.SetTableBorderTop("single", 0, 0, 255, 255, 255);
                oTablePr.SetTableBorderInsideH("single", 0, 0, 255, 255, 255);
                oTablePr.SetTableBorderInsideV("single", 0, 0, 255, 255, 255);
                oTablePr.SetWidth("percent", 100);
                oTable.SetStyle(oTableStyle);
                oDocument.AddElement(nIndex, oTable);
                return nIndex;
            }

            function createTextAndStyle(audit, oLastParagraph) {
                // 插入位置
                let nIndex = oLastParagraph.Index + 1;
                for (let i = 0; i < audit.length; i++) {
                    const oObj = audit[i];
                    // 创建下一个段落
                    nIndex = createNextParagraph(oDocument, nIndex);
                    const oApiParagraph = oDocument.GetElement(nIndex - 1);
                    oApiParagraph.Paragraph.Clear_Formatting();
                    const oRun = this.api.CreateRun();
                    oRun.AddText(oObj.text);

                    // 设置text字体
                    const oTextPr = this.api.CreateTextPr();
                    oTextPr.SetFontSize(oObj.textFontSize);
                    oTextPr.SetFontFamily(oObj.textFont);
                    oRun.SetTextPr(oTextPr);

                    todoRun.push(oRun);
                    todoParagraph[i] = [];
                    // 添加text
                    oApiParagraph.AddElement(oRun);


                    // 设置people字体
                    const oPeopleTextPr = this.api.CreateTextPr();
                    oPeopleTextPr.SetFontSize(oObj.peopleFontSize);
                    oPeopleTextPr.SetFontFamily(oObj.peopleFont);

                    // 添加人员 每行4个人
                    const aSplitArrayIntoChunks = splitArrayIntoChunks(oObj.photos, 4);
                    const aLastSplitArrayIntoChunks = aSplitArrayIntoChunks[aSplitArrayIntoChunks.length - 1];
                    aSplitArrayIntoChunks.forEach((photos, index) => {
                        oApiParagraph.SetJc("left");
                        const oRunDrawing = this.api.CreateRun();
                        photos.forEach(photo => {
                            // 1.6cm * 0.93cm
                            const oDrawing = this.api.CreateImage(photo, oObj.imageWidth * 10 * 36000, oObj.imageHeight * 10 * 36000);
                            oRunDrawing.AddDrawing(oDrawing);
                            oRunDrawing.AddText(oObj.peopleSplit);
                        });
                        oRunDrawing.SetTextPr(oPeopleTextPr);


                        if (index === 0) {
                            oApiParagraph.AddElement(oRunDrawing);
                            if (aSplitArrayIntoChunks.length === 1) {
                                oApiParagraph.AddLineBreak();
                            }
                        } else {
                            // 第二组开始，设置缩进
                            // 创建下一个段落
                            nIndex = createNextParagraph(oDocument, nIndex);
                            const oNextApiParagraph = oDocument.GetElement(nIndex - 1);
                            oNextApiParagraph.Paragraph.Clear_Formatting();
                            oNextApiParagraph.AddElement(oRunDrawing);
                            todoParagraph[i].push(oNextApiParagraph);
                            // 最后一组设置换行
                            if (aLastSplitArrayIntoChunks === photos) {
                                oNextApiParagraph.AddLineBreak();
                            }
                        }

                    });

                    if (i === audit.length - 1 || i === audit.length - 2) {
                        const oPeopleRun = this.api.CreateRun();
                        oObj.peoples.forEach((people, index) => {
                            oPeopleRun.AddText(people);
                            oPeopleRun.AddText(oObj.peopleSplit);
                        });
                        oPeopleRun.SetTextPr(oPeopleTextPr);
                        oApiParagraph.AddElement(oPeopleRun);
                        oApiParagraph.SetJc("left");
                        oApiParagraph.AddLineBreak();
                    }
                    // 添加人员 每行4个人
                    // const aSplitArrayIntoChunks = splitArrayIntoChunks(oObj.peoples, 4);
                    // const aLastSplitArrayIntoChunks = aSplitArrayIntoChunks[aSplitArrayIntoChunks.length - 1];
                    // aSplitArrayIntoChunks.forEach((peoples, index) => {
                    //     const oPeopleRun = this.api.CreateRun();
                    //     peoples.forEach((people, index) => {
                    //         oPeopleRun.AddText(people);
                    //         oPeopleRun.AddText(oObj.peopleSplit);
                    //     });
                    //     oPeopleRun.SetTextPr(oPeopleTextPr);
                    //
                    //     if (index === 0) {
                    //         oApiParagraph.AddElement(oPeopleRun);
                    //         oApiParagraph.SetJc("left");
                    //         if (aSplitArrayIntoChunks.length === 1) {
                    //             if (oObj.photos.length <= 0) {
                    //                 oApiParagraph.AddLineBreak();
                    //             }
                    //             renderImage.call(this, oObj, nIndex, function (oNextApiParagraph, index) {
                    //                 todoParagraph[i].push(oNextApiParagraph);
                    //                 nIndex = index;
                    //             });
                    //         }
                    //     } else {
                    //         // 第二组开始，设置缩进
                    //         // 创建下一个段落
                    //         nIndex = createNextParagraph(oDocument, nIndex);
                    //         const oNextApiParagraph = oDocument.GetElement(nIndex - 1);
                    //         oNextApiParagraph.Paragraph.Clear_Formatting();
                    //         oNextApiParagraph.AddElement(oPeopleRun);
                    //         oNextApiParagraph.SetJc("left");
                    //         todoParagraph[i].push(oNextApiParagraph);
                    //         // 最后一组设置换行
                    //         if (aLastSplitArrayIntoChunks === peoples) {
                    //             renderImage.call(this, oObj, nIndex, function (oNextApiParagraph, index) {
                    //                 todoParagraph[i].push(oNextApiParagraph)
                    //                 nIndex = index;
                    //             });
                    //             if (oObj.photos.length <= 0) {
                    //                 oNextApiParagraph.AddLineBreak();
                    //             }
                    //         }
                    //     }
                    //
                    // });

                }
                return nIndex;
            }

            function renderImage(oData, nIndex, callback) {

                // 设置people字体
                const oPeopleTextPr = this.api.CreateTextPr();
                oPeopleTextPr.SetFontSize(oData.peopleFontSize);
                oPeopleTextPr.SetFontFamily(oData.peopleFont);
                // 添加人员 每行4个人
                const aSplitArrayIntoChunks = splitArrayIntoChunks(oData.photos, 4);
                aSplitArrayIntoChunks.forEach((photos, index) => {
                    // 创建下一个段落
                    nIndex = createNextParagraph(oDocument, nIndex);
                    let oNextApiParagraph = oDocument.GetElement(nIndex - 1);
                    oNextApiParagraph.Paragraph.Clear_Formatting();
                    oNextApiParagraph.SetJc("left");
                    const oRunDrawing = this.api.CreateRun();
                    photos.forEach(photo => {
                        // 1.6cm * 0.93cm
                        const oDrawing = this.api.CreateImage(photo, oData.imageWidth * 10 * 36000, oData.imageHeight * 10 * 36000);
                        oRunDrawing.AddDrawing(oDrawing);
                        oRunDrawing.AddText(oData.peopleSplit);
                    });
                    oRunDrawing.SetTextPr(oPeopleTextPr);
                    oNextApiParagraph.AddElement(oRunDrawing)
                    callback(oNextApiParagraph, nIndex);

                    if (photos === aSplitArrayIntoChunks[aSplitArrayIntoChunks.length - 1]) {
                        oNextApiParagraph.AddLineBreak();
                    }
                });
            }

            if (!oData.review) {
                return
            }
            const oDocument = this.api.GetDocument();
            // 如果文档没有计算完成，等待10毫秒后再次执行
            if (!oDocument.Document._Calculated) {
                setTimeout(this.signature.bind(this, oData), 10);
                return;
            }
            oDocument.Document.StartAction(AscDFH.historydescription_Document_AddCaption);
            let oBreakParagraph = this.addBlankPage(0);
            // 第一页内容是否填满
            const bFirstFull = oBreakParagraph.PageNum !== oBreakParagraph.Next.PageNum;
            if (bFirstFull) {
                oBreakParagraph = oBreakParagraph.Prev;
                oDocument.RemoveElement(oBreakParagraph.Next.Index);
            }
            console.log("oBreakParagraph", oBreakParagraph.PageNum);
            const aFont = new Set();
            oData.review.forEach(item => {
                aFont.add(item.textFont);
                aFont.add(item.peopleFont);
            });
            const aPrepeareFonts = [];
            const mergedArray = [...new Set(oData.review.map(item => item.photos).flat())];
            const aPrepeareImages = mergedArray.reduce((acc, curr, index) => {
                acc[index] = curr;
                return acc;
            }, {});

            aFont.forEach(font => {
                aPrepeareFonts.push(new AscFonts.CFont(font, 0, "", 0));
            })
            let todoRun = [];
            let todoParagraph = {};
            const todoTableRun = {};
            // oData.review 拆成两份 第一个数组是除去后2份，第二个数组是最后2份
            // 复制 oData.review 数组
            const reviewCopy = oData.review.slice();
            // 去除最后两个元素的数组
            const firstArray = reviewCopy.slice(0, reviewCopy.length - 2);
            // 最后两个元素的数组
            const secondArray = reviewCopy.slice(reviewCopy.length - 2);
            // 绑定this
            // 设置光标位置
            oDocument.Document.CurPos.ContentPos = createTable.call(this, firstArray, oBreakParagraph);
            oDocument.Document.CurPos.ContentPos = createTextAndStyle.call(this, secondArray, {Index: oDocument.Document.CurPos.ContentPos});
            oDocument.Document.Content[oDocument.Document.CurPos.ContentPos].MoveCursorToStartPos(false);

            this.api.pre_Paste(aPrepeareFonts, aPrepeareImages, function () {
                // oDocument.Document.Recalculate();
                // todoRun.forEach((run, index) => {
                //     if (todoParagraph[index].length > 0) {
                //         todoParagraph[index].forEach(paragraph => {
                //             console.log("SetIndFirstLine", calculateIndent(run))
                //             paragraph.SetIndFirstLine(calculateIndent(run));
                //         });
                //     }
                // });
                const oTable = todoTableRun[-1];
                delete todoTableRun[-1];
                Object.keys(todoTableRun).forEach(function (key) {
                    let run = todoTableRun[key];
                    oTable.GetRow(key).GetCell(0).SetWidth("twips", calculateIndent(run) + 243);
                });
                // 结束活动，通知文档修改结束
                oDocument.Document.Recalculate();
                oDocument.Document.UpdateInterface();
                oDocument.Document.UpdateSelection();
                oDocument.Document.FinalizeAction();
            });


        },
        /**
         * 获取目录列表
         * @param isActive 是否是主动发起（是否是由postmessage发起请求）
         * @returns {*[]} 目录列表
         */
        getCatalogList(isActive = false) {
            // const vCatalogList = [];
            // // 获取文档对象
            // const vDocument = this.api.GetDocument();
            // const vStyles = vDocument.Document.GetStyles();
            // // 获取标题集合 TODO vHeadings目前只能识别默认的标题样式 1-9 需要识别需要迭代 vStyle.Style所有数组找出所有关键词
            // const vHeadings = vStyles.Default.Headings;
            // for (let i = 0; i < vDocument.GetElementsCount(); i++) {
            //     const vElement = vDocument.GetElement(i);
            //     // 判断是否是标题元素
            //     if (vElement.GetClassType() === 'paragraph' && vElement.GetStyle() && vHeadings.includes(vElement.GetStyle().Style.Id)) {
            //         let vParagraph = vElement.Paragraph;
            //         let vCatalogString = this.getParagraphText(vParagraph);
            //         vCatalogList.push({
            //             titleLevel: vStyles.Get(vElement.GetStyle().Style.Id).Name,
            //             titleText: vCatalogString,
            //             Paragraph: vParagraph,
            //         });
            //     }
            // }
            // const res = this.generateCatalogNumber(vCatalogList);
            const res = this.getDocumentOutlineList();
            if (isActive) {
                // 主动发起请求 返回结果
                Common.Gateway.getCatalogListFunResult(res);
            }


            // const oDocument = this.api.GetDocument();
            // const nCount = oDocument.GetElementsCount();
            // Object.defineProperty(oDocument.GetElement(nCount-1).Paragraph, "PageNum", {
            //     set(newValue) {
            //         console.log("set", newValue);
            //         oDocument.GetElement(nCount-1).Paragraph.PageNum = newValue;
            //     }
            // });
            return res;
        },
        switchDisableEditor: function (disable = true) {
            this.getApplication().getController('Main').disableEditing(disable);
            this.switchCursor(disable);
            if (disable) {
                this.api.asc_addRestriction(Asc.c_oAscRestrictionType.View);
            } else {
                this.api.asc_removeRestriction(Asc.c_oAscRestrictionType.View);
            }

            // Common.NotificationCenter.trigger('protect:doclock', props);
            // window['AscCommon'].g_inputContext.setReadOnlyWrapper(disable);
            // window['AscCommon'].g_inputContext.setReadOnly(disable);
            // window.DE.controllers.DocumentHolder.SetDisabled(disable);
            // window.DE.controllers.Main.disableEditing(disable);
            // this.setDisabledKeyboard(disable);
        },

        switchCursor: function (disable = true) {
            const cursor = document.querySelector('#id_viewer_overlay');
            if (disable) {
                cursor.style.cursor = 'not-allowed';
            } else {
                // delete cursor.style.cursor;
                cursor.style.cursor = 'auto';
            }
        },
        bindListenerCursorEvent: function () {
            // 选择要观察的目标节点
            const cursor = document.querySelector('#id_target_cursor');

            // 创建一个观察器实例并传入回调函数
            this.observer = new MutationObserver((mutationsList, observer) => {
                // 遍历每个变化的属性
                for (const mutation of mutationsList) {
                    if (mutation.type === 'attributes' && mutation.attributeName === 'style') {
                        const oDocument = this.api.GetDocument();
                        // let controllerGetCurPage = oDocument.Document.controller_GetCurPage();
                        // let xy = oDocument.Document.controller_GetCursorPosXY();
                        const Item = oDocument.Document.Content[this.api.GetDocument().Document.CurPos.ContentPos]
                        Common.Gateway.onCursorChange({lineIndex: Item.Index, pageNum: Item.PageNum});
                    }
                }
            });

            // 配置观察选项
            const config = {attributes: true, attributeFilter: ['style']};
            // 传入目标节点和观察选项
            this.observer.observe(cursor, config);
        },
        unbindListenerCursorEvent: function () {
            if (this.observer) {
                this.observer.disconnect();
                this.observer = null;
            }
        },
        getBookmarkPosition: function (array) {
            array = JSON.parse(array);
            // [{startBookmarkName: '', endBookmarkName: ''}] 帮我变量所有调用GetBookmarkByName获取Index和PageNum并存入该数组
            array.forEach(item => {
                if (!item.startBookmarkName || !item.endBookmarkName) {
                    console.error('startBookmarkName or endBookmarkName not null');
                }
                let startBookmark = this.api.asc_GetDocumentOutlineManager().Elements[+item.startBookmarkName];
                let endBookmark = this.api.asc_GetDocumentOutlineManager().Elements[+item.endBookmarkName];
                if (!startBookmark || !endBookmark) {
                    return null;
                }
                // 是否是最后一个元素
                if (+item.startBookmarkName === this.api.asc_GetDocumentOutlineManager().get_ElementsCount() - 1) {
                    item.isLast = true;
                }
                item.startLineIndex = startBookmark.Paragraph.Index;
                item.startPageNum = startBookmark.Paragraph.PageNum;
                item.endLineIndex = endBookmark.Paragraph.Index - 1;
                item.endPageNum = endBookmark.Paragraph.PageNum;
            });
            Common.Gateway.getBookmarkPositionFunResult(array);
            return array;
        },
        getSettings: function (text) {
            var props = this.api.asc_GetWatermarkProps();
            props.Type = 1;
            var val = props.get_Type();
            if (val === Asc.c_oAscWatermarkType.Image) {
                val = parseInt(this.cmbScale.getValue());
                isNaN(val) && (val = -1);
                props.put_Scale((val < 0) ? val : val / 100);
            } else {
                props.put_Text(text);
                props.put_IsDiagonal(true);
                props.put_Opacity(128);

                val = props.get_TextPr() || new Asc.CTextProp();
                if (val) {
                    val.put_FontSize(105);
                    var font = new AscCommon.asc_CTextFontFamily();
                    font.put_Name('Arial');
                    font.put_Index(-1);
                    val.put_FontFamily(font);
                    val.put_Bold(false);
                    val.put_Italic(false);
                    val.put_Underline(false);
                    val.put_Strikeout(false);

                    val.put_Lang(parseInt(Common.util.LanguageInfo.getLocalLanguageCode('zh-CN')));

                    var color = new Asc.asc_CColor();
                    if (this.isAutoColor) {
                        color.put_auto(true);
                    } else {
                        color = Common.Utils.ThemeColor.getRgbColor('c0c0c0');
                    }
                    val.put_Color(color);
                    props.put_TextPr(val);
                }
            }

            return props;
        },
        setWatermark(text, type = 1) {
            const props = this.getSettings(text);
            this.api.asc_SetWatermarkProps(props);
        },
        removeWatermark() {
            this.api.asc_WatermarkRemove();
        },

    }), DE.Controllers.Banshion);
});
