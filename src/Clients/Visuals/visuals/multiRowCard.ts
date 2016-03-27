/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved. 
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *   
 *  The above copyright notice and this permission notice shall be included in 
 *  all copies or substantial portions of the Software.
 *   
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

/// <reference path="../_references.ts"/>

module powerbi.visuals {
    import getKpiImageMetadata = powerbi.visuals.KpiUtil.getKpiImageMetadata;

    export interface CardItemData {
        caption: string;
        details: string;
        showURL: boolean;
        showImage: boolean;
        showKPI: boolean;
        columnIndex: number;
    }

    export interface CardData {
        title?: string;
        showTitleAsURL?: boolean;
        showTitleAsImage?: boolean;
        showTitleAsKPI?: boolean;
        cardItemsData: CardItemData[];
    }

    interface ImageStyle {
        maxWidth?: number;
        maxHeight?: number;
    }

    interface MediaQuery {
        maxWidth?: number;
        style?: MultiRowCardStyle;
    }

    interface MultiRowCardStyle {
        card?: {
            marginBottom?: number;
            maxRows?: number;
        };
        cardItemContainer?: {
            paddingRight?: number;
            minWidth?: number;
        };
        caption?: {
            lineHeight?: number;
        };
        imageCaption?: ImageStyle;
        imageTitle?: ImageStyle;
    }

    export class MultiRowCard implements IVisual {
        private currentViewport: IViewport;
        private options: VisualInitOptions;
        private dataView: DataView;
        private style: IVisualStyle;
        private element: JQuery;
        private listView: IListView;
        /**
         * This includes card height with margin that will be passed to list view.
         */
        private settings: MultiRowCardStyle;
        private dataModel: CardData[];
        private interactivity: InteractivityOptions;
        private isInteractivityOverflowHidden: boolean = false;
        private waitingForData: boolean;
        private cardHasTitle: boolean;
        private isSingleRowCard: boolean;
        private maxColPerRow: number;
        
        /**
         * Note: Public for testability.
         */
        public static formatStringProp: DataViewObjectPropertyIdentifier = {
            objectName: 'general',
            propertyName: 'formatString',
        };
        private static multiRowCardClass = 'multiRowCard';
        private static Card: ClassAndSelector = {
            class: 'card',
            selector: '.card'
        };

        private static Title: ClassAndSelector = {
            class: 'title',
            selector: '.title'
        };

        private static ImageTitle: ClassAndSelector = {
            class: 'title',
            selector: MultiRowCard.Title.selector + ' img'
        };

        private static KPITitle: ClassAndSelector = {
            class: 'kpiTitle',
            selector: MultiRowCard.Title.selector + ' kpiTitle'
        };

        private static CardItemContainer: ClassAndSelector = {
            class: 'cardItemContainer',
            selector: '.cardItemContainer'
        };

        private static Caption: ClassAndSelector = {
            class: 'caption',
            selector: '.caption'
        };

        private static ImageCaption: ClassAndSelector = {
            class: '',
            selector: MultiRowCard.Caption.selector + ' img'
        };

        private static Details: ClassAndSelector = {
            class: 'details',
            selector: '.details'
        };
        private static TitleUrlSelector: string = MultiRowCard.Title.selector + ' a';
        private static CaptionUrlSelector: string = MultiRowCard.Caption.selector + ' a';
        /**
         * Cards have specific styling so defined inline styles and also to support theming and improve performance.
         */
        private static DefaultStyle: MultiRowCardStyle = {
            card: {
                marginBottom: 20
            },
            cardItemContainer: {
                paddingRight: 20,
                minWidth: 120,
            },
            caption: {
                lineHeight: 20,
            },
            imageCaption: {
                maxHeight: 75,
                maxWidth: 100,
            },
            imageTitle: {
                maxHeight: 75,
                maxWidth: 100,
            }
        };

        // queries should be ordered by maxWidth in ascending order
        private static tileMediaQueries: MediaQuery[] = [
            {
                maxWidth: 250,
                style: {
                    card: {
                        maxRows: 2,
                    },
                    cardItemContainer: {
                        minWidth: 110,
                    },
                    imageCaption: {
                        maxHeight: 45,
                    }
                }
            },
            {
                maxWidth: 490,
                style: {
                    card: {
                        maxRows: 2,
                    },
                    cardItemContainer: {
                        minWidth: 130,
                    },
                    caption: {
                        lineHeight: 24,
                    },
                    imageCaption: {
                        maxHeight: 52,
                    }
                }
            },
            {
                maxWidth: 750,
                style: {
                    card: {
                        maxRows: 1,
                    },
                    cardItemContainer: {
                        minWidth: 120,
                    },
                    caption: {
                        lineHeight: 20,
                    },
                    imageCaption: {
                        maxHeight: 53,
                    }
                }
            }
        ];

        public init(options: VisualInitOptions) {
            debug.assertValue(options, 'options');
            this.options = options;
            this.style = options.style;
            let viewport = this.currentViewport = options.viewport;
            let interactivity = this.interactivity = options.interactivity;

            if (interactivity && interactivity.overflow === 'hidden')
                this.isInteractivityOverflowHidden = true;

            this.settings = $.extend(true, {}, MultiRowCard.DefaultStyle);
            var multiRowCardDiv = this.element = $('<div/>')
                .addClass(MultiRowCard.multiRowCardClass)
                .css({
                    'height': getPixelString(viewport.height),
                });
            options.element.append(multiRowCardDiv);
            this.initializeCardRowSelection();
        }

        public onDataChanged(options: VisualDataChangedOptions): void {
            debug.assertValue(options, 'options');

            let dataViews = options.dataViews;
            if (dataViews && dataViews.length > 0) {
                var dataView = this.dataView = dataViews[0];
                var columnMetadata: DataViewMetadataColumn[] = dataView.table.columns;
                var tableRows: any[][] = dataView.table.rows;
                var resetScrollbarPosition = options.operationKind !== VisualDataChangeOperationKind.Append;
                var dataModel = this.dataModel = MultiRowCard.converter(dataView, columnMetadata.length, tableRows.length, this.isInteractivityOverflowHidden);
                this.setCardDimensions();
                this.listView.data(dataModel, (d: CardData) => dataModel.indexOf(d), resetScrollbarPosition);
            }

            this.waitingForData = false;
        }

        public onResizing(viewport: IViewport): void {
            var heightNotChanged = (this.currentViewport.height === viewport.height);
            this.currentViewport = viewport;
            this.element.css('height', getPixelString(viewport.height));
            if (!this.dataView)
                return;

            var previousMaxColPerRow = this.maxColPerRow;
            this.maxColPerRow = this.getMaxColPerRow();
            var widthNotChanged = (previousMaxColPerRow === this.maxColPerRow);
            if (heightNotChanged && widthNotChanged)
                return;

            this.listView.viewport(viewport);
        }

        public static converter(dataView: DataView, columnCount: number, maxCards: number, isDashboardVisual: boolean = false): CardData[] {
            let details: CardData[] = [];
            let tableDataRows = dataView.table.rows;
            let columnMetadata: DataViewMetadataColumn[] = dataView.table.columns;

            for (let i = 0, len = maxCards; i < len; i++) {
                let row = tableDataRows[i];
                let isValuePromoted: boolean = undefined;
                var title: string = undefined;
                let showTitleAsURL: boolean = false;
                let showTitleAsImage: boolean = false;
                let showTitleAsKPI: boolean = false;
                let cardData: CardItemData[] = [];
                for (var j = 0; j < columnCount; j++) {
                    let column = columnMetadata[j];

                    let statusGraphicInfo = getKpiImageMetadata(column, row[j]);
                    let columnCaption: string;
                    let statusGraphic: string;

                    if (statusGraphicInfo) {
                        columnCaption = statusGraphicInfo.caption;
                        statusGraphic = statusGraphicInfo.statusGraphic;
                    }

                    //TODO: seems we are duplicating this logic in many places. Consider putting it in KPIUtil
                    if (!columnCaption)
                        columnCaption = valueFormatter.format(row[j], valueFormatter.getFormatString(column, MultiRowCard.formatStringProp));

                    let showKPI = statusGraphicInfo !== undefined && statusGraphicInfo.caption !== undefined;

                    // The columnDetail represents column name. In card the column name is shown as details
                    let columnDetail: string = columnMetadata[j].displayName;

                    //Title is shown only on Canvas and only if there is one Category field.
                    if (!isDashboardVisual && !column.type.numeric) {
                        if (isValuePromoted === undefined) {
                            isValuePromoted = true;
                            title = columnCaption;
                            showTitleAsURL = UrlHelper.isValidUrl(column, title);
                            showTitleAsImage = UrlHelper.isValidImage(column, columnCaption);
                            showTitleAsKPI = showKPI;
                        }
                        else if (isValuePromoted) {
                            isValuePromoted = false;
                        }
                    }
                    cardData.push({
                        caption: columnCaption,
                        details: columnDetail,
                        showURL: UrlHelper.isValidUrl(column, columnCaption),
                        showImage: UrlHelper.isValidImage(column, columnCaption),
                        showKPI: showKPI,
                        columnIndex: j
                    });
                }
                details.push({
                    title: isValuePromoted ? title : undefined,
                    showTitleAsURL: showTitleAsURL,
                    showTitleAsImage: showTitleAsImage,
                    showTitleAsKPI: showTitleAsKPI,
                    cardItemsData: isValuePromoted ? cardData.filter((d: CardItemData) => d.caption !== title) : cardData
                });
            }
            return details;
        }

        private initializeCardRowSelection() {
            let isDashboardVisual = this.isInteractivityOverflowHidden;

            let rowEnter = (rowSelection: D3.Selection) => {
                let cardRow = rowSelection
                    .append("div")
                    .classed(MultiRowCard.Card.class, true);

                // The card top padding is not needed when card items are wrapped as top padding is added to each carditemcontainer when wrapped
                if (isDashboardVisual) {
                    cardRow.classed('mrtile', true);
                }
                else {
                    if (this.cardHasTitle) {
                        cardRow.append("div").classed(MultiRowCard.Title.class, true)
                            .each(function (d: CardData) {
                                if (d.showTitleAsImage)
                                    appendImage(d3.select(this));
                                else if (d.showTitleAsURL)
                                    d3.select(this).append('a');
                                else if (d.showTitleAsKPI)
                                    d3.select(this).append('div')
                                        .classed(MultiRowCard.KPITitle.class, true)
                                        .classed(d.title, true)
                                        .style({
                                            display: 'inline-block',
                                            verticalAlign: 'sub'
                                        });
                            });
                    }
                }

                let cardItem = cardRow
                    .selectAll(MultiRowCard.CardItemContainer.selector)
                    .data((d: CardData) => d.cardItemsData)
                    .enter()
                    .append('div')
                    .classed(MultiRowCard.CardItemContainer.class, true);

                cardItem
                    .append('div')
                    .classed(MultiRowCard.Caption.class, true)
                    .each(function (d: CardItemData) {
                        if (d.showURL) {
                            d3.select(this).append('a');
                        }
                        else if (d.showImage) {
                            appendImage(d3.select(this));
                        }
                        else if (d.showKPI) {
                            d3.select(this).append('div')
                                .classed(d.caption, true)
                                .style({
                                    display: 'inline-block',
                                    verticalAlign: 'sub'
                                });
                        }
                    });

                cardItem
                    .append('div')
                    .classed(MultiRowCard.Details.class, true);
            };

            /**
            * Row update should:
            * 1. bind Data
            * 2. Manipulate DOM (likely just updating CSS properties) affected by data
            */
            let rowUpdate = (rowSelection: D3.Selection) => {
                let style = this.getStyle();
                if (!isDashboardVisual && this.cardHasTitle) {
                    rowSelection.selectAll(MultiRowCard.Title.selector).filter((d: CardData) => !d.showTitleAsURL && !d.showTitleAsImage && !d.showTitleAsKPI).text((d: CardData) => d.title);
                    rowSelection
                        .selectAll(MultiRowCard.TitleUrlSelector)
                        .attr('href', (d: CardData) => d.title)
                        .attr('target', '_blank')
                        .text((d: CardData) => d.title);

                    rowSelection
                        .selectAll(MultiRowCard.ImageTitle.selector)
                        .attr('src', (d: CardData) => d.title);
                    setImageStyle(rowSelection.selectAll(MultiRowCard.Title.selector), style.imageTitle);

                    rowSelection
                        .selectAll(MultiRowCard.KPITitle.selector)
                        .each(function (d: CardData) {
                            var element = d3.select(this);
                            element.classed(d.title);
                        });
                }

                let cardSelection = rowSelection.selectAll(MultiRowCard.Card.selector);

                cardSelection
                    .selectAll(MultiRowCard.Caption.selector)
                    .filter((d: CardItemData) => !(d.showURL || d.showImage || d.showKPI))
                    .text((d: CardItemData) => d.caption)
                    .style({
                        'line-height': getPixelString(style.caption.lineHeight),
                    });

                cardSelection
                    .selectAll(MultiRowCard.ImageCaption.selector)
                    .attr('src', (d: CardItemData) => d.caption)
                    .style(style.imageCaption);

                cardSelection
                    .selectAll(MultiRowCard.CardItemContainer.selector)
                    .style({
                        'padding-right': (d: CardItemData) => {
                            return this.isLastRowItem(d.columnIndex, this.dataView.metadata.columns.length) ? '0px' : getPixelString(style.cardItemContainer.paddingRight);
                        },
                        'width': (d: CardItemData) => {
                            return this.getColumnWidth(d.columnIndex, this.dataView.metadata.columns.length);
                        },
                        'display': (d: CardItemData) => {
                            return (this.hideColumn(d.columnIndex) ? 'none' : 'inline-block');
                        },
                    });

                setImageStyle(cardSelection.selectAll(MultiRowCard.Caption.selector), style.imageCaption);

                cardSelection
                    .selectAll(MultiRowCard.CaptionUrlSelector)
                    .attr('href', (d: CardItemData) => d.caption)
                    .attr('target', '_blank')
                    .text((d: CardItemData) => d.caption);

                cardSelection
                    .selectAll(MultiRowCard.Details.selector).text((d: CardItemData) => d.details);

                cardSelection
                    .style('margin-bottom', isDashboardVisual ? '0px' : (this.isSingleRowCard ? '0px' : getPixelString(style.card.marginBottom)));
            };

            let rowExit = (rowSelection: D3.Selection) => {
                rowSelection.remove();
            };

            let listViewOptions: ListViewOptions = {
                rowHeight: undefined,
                enter: rowEnter,
                exit: rowExit,
                update: rowUpdate,
                loadMoreData: () => this.onLoadMoreData(),
                viewport: this.currentViewport,
                baseContainer: d3.select(this.element.get(0)),
                scrollEnabled: !this.isInteractivityOverflowHidden,
            };

            this.listView = ListViewFactory.createListView(listViewOptions);
        }

        private getMaxColPerRow(): number {
            var rowWidth = this.currentViewport.width;
            var minColumnWidth = this.getStyle().cardItemContainer.minWidth;
            var columnCount = this.dataView.metadata.columns.length;
            //atleast one column fits in a row
            var maxColumnPerRow = Math.floor(rowWidth / minColumnWidth) || 1;
            return Math.min(columnCount, maxColumnPerRow);
        }

        private getRowIndex(fieldIndex: number): number {
            return Math.floor((fieldIndex * 1.0) / this.getMaxColPerRow());
        }

        private getStyle(): MultiRowCardStyle {
            var defaultStyle = this.settings;
            if (!this.isInteractivityOverflowHidden)
                return $.extend(true, {}, defaultStyle);

            let viewportWidth = this.currentViewport.width;
            let overrideStyle: MultiRowCardStyle = {};
            for (let currentQuery of MultiRowCard.tileMediaQueries)
                if (viewportWidth <= currentQuery.maxWidth) {
                    overrideStyle = currentQuery.style;
                    break;
                }
            return $.extend(true, {}, defaultStyle, overrideStyle);
        }

        private hideColumn(fieldIndex: number): boolean {
            //calculate the number of items apearing in the same row as the columnIndex
            var rowIndex = this.getRowIndex(fieldIndex);

            // when interactivity is disabled (pinned tile), don't wrap the row
            var maxRows = this.getStyle().card.maxRows;
            return (maxRows && rowIndex >= maxRows);
        }

        private getColumnWidth(fieldIndex: number, columnCount: number): string {
            //atleast one column fits in a row
            var maxColumnPerRow = this.getMaxColPerRow();
            if (maxColumnPerRow >= columnCount)
                //all columns fit in the same row, divide the space equaly
                return (100.0 / columnCount) + '%';

            //calculate the number of items apearing in the same row as the columnIndex
            var rowIndex = this.getRowIndex(fieldIndex);

            var totalRows = Math.ceil((columnCount * 1.0) / maxColumnPerRow);
            var lastRowCount = columnCount % maxColumnPerRow;
            if (rowIndex < totalRows || lastRowCount === 0)
                // items is not on the last row or last row contains max columns allowed per row
                return (100.0 / maxColumnPerRow) + '%';

            // items is on the last row
            return (100.0 / lastRowCount) + '%';
        }

        private isLastRowItem(fieldIndex: number, columnCount: number) {
            if (fieldIndex + 1 === columnCount)
                return true;
            var maxColumnPerRow = this.getMaxColPerRow();
            if (maxColumnPerRow - (fieldIndex % maxColumnPerRow) === 1)
                return true;

            return false;
        }

        /**
         * This contains the card column wrapping logic.
         * Determines how many columns can be shown per each row inside a Card.
         * To place the fields evenly along the card,
         * the width of each card item is calculated based on the available viewport width.
         */
        private setCardDimensions(): void {
            this.cardHasTitle = false;

            var dataModel = this.dataModel;

            if (!this.isInteractivityOverflowHidden && dataModel && dataModel.length > 0) {
                this.cardHasTitle = dataModel[0].title !== undefined;
                this.isSingleRowCard = dataModel.length === 1 ? true : false;
            }
        }

        private onLoadMoreData(): void {
            if (!this.waitingForData && this.dataView.metadata && this.dataView.metadata.segment) {
                this.options.host.loadMoreData();
                this.waitingForData = true;
            }
        }
    }

    function appendImage(selection: D3.Selection): void {
        selection
            .append('div')
            .classed('imgCon', true)
            .append('img');
    }

    function setImageStyle(selection: D3.Selection, imageStyle: ImageStyle): void {
        selection
            .selectAll('.imgCon')
            .style({
                'height': getPixelString(imageStyle.maxHeight),
            })
            .selectAll('img')
            .style({
                'max-height': getPixelString(imageStyle.maxHeight),
                'max-width': getPixelString(imageStyle.maxWidth),
            });
    }

    function getPixelString(value: number): string {
        return value + "px";
    }
}
