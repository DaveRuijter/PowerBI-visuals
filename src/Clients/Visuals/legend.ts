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

/// <reference path="_references.ts"/>

module powerbi.visuals {
    export enum LegendIcon {
        Box,
        Circle,
        Line
    }

    export enum LegendPosition {
        Top,
        Bottom,
        Right,
        Left,
        None,
        TopCenter,
        BottomCenter,
        RightCenter,
        LeftCenter,          
    }

    export interface LegendPosition2D {
        textPosition?: Point;
        glyphPosition?: Point;
    }

    export interface LegendDataPoint extends SelectableDataPoint, LegendPosition2D {
        label: string;
        color: string;
        icon: LegendIcon;
        category?: string;
        measure?: any;
        iconOnlyOnLabel?: boolean;
        tooltip?: string;
    }

    export interface LegendData {
        title?: string;
        dataPoints: LegendDataPoint[];
        grouped?: boolean;
    }

    export var legendProps = {
        show: 'show',
        position: 'position',
        titleText: 'titleText',
        showTitle: 'showTitle',
    };

    export function createLegend(legendParentElement: JQuery,
        interactive: boolean,
        interactivityService: IInteractivityService,
        isScrollable: boolean = false,
        legendPosition: LegendPosition = LegendPosition.Top): ILegend {
        if (interactive) return new CartesianChartInteractiveLegend(legendParentElement);
        else return new SVGLegend(legendParentElement, legendPosition, interactivityService, isScrollable);
    }

    export interface ILegend {
        getMargins(): IViewport;

        isVisible(): boolean;
        changeOrientation(orientation: LegendPosition): void;
        getOrientation(): LegendPosition;
        drawLegend(data: LegendData, viewport: IViewport);
        /**
         * Reset the legend by clearing it
         */
        reset(): void;
    }

    export function getIconClass(iconType: LegendIcon): string {
        switch (iconType) {
            case LegendIcon.Circle:
                return 'icon circle';
            case LegendIcon.Box:
                return 'icon tall';
            case LegendIcon.Line:
                return 'icon short';
            default:
                debug.assertFail('Invalid Chart type: ' + iconType);
        }
    }

    export function getLabelMaxSize(currentViewport: IViewport, numItems: number, hasTitle: boolean): string {
        let viewportWidth = currentViewport.width;
        let smallTileWidth = 250;
        let mediumTileWidth = 490;
        let largeTileWidth = 750;
        let tileMargins = 20;
        let legendMarkerWidth = 28;
        let legendItems;

        if (numItems < 1)
            return '48px';

        if (viewportWidth <= smallTileWidth) {
            legendItems = hasTitle ? 4 : 3;
            //Max width based on minimum number of items design of 3 i.e. '48px' or 4 (with title) - '29px' max-width at least   
            return Math.floor((smallTileWidth - tileMargins - (legendMarkerWidth * legendItems)) / Math.min(numItems, legendItems)) + 'px';
        }

        if (viewportWidth <= mediumTileWidth) {
            legendItems = hasTitle ? 6 : 5;
            //Max width based on minimum number of items design of 5 i.e. '66px' or 6 (with title) - '50px' max-width at least  
            return Math.floor((mediumTileWidth - tileMargins - (legendMarkerWidth * legendItems)) / Math.min(numItems, legendItems)) + 'px';
        }

        if (viewportWidth <= largeTileWidth) {
            legendItems = hasTitle ? 8 : 7;
            //Max width based on minimum number of items design of 7 i.e. '76px' or 8 (with title) - '63px' max-width at least
            return Math.floor((largeTileWidth - tileMargins - (legendMarkerWidth * legendItems)) / Math.min(numItems, legendItems)) + 'px';
        }

        //Wide viewport
        legendItems = hasTitle ? 10 : 9;
        return Math.floor((viewportWidth - tileMargins - (legendMarkerWidth * legendItems)) / Math.min(numItems, legendItems)) + 'px';
    }

    interface TitleLayout {
        x: number;
        y: number;
        text: string;
        width: number;
    }

    enum NavigationArrowType {
        Increase,
        Decrease
    }

    interface NavigationArrow {
        x: number;
        y: number;
        path: string;
        rotateTransform: string;
        type: NavigationArrowType;
    }

    interface LegendLayout {
        numberOfItems: number;
        title: TitleLayout;
        navigationArrows: NavigationArrow[];
    }

    class SVGLegend implements ILegend {
        private orientation: LegendPosition;
        private viewport: IViewport;
        private parentViewport: IViewport;
        private svg: D3.Selection;
        private group: D3.Selection;
        private clearCatcher: D3.Selection;
        private element: JQuery;
        private interactivityService: IInteractivityService;
        private legendDataStartIndex = 0;
        private arrowPosWindow = 1;
        private data: LegendData;
        private isScrollable: boolean;

        private lastCalculatedWidth = 0;
        private visibleLegendWidth = 0;
        private visibleLegendHeight = 0;

        private static LegendIconRadius = 5;
        private static MaxTextLength = 60;
        private static MaxTitleLength = 80;
        private static TextAndIconPadding = 5;
        private static TitlePadding = 15;
        private static LegendEdgeMariginWidth = 10;
        private static LegendMaxWidthFactor = 0.3;
        private static TopLegendHeight = 24;
        
        // Navigation Arrow constants
        private static LegendArrowOffset = 10;
        private static LegendArrowHeight = 15;
        private static LegendArrowWidth = 7.5;
        private static LegendArrowTranslateY = 3.5;

        private static LegendTextProperties: TextProperties = {
            fontFamily: 'wf_segoe-ui_normal',
            fontSize: '11px'
        };

        private static LegendTitleTextProperties: TextProperties = {
            fontFamily: 'wf_segoe-ui_Semibold',
            fontSize: '11px'
        };

        private static LegendItem: ClassAndSelector = {
            class: 'legendItem',
            selector: '.legendItem'
        };

        private static LegendText: ClassAndSelector = {
            class: 'legendText',
            selector: '.legendText'
        };

        private static LegendIcon: ClassAndSelector = {
            class: 'legendIcon',
            selector: '.legendIcon'
        };

        private static LegendTitle: ClassAndSelector = {
            class: 'legendTitle',
            selector: '.legendTitle'
        };

        private static NavigationArrow: ClassAndSelector = {
            class: 'navArrow',
            selector: '.navArrow'
        };

        constructor(
            element: JQuery,
            legendPosition: LegendPosition,
            interactivityService: IInteractivityService,
            isScrollable: boolean) {

            this.svg = d3.select(element.get(0)).insert('svg', ':first-child');
            this.svg.style('display', 'inherit');
            this.svg.classed('legend', true);
            if (interactivityService)
                this.clearCatcher = appendClearCatcher(this.svg);
            this.group = this.svg.append('g').attr('id', 'legendGroup');
            this.interactivityService = interactivityService;
            this.isScrollable = isScrollable;
            this.element = element;
            this.changeOrientation(legendPosition);
            this.parentViewport = { height: 0, width: 0 };
            this.calculateViewport();
            this.updateLayout();
        }

        private updateLayout() {
            let viewport = this.viewport;
            let orientation = this.orientation;
            this.svg
                .attr({
                    'height': viewport.height || (orientation === LegendPosition.None ? 0 : this.parentViewport.height),
                    'width': viewport.width || (orientation === LegendPosition.None ? 0 : this.parentViewport.width)
                })
            /*
             * Workaround for web-kit browsers, since isn't doesn't invalidate dom with correct attr size.
             * This happens intermittently in dashboard, and corrects itself on dom manupilation.
             */
                .style('max-width', '100%');

            let isBottom = orientation === LegendPosition.Bottom || orientation === LegendPosition.BottomCenter;            
            this.svg.style({
                'float': this.getFloat(),
                'position': isBottom ? 'absolute' : '',
                'bottom': isBottom ? '0px' : '',
            });
        }

        private calculateViewport(): void {
            switch (this.orientation) {
                case LegendPosition.Top:
                case LegendPosition.Bottom:
                case LegendPosition.TopCenter:
                case LegendPosition.BottomCenter:
                    this.viewport = { height: SVGLegend.TopLegendHeight, width: 0 };
                    return;
                case LegendPosition.Right:
                case LegendPosition.Left:
                case LegendPosition.RightCenter:
                case LegendPosition.LeftCenter:
                    let width = this.lastCalculatedWidth ? this.lastCalculatedWidth : this.parentViewport.width * SVGLegend.LegendMaxWidthFactor;
                    this.viewport = { height: 0, width: width };
                    return;

                case LegendPosition.None:
                    this.viewport = { height: 0, width: 0 };
            }
        }

        private getFloat(): string {
            switch (this.orientation) {
                case LegendPosition.Right:
                case LegendPosition.RightCenter:
                    return 'right';
                case LegendPosition.Left:
                case LegendPosition.LeftCenter:
                    return 'left';
                default: return '';
            }
        }

        public getMargins(): IViewport {
            return this.viewport;
        }

        public isVisible(): boolean {
            return this.orientation !== LegendPosition.None;
        }

        public changeOrientation(orientation: LegendPosition): void {
            if (orientation) {
                this.orientation = orientation;
            } else {
                this.orientation = LegendPosition.Top;
            }
            this.svg.attr('orientation', orientation);
        }

        public getOrientation(): LegendPosition {
            return this.orientation;
        }

        public drawLegend(data: LegendData, viewport: IViewport): void {
            this.drawLegendInternal(data, viewport, true /* perform auto width */);
        }

        public drawLegendInternal(data: LegendData, viewport: IViewport, autoWidth: boolean): void {
            this.parentViewport = viewport;
            this.data = data;

            if (data.dataPoints.length === 0) {
                this.changeOrientation(LegendPosition.None);
            }

            if (this.getOrientation() === LegendPosition.None) {
                data.dataPoints = [];
            }

            // Adding back the workaround for Legend Left/Right position for Map
            let mapControl = this.element.children(".mapControl");
            if (mapControl.length > 0 && !this.isTopOrBottom(this.orientation)) {
                mapControl.css("display", "inline-block");
            }

            this.calculateViewport();

            let layout = this.calculateLayout(data, autoWidth);
            let titleLayout = layout.title;
            let titleData = titleLayout ? [titleLayout] : [];
            let hasSelection = this.interactivityService && this.interactivityService.legendHasSelection();

            let group = this.group;

            //transform the wrapping group if position is centered           
            if (this.isCentered(this.orientation)) {
                let centerOffset = 0;
                if (this.isTopOrBottom(this.orientation)) {
                    centerOffset = Math.max(0, (this.parentViewport.width - this.visibleLegendWidth) / 2);                  
                    group.attr('transform', SVGUtil.translate(centerOffset,0));
                }
                else {
                    centerOffset = Math.max((this.parentViewport.height - this.visibleLegendHeight) / 2);
                    group.attr('transform', SVGUtil.translate(0,centerOffset));
                }
            }
            else {
                group.attr('transform', null);
            }

            let legendTitle = group
                .selectAll(SVGLegend.LegendTitle.selector)
                .data(titleData);

            legendTitle.enter()
                .append('text')
                .style({
                    'font-size': SVGLegend.LegendTitleTextProperties.fontSize,
                    'font-family': SVGLegend.LegendTitleTextProperties.fontFamily
                })
                .classed(SVGLegend.LegendTitle.class, true);

            legendTitle
                .text((d: TitleLayout) => d.text)
                .attr({
                    'x': (d: TitleLayout) => d.x,
                    'y': (d: TitleLayout) => d.y
                });

            legendTitle.exit().remove();
            
            let virtualizedDataPoints = data.dataPoints.slice(this.legendDataStartIndex, this.legendDataStartIndex + layout.numberOfItems);

            let legendItems = group
                .selectAll(SVGLegend.LegendItem.selector)
                .data(virtualizedDataPoints, (d: LegendDataPoint) => d.label + d.color);

            let itemsEnter = legendItems.enter()
                .append('g')
                .classed(SVGLegend.LegendItem.class, true);

            itemsEnter
                .append('circle')
                .classed(SVGLegend.LegendIcon.class, true);

            itemsEnter.append('title');

            itemsEnter
                .append('text')
                .classed(SVGLegend.LegendText.class, true)
                .style({
                    'font-size': SVGLegend.LegendTextProperties.fontSize,
                    'font-family': SVGLegend.LegendTextProperties.fontFamily
                });

            legendItems
                .select(SVGLegend.LegendIcon.selector)
                .attr({
                    'cx': (d: LegendDataPoint, i) => d.glyphPosition.x,
                    'cy': (d: LegendDataPoint) => d.glyphPosition.y,
                    'r': SVGLegend.LegendIconRadius,
                })
                .style({
                    'fill': (d: LegendDataPoint) => {
                        if (hasSelection && !d.selected)
                            return LegendBehavior.dimmedLegendColor;
                        else
                            return d.color;
                    }
                });

            legendItems
                .select('title')
                .text((d: LegendDataPoint) => d.tooltip);

            legendItems
                .select(SVGLegend.LegendText.selector)
                .attr({
                    'x': (d: LegendDataPoint) => d.textPosition.x,
                    'y': (d: LegendDataPoint) => d.textPosition.y,
                })
                .text((d: LegendDataPoint) => d.label);

            if (this.interactivityService) {
                let iconsSelection = legendItems.select(SVGLegend.LegendIcon.selector);
                let behaviorOptions: LegendBehaviorOptions = {
                    legendItems: legendItems,
                    legendIcons: iconsSelection,
                    clearCatcher: this.clearCatcher,
                };

                this.interactivityService.bind(data.dataPoints, new LegendBehavior(), behaviorOptions, { isLegend: true });
            }

            legendItems.exit().remove();

            this.drawNavigationArrows(layout.navigationArrows);

            this.updateLayout();
        }

        private normalizePosition(points: any[]): void {
            if (this.legendDataStartIndex >= points.length) {
                this.legendDataStartIndex = points.length - 1;
            }

            if (this.legendDataStartIndex < 0) {
                this.legendDataStartIndex = 0;
            }
        }

        private calculateTitleLayout(title: string): TitleLayout {
            let width = 0;
            let hasTitle = !jsCommon.StringExtensions.isNullOrEmpty(title);

            if (hasTitle) {
                let properties = SVGLegend.LegendTextProperties;
                let text = properties.text = title;
                let isHorizontal = this.isTopOrBottom(this.orientation);
                let fixedHorizontalIconShift = SVGLegend.TextAndIconPadding + SVGLegend.LegendIconRadius;
                let fixedHorizontalTextShift = SVGLegend.LegendIconRadius + SVGLegend.TextAndIconPadding + fixedHorizontalIconShift;
                let maxHorizontalSpaceAvaliable = this.parentViewport.width * SVGLegend.LegendMaxWidthFactor
                    - fixedHorizontalTextShift - SVGLegend.LegendEdgeMariginWidth;

                let maxMeasureLength = isHorizontal ? SVGLegend.MaxTitleLength : maxHorizontalSpaceAvaliable;

                width = TextMeasurementService.measureSvgTextWidth(properties);

                if (width > maxMeasureLength) {
                    text = TextMeasurementService.getTailoredTextOrDefault(properties, maxMeasureLength);
                    width = maxMeasureLength;
                };

                if (isHorizontal)
                    width += SVGLegend.TitlePadding;

                return {
                    x: 0,
                    y: 0,
                    text: text,
                    width: width
                };
            }
            return null;

        }
        /** Performs layout offline for optimal perfomance */
        private calculateLayout(data: LegendData, autoWidth: boolean): LegendLayout {
            let dataPoints = data.dataPoints;
            if (data.dataPoints.length === 0) {
                return {
                    startIndex: null,
                    numberOfItems: 0,
                    title: null,
                    navigationArrows: []
                };
            }

            this.normalizePosition(dataPoints);
            if (this.legendDataStartIndex < dataPoints.length) {
                dataPoints = dataPoints.slice(this.legendDataStartIndex);
            }

            let title = this.calculateTitleLayout(data.title);

            let navArrows: NavigationArrow[];
            let numberOfItems: number;
            if (this.isTopOrBottom(this.orientation)) {
                navArrows = this.isScrollable ? this.calculateHorizontalNavigationArrowsLayout(title) : [];
                numberOfItems = this.calculateHorizontalLayout(dataPoints, title, navArrows);
            }
            else {
                navArrows = this.isScrollable ? this.calculateVerticalNavigationArrowsLayout(title) : [];
                numberOfItems = this.calculateVerticalLayout(dataPoints, title, navArrows, autoWidth);
            }
            return {
                numberOfItems: numberOfItems,
                title: title,
                navigationArrows: navArrows
            };
        }

        private updateNavigationArrowLayout(navigationArrows: NavigationArrow[], remainingDataLength, visibleDataLength) {
            if (this.legendDataStartIndex === 0) {
                navigationArrows.shift();
            }

            let lastWindow = this.arrowPosWindow;
            this.arrowPosWindow = visibleDataLength;

            if (navigationArrows && navigationArrows.length > 0 && this.arrowPosWindow === remainingDataLength) {
                this.arrowPosWindow = lastWindow;
                navigationArrows.length = navigationArrows.length - 1;
            }
        }

        private calculateHorizontalNavigationArrowsLayout(title: TitleLayout): NavigationArrow[] {
            let height = SVGLegend.LegendArrowHeight;
            let width = SVGLegend.LegendArrowWidth;
            let translateY = SVGLegend.LegendArrowTranslateY;

            let data: NavigationArrow[] = [];
            let rightShift = title ? title.x + title.width : 0;
            let arrowLeft = SVGUtil.createArrow(width, height, 180 /*angle*/);
            let arrowRight = SVGUtil.createArrow(width, height, 0 /*angle*/);

            data.push({
                x: rightShift,
                y: translateY,
                path: arrowLeft.path,
                rotateTransform: arrowLeft.transform,
                type: NavigationArrowType.Decrease
            });

            data.push({
                x: this.parentViewport.width - width,
                y: translateY,
                path: arrowRight.path,
                rotateTransform: arrowRight.transform,
                type: NavigationArrowType.Increase
            });

            return data;
        }

        private calculateVerticalNavigationArrowsLayout(title: TitleLayout): NavigationArrow[] {
            let height = SVGLegend.LegendArrowHeight;
            let width = SVGLegend.LegendArrowWidth;

            let data: NavigationArrow[] = [];
            let rightShift = 40;
            let arrowTop = SVGUtil.createArrow(width, height, 270 /*angle*/);
            let arrowBottom = SVGUtil.createArrow(width, height, 90 /*angle*/);

            data.push({
                x: rightShift,
                y: height + SVGLegend.LegendArrowOffset / 2,
                path: arrowTop.path,
                rotateTransform: arrowTop.transform,
                type: NavigationArrowType.Decrease
            });

            data.push({
                x: rightShift,
                y: this.parentViewport.height - height,
                path: arrowBottom.path,
                rotateTransform: arrowBottom.transform,
                type: NavigationArrowType.Increase
            });

            return data;
        }

        private calculateHorizontalLayout(dataPoints: LegendDataPoint[], title: TitleLayout, navigationArrows: NavigationArrow[]): number {
            debug.assertValue(navigationArrows, 'navigationArrows');

            let fixedTextShift = SVGLegend.LegendIconRadius + SVGLegend.TextAndIconPadding;
            let fixedIconShift = 11;
            fixedTextShift = fixedIconShift + 4;
            let totalSpaceOccupiedThusFar = 0;
            let iconTotalItemPadding = SVGLegend.LegendIconRadius * 2 + SVGLegend.TextAndIconPadding * 3;
            let numberOfItems: number = dataPoints.length;

            if (title) {
                totalSpaceOccupiedThusFar = title.width;
                title.y = fixedTextShift;
            }

            if (this.legendDataStartIndex > 0) {
                totalSpaceOccupiedThusFar += SVGLegend.LegendArrowOffset;
            }

            // This bit expands the max lengh if there are only a few items
            // so longer labels can potentially get more space, and not be
            // ellipsed. 
            let dataPointsLength = dataPoints.length;
            let parentWidth = this.parentViewport.width;
            let maxTextLength = dataPointsLength > 0
                ? (((parentWidth - totalSpaceOccupiedThusFar) - (iconTotalItemPadding * dataPointsLength)) / dataPointsLength) | 0
                : 0;
            maxTextLength = maxTextLength > SVGLegend.MaxTextLength ? maxTextLength : SVGLegend.MaxTextLength;

            for (let i = 0; i < dataPointsLength; i++) {
                let dp = dataPoints[i];

                dp.glyphPosition = {
                    x: totalSpaceOccupiedThusFar + SVGLegend.LegendIconRadius,
                    y: fixedIconShift
                };

                dp.textPosition = {
                    x: totalSpaceOccupiedThusFar + fixedTextShift,
                    y: fixedTextShift
                };

                let properties = SVGLegend.LegendTextProperties;
                properties.text = dp.label;
                dp.tooltip = dp.label;

                let width = TextMeasurementService.measureSvgTextWidth(properties);
                let spaceTakenByItem = 0;
                if (width < maxTextLength) {
                    spaceTakenByItem = iconTotalItemPadding + width;
                } else {
                    let text = TextMeasurementService.getTailoredTextOrDefault(
                        properties,
                        maxTextLength);
                    dp.label = text;
                    spaceTakenByItem = iconTotalItemPadding + maxTextLength;
                }

                totalSpaceOccupiedThusFar += spaceTakenByItem;

                if (totalSpaceOccupiedThusFar > parentWidth) {
                    numberOfItems = i;
                    break;
                }
            }

            this.visibleLegendWidth = totalSpaceOccupiedThusFar;

            this.updateNavigationArrowLayout(navigationArrows, dataPointsLength, numberOfItems);

            return numberOfItems;
        }

        private calculateVerticalLayout(
            dataPoints: LegendDataPoint[],
            title: TitleLayout,
            navigationArrows: NavigationArrow[],
            autoWidth: boolean): number {
            let verticalLegendHeight = 20;
            let spaceNeededByTitle = 15;
            let totalSpaceOccupiedThusFar = verticalLegendHeight;
            let extraShiftForTextAlignmentToIcon = 4;
            let fixedHorizontalIconShift = SVGLegend.TextAndIconPadding + SVGLegend.LegendIconRadius;
            let fixedHorizontalTextShift = SVGLegend.LegendIconRadius + SVGLegend.TextAndIconPadding + fixedHorizontalIconShift;
            let maxHorizontalSpaceAvaliable = autoWidth
                ? this.parentViewport.width * SVGLegend.LegendMaxWidthFactor
                - fixedHorizontalTextShift - SVGLegend.LegendEdgeMariginWidth
                : this.lastCalculatedWidth
                - fixedHorizontalTextShift - SVGLegend.LegendEdgeMariginWidth;
            let numberOfItems: number = dataPoints.length;

            let maxHorizontalSpaceUsed = 0;
            let parentHeight = this.parentViewport.height;

            if (title) {
                totalSpaceOccupiedThusFar += spaceNeededByTitle;
                title.x = SVGLegend.TextAndIconPadding;
                title.y = spaceNeededByTitle;
                maxHorizontalSpaceUsed = title.width || 0;
            }

            if (this.legendDataStartIndex > 0)
                totalSpaceOccupiedThusFar += SVGLegend.LegendArrowOffset;

            let dataPointsLength = dataPoints.length;
            for (let i = 0; i < dataPointsLength; i++) {
                let dp = dataPoints[i];

                dp.glyphPosition = {
                    x: fixedHorizontalIconShift,
                    y: totalSpaceOccupiedThusFar
                };

                dp.textPosition = {
                    x: fixedHorizontalTextShift,
                    y: totalSpaceOccupiedThusFar + extraShiftForTextAlignmentToIcon
                };

                let properties = SVGLegend.LegendTextProperties;
                properties.text = dp.label;
                dp.tooltip = dp.label;

                // TODO: [PERF] Get rid of this extra measurement, and modify
                // getTailoredTextToReturnWidth + Text
                let width = TextMeasurementService.measureSvgTextWidth(properties);
                if (width > maxHorizontalSpaceUsed) {
                    maxHorizontalSpaceUsed = width;
                }

                if (width > maxHorizontalSpaceAvaliable) {
                    let text = TextMeasurementService.getTailoredTextOrDefault(
                        properties,
                        maxHorizontalSpaceAvaliable);
                    dp.label = text;
                }

                totalSpaceOccupiedThusFar += verticalLegendHeight;

                if (totalSpaceOccupiedThusFar > parentHeight) {
                    numberOfItems = i;
                    break;
                }
            }

            if (autoWidth) {
                if (maxHorizontalSpaceUsed < maxHorizontalSpaceAvaliable) {
                    this.lastCalculatedWidth = this.viewport.width = Math.ceil(maxHorizontalSpaceUsed + fixedHorizontalTextShift + SVGLegend.LegendEdgeMariginWidth);
                } else {
                    this.lastCalculatedWidth = this.viewport.width = Math.ceil(this.parentViewport.width * SVGLegend.LegendMaxWidthFactor);
                }
            }
            else {
                this.viewport.width = this.lastCalculatedWidth;
            }

            this.visibleLegendHeight = totalSpaceOccupiedThusFar;

            navigationArrows.forEach(d => d.x = this.lastCalculatedWidth / 2);
            this.updateNavigationArrowLayout(navigationArrows, dataPointsLength, numberOfItems);

            return numberOfItems;
        }

        private drawNavigationArrows(layout: NavigationArrow[]) {
            let arrows = this.group.selectAll(SVGLegend.NavigationArrow.selector)
                .data(layout);

            arrows
                .enter()
                .append('g')
                .on('click', (d: NavigationArrow) => {
                    let pos = this.legendDataStartIndex;
                    this.legendDataStartIndex = d.type === NavigationArrowType.Increase
                        ? pos + this.arrowPosWindow : pos - this.arrowPosWindow;
                    this.drawLegendInternal(this.data, this.parentViewport, false);
                })
                .classed(SVGLegend.NavigationArrow.class, true)
                .append('path');

            arrows
                .attr('transform', (d: NavigationArrow) => SVGUtil.translate(d.x, d.y))
                .select('path')
                .attr({
                    'd': (d: NavigationArrow) => d.path,
                    'transform': (d: NavigationArrow) => d.rotateTransform
                });

            arrows.exit().remove();
        }

        private isTopOrBottom(orientation: LegendPosition) {
            switch (orientation) {
                case LegendPosition.Top:
                case LegendPosition.Bottom:
                case LegendPosition.BottomCenter:     
                case LegendPosition.TopCenter:                                   
                    return true;
                default:
                    return false;
            }
        }

        private isCentered(orientation: LegendPosition): boolean {
            switch (orientation) {
                case LegendPosition.BottomCenter:
                case LegendPosition.LeftCenter:
                case LegendPosition.RightCenter:
                case LegendPosition.TopCenter:
                    return true;
                default:
                    return false;
            }
        }

        public reset(): void {
            // Intentionally left blank. 
        }
    }

    class CartesianChartInteractiveLegend implements ILegend {
        private static LegendHeight = 65;
        private static LegendContainerClass = 'interactive-legend';
        private static LegendTitleClass = 'title';
        private static LegendItem = 'item';
        private static legendPlaceSelector = '\u25A0';
        private static legendIconClass = 'icon';
        private static legendColorCss = 'color';
        private static legendItemNameClass = 'itemName';
        private static legendItemMeasureClass = 'itemMeasure';
        private element: JQuery;
        private legendContainerDiv: D3.Selection;

        constructor(element: JQuery) {
            this.element = element;
        }

        public static getIconClass(chartType: LegendIcon): string {
            switch (chartType) {
                case LegendIcon.Circle:
                case LegendIcon.Box:
                case LegendIcon.Line:
                    return 'icon';
                default:
                    debug.assertFail('Invalid Chart type: ' + chartType);
            }
        }

        public getMargins(): IViewport {
            return {
                height: CartesianChartInteractiveLegend.LegendHeight,
                width: 0
            };
        }

        public drawLegend(legendData: LegendData) {
            debug.assertValue(legendData, 'legendData');
            let data = legendData.dataPoints;
            debug.assertValue(data, 'dataPoints');
            if (data.length < 1) return;
            let legendContainerDiv = this.legendContainerDiv;
            if (!legendContainerDiv) {
                if (!data.length) return;
                let divToPrepend = $('<div></div>')
                    .height(this.getMargins().height)
                    .addClass(CartesianChartInteractiveLegend.LegendContainerClass);
                // Prepending, as legend should always be on topmost visual.
                this.element.prepend(divToPrepend);
                this.legendContainerDiv = legendContainerDiv = d3.select(divToPrepend.get(0));
            }

            // Construct the legend title and items.
            this.drawTitle(data);
            this.drawLegendItems(data);
        }

        public reset(): void {
            if (this.legendContainerDiv) {
                this.legendContainerDiv.remove();
                this.legendContainerDiv = null;
            }
        }

        public isVisible(): boolean {
            return true;
        }

        public changeOrientation(orientation: LegendPosition) {
            // Not supported
        }

        public getOrientation(): LegendPosition {
            return LegendPosition.Top;
        }

        /**
         * Draw the legend title
         */
        private drawTitle(data: LegendDataPoint[]): void {
            debug.assert(data && data.length > 0, 'data is null or empty');
            let titleDiv: D3.Selection = this.legendContainerDiv.selectAll('div.' + CartesianChartInteractiveLegend.LegendTitleClass);
            let item: D3.UpdateSelection = titleDiv.data([data[0]]);

            // Enter
            let itemEnter: D3.EnterSelection = item.enter();
            let titleDivEnter: D3.Selection = itemEnter.append('div').attr('class', CartesianChartInteractiveLegend.LegendTitleClass);
            titleDivEnter
                .filter((d: LegendDataPoint) => d.iconOnlyOnLabel)
                .append('span')
                .attr('class', CartesianChartInteractiveLegend.legendIconClass)
                .html(CartesianChartInteractiveLegend.legendPlaceSelector);
            titleDivEnter.append('span');

            // Update
            item.filter((d: LegendDataPoint) => d.iconOnlyOnLabel)
                .select('span.' + CartesianChartInteractiveLegend.legendIconClass)
                .style(CartesianChartInteractiveLegend.legendColorCss, (d: LegendDataPoint) => d.color);
            item.select('span:last-child').text((d: LegendDataPoint) => d.category);
        }

        /**
         * Draw the legend items
         */
        private drawLegendItems(data: LegendDataPoint[]): void {
            // Add Mesaures - the items of the category in the legend
            this.ensureLegendTableCreated();
            let dataPointsMatrix: LegendDataPoint[][] = CartesianChartInteractiveLegend.splitArrayToOddEven(data);
            let legendItemsContainer: D3.UpdateSelection = this.legendContainerDiv.select('tbody').selectAll('tr').data(dataPointsMatrix);

            // trs is table rows. 
            // there are two table rows.
            // the order of insertion to the legend table is:
            // Even data points got inserted into the 1st line
            // Odd data points got inserted into the 2nd line
            // ----------------------------
            // | value0 | value 2 | value 4
            // ----------------------------
            // | value1 | value 3 | 
            // ----------------------------
            // 

            // Enter
            let legendItemsEnter: D3.EnterSelection = legendItemsContainer.enter();
            let rowEnter: D3.Selection = legendItemsEnter.append('tr');
            let cellEnter: D3.Selection = rowEnter.selectAll('td')
                .data((d: LegendDataPoint[]) => d, (d: LegendDataPoint) => d.label)
                .enter()
                .append('td').attr('class', CartesianChartInteractiveLegend.LegendItem);
            let cellSpanEnter: D3.Selection = cellEnter.append('span');
            cellSpanEnter.filter((d: LegendDataPoint) => !d.iconOnlyOnLabel)
                .append('span')
                .html(CartesianChartInteractiveLegend.legendPlaceSelector)
                .attr('class', CartesianChartInteractiveLegend.legendIconClass)
                .attr('white-space', 'nowrap');
            cellSpanEnter.append('span').attr('class', CartesianChartInteractiveLegend.legendItemNameClass);
            cellSpanEnter.append('span').attr('class', CartesianChartInteractiveLegend.legendItemMeasureClass);

            // Update
            let legendCells: D3.UpdateSelection = legendItemsContainer.selectAll('td').data((d: LegendDataPoint[]) => d, (d: LegendDataPoint) => d.label);
            legendCells.select('span.' + CartesianChartInteractiveLegend.legendItemNameClass).html((d: LegendDataPoint) => powerbi.visuals.TextUtil.removeBreakingSpaces(d.label));
            legendCells.select('span.' + CartesianChartInteractiveLegend.legendItemMeasureClass).html((d: LegendDataPoint) => '&nbsp;' + d.measure);
            legendCells.select('span.' + CartesianChartInteractiveLegend.legendIconClass).style('color', (d: LegendDataPoint) => d.color);

            // Exit
            legendCells.exit().remove();
        }

        /**
         * Ensure legend table is created and set horizontal pan gestures on it
         */
        private ensureLegendTableCreated(): void {
            if (this.legendContainerDiv.select('div table').empty()) {
                let legendTable: D3.Selection = this.legendContainerDiv.append('div').append('table');
                legendTable.style('table-layout', 'fixed').append('tbody');
                // Setup Pan Gestures of the legend
                this.setPanGestureOnLegend(legendTable);
            }
        }

        /**
         * Set Horizontal Pan gesture for the legend
         */
        private setPanGestureOnLegend(legendTable: D3.Selection): void {
            let viewportWidth: number = $(this.legendContainerDiv.select('div:nth-child(2)')[0]).width();
            let xscale: D3.Scale.LinearScale = d3.scale.linear().domain([0, viewportWidth]).range([0, viewportWidth]);
            let zoom: D3.Behavior.Zoom = d3.behavior.zoom()
                .scaleExtent([1, 1]) // disable scaling
                .x(xscale)
                .on("zoom", () => {
                    // horizontal pan is valid only in case the legend items width are bigger than the viewport width
                    if ($(legendTable[0]).width() > viewportWidth) {
                        let t: number[] = zoom.translate();
                        let tx: number = t[0];
                        let ty: number = t[1];
                        tx = Math.min(tx, 0);
                        tx = Math.max(tx, viewportWidth - $(legendTable[0]).width());
                        zoom.translate([tx, ty]);
                    legendTable.style("-ms-transform",() => { /* IE 9 */
                        return SVGUtil.translateXWithPixels(tx);
                    });
                    legendTable.style("-webkit-transform",() => { /* Safari */
                        return SVGUtil.translateXWithPixels(tx);
                    });
                        legendTable.style("transform", () => {
                            return SVGUtil.translateXWithPixels(tx);
                        });
                    }
                });
            if (this.legendContainerDiv) {
                this.legendContainerDiv.call(zoom);
            } else {
                legendTable.call(zoom);
            }
        }

        /**
         * Split legend data points array into odd and even arrays
         * Even array will be the legend first line and Odd array will be the 2nd legend line 
         */
        private static splitArrayToOddEven(data: LegendDataPoint[]): LegendDataPoint[][] {
            let oddData: LegendDataPoint[] = [];
            let evenData: LegendDataPoint[] = [];
            for (let i = 0; i < data.length; ++i) {
                if (i % 2 === 0) {
                    evenData.push(data[i]);
                }
                else {
                    oddData.push(data[i]);
                }
            }
            return [evenData, oddData];
        }
    }

    export module LegendData {
        export function update(legendData: LegendData, legendObject: DataViewObject): void {
            debug.assertValue(legendData, 'legendData');
            debug.assertValue(legendObject, 'legendObject');

            if (legendObject[legendProps.show] == null) {
                legendObject[legendProps.show] = true;
            }

            if (legendObject[legendProps.show] === false)
                legendData.dataPoints = [];

            if (legendObject[legendProps.show] === true && legendObject[legendProps.position] == null) {
                legendObject[legendProps.position] = legendPosition.top;
            }

            if (legendObject[legendProps.showTitle] === false)
                legendData.title = "";
            else if (legendObject[legendProps.titleText] !== undefined) {
                legendData.title = <string>legendObject[legendProps.titleText];
            }
        }
    }
}