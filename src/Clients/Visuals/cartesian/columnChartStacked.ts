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
    export class StackedColumnChartStrategy implements IColumnChartStrategy {
        private static classes = {
            item: {
                class: 'column',
                selector: '.column'
            },
            highlightItem: {
                class: 'highlightColumn',
                selector: '.highlightColumn'
            },
        };

        private data: ColumnChartData;
        private graphicsContext: ColumnChartContext;
        private width: number;
        private height: number;
        private margin: IMargin;
        private xProps: IAxisProperties;
        private yProps: IAxisProperties;
        private categoryLayout: CategoryLayout;
        private columnsCenters: number[];
        private columnSelectionLineHandle: D3.Selection;
        private animator: IColumnChartAnimator;
        private interactivityService: IInteractivityService;
        private viewportHeight: number;
        private viewportWidth: number;
        private static validLabelPositions: RectLabelPosition[] = [RectLabelPosition.InsideCenter, RectLabelPosition.InsideEnd, RectLabelPosition.InsideBase];
        private layout: IColumnLayout;

        public setupVisualProps(columnChartProps: ColumnChartContext): void {
            this.graphicsContext = columnChartProps;
            this.margin = columnChartProps.margin;
            this.width = this.graphicsContext.width;
            this.height = this.graphicsContext.height;
            this.categoryLayout = columnChartProps.layout;
            this.animator = columnChartProps.animator;
            this.interactivityService = columnChartProps.interactivityService;
            this.viewportHeight = columnChartProps.viewportHeight;
            this.viewportWidth = columnChartProps.viewportWidth;
        }

        public setData(data: ColumnChartData) {
            this.data = data;
        }

        public setXScale(is100Pct: boolean, forcedTickCount?: number, forcedXDomain?: any[], axisScaleType?: string): IAxisProperties {
            let width = this.width;

            let forcedXMin, forcedXMax;

            if (forcedXDomain && forcedXDomain.length === 2) {
                forcedXMin = forcedXDomain[0];
                forcedXMax = forcedXDomain[1];
            }

            let props = this.xProps = ColumnUtil.getCategoryAxis(
                this.data,
                width,
                this.categoryLayout,
                false,
                forcedXMin,
                forcedXMax,
                axisScaleType);

            return props;
        }

        public setYScale(is100Pct: boolean, forcedTickCount?: number, forcedYDomain?: any[], axisScaleType?: string): IAxisProperties {
            let height = this.viewportHeight;
            let yProps = this.yProps = StackedUtil.getValueAxis(
                this.data,
                is100Pct,
                height,
                [height, 0],
                forcedTickCount,
                forcedYDomain,
                axisScaleType);

            return yProps;
        }

        public drawColumns(useAnimation: boolean): ColumnChartDrawInfo {
            let data = this.data;
            debug.assertValue(data, 'data should not be null or undefined');

            this.columnsCenters = null; // invalidate the columnsCenters so that will be calculated again

            let axisOptions: ColumnAxisOptions = {
                columnWidth: this.categoryLayout.categoryThickness * (1 - CartesianChart.InnerPaddingRatio),
                xScale: this.xProps.scale,
                yScale: this.yProps.scale,
                isScalar: this.categoryLayout.isScalar,
                margin: this.margin,
            };
            let stackedColumnLayout = this.layout = StackedColumnChartStrategy.getLayout(data, axisOptions);
            let dataLabelSettings = data.labelSettings;
            let labelDataPoints: LabelDataPoint[] = [];
            if (dataLabelSettings && dataLabelSettings.show) {
                labelDataPoints = this.createLabelDataPoints();
            }

            let result: ColumnChartAnimationResult;
            let shapes: D3.UpdateSelection;
            let series = ColumnUtil.drawSeries(data, this.graphicsContext.mainGraphicsContext, axisOptions);
            if (this.animator && useAnimation) {
                result = this.animator.animate({
                    viewModel: data,
                    series: series,
                    layout: stackedColumnLayout,
                    itemCS: StackedColumnChartStrategy.classes.item,
                    interactivityService: this.interactivityService,
                    mainGraphicsContext: this.graphicsContext.mainGraphicsContext,
                    viewPort: { height: this.height, width: this.width },
                });
                shapes = result.shapes;
            }
            if (!this.animator || !useAnimation || result.failed) {
                shapes = ColumnUtil.drawDefaultShapes(data, series, stackedColumnLayout, StackedColumnChartStrategy.classes.item, !this.animator, this.interactivityService && this.interactivityService.hasSelection());
            }

            ColumnUtil.applyInteractivity(shapes, this.graphicsContext.onDragStart);

            return {
                shapesSelection: shapes,
                viewport: { height: this.height, width: this.width },
                axisOptions,
                labelDataPoints: labelDataPoints,
            };
        }

        public selectColumn(selectedColumnIndex: number, lastSelectedColumnIndex: number): void {
            ColumnUtil.setChosenColumnOpacity(this.graphicsContext.mainGraphicsContext, StackedColumnChartStrategy.classes.item.selector, selectedColumnIndex, lastSelectedColumnIndex);
            this.moveHandle(selectedColumnIndex);
        }

        public getClosestColumnIndex(x: number, y: number): number {
            return ColumnUtil.getClosestColumnIndex(x, this.getColumnsCenters());
        }

        /**
         * Get the chart's columns centers (x value).
         */
        private getColumnsCenters(): number[] {
            if (!this.columnsCenters) { // lazy creation
                let categoryWidth: number = this.categoryLayout.categoryThickness * (1 - CartesianChart.InnerPaddingRatio);
                // use the axis scale and first series data to get category centers
                if (this.data.series.length > 0) {
                    let xScaleOffset = 0;
                    if (!this.categoryLayout.isScalar)
                        xScaleOffset = categoryWidth / 2;
                    let firstSeries = this.data.series[0];
                    this.columnsCenters = firstSeries.data.map(d => this.xProps.scale(this.categoryLayout.isScalar ? d.categoryValue : d.categoryIndex) + xScaleOffset);
                }
            }
            return this.columnsCenters;
        }

        private moveHandle(selectedColumnIndex: number) {
            let columnCenters = this.getColumnsCenters();
            let x = columnCenters[selectedColumnIndex];

            if (!this.columnSelectionLineHandle) {
                let handle = this.columnSelectionLineHandle = this.graphicsContext.mainGraphicsContext.append('g');
                handle.append('line')
                    .classed('interactive-hover-line', true)
                    .attr({
                        x1: x,
                        x2: x,
                        y1: 0,
                        y2: this.height,
                    });

                handle.append('circle')
                    .attr({
                        cx: x,
                        cy: this.height,
                        r: '6px',
                    })
                    .classed('drag-handle', true);
            }
            else {
                let handle = this.columnSelectionLineHandle;
                handle.select('line').attr({ x1: x, x2: x });
                handle.select('circle').attr({ cx: x });
            }
        }

        public static getLayout(data: ColumnChartData, axisOptions: ColumnAxisOptions): IColumnLayout {
            let columnWidth = axisOptions.columnWidth;
            let isScalar = axisOptions.isScalar;
            let xScale = axisOptions.xScale;
            let yScale = axisOptions.yScale;
            let scaledY0 = yScale(0);
            let xScaleOffset = 0;

            if (isScalar)
                xScaleOffset = columnWidth / 2;

            return {
                shapeLayout: {
                    width: (d: ColumnChartDataPoint) => columnWidth,
                    x: (d: ColumnChartDataPoint) => xScale(isScalar ? d.categoryValue : d.categoryIndex) - xScaleOffset,
                    y: (d: ColumnChartDataPoint) => scaledY0 + AxisHelper.diffScaled(yScale, d.position, 0),
                    height: (d: ColumnChartDataPoint) => StackedUtil.getSize(yScale, d.valueAbsolute)
                },
                shapeLayoutWithoutHighlights: {
                    width: (d: ColumnChartDataPoint) => columnWidth,
                    x: (d: ColumnChartDataPoint) => xScale(isScalar ? d.categoryValue : d.categoryIndex) - xScaleOffset,
                    y: (d: ColumnChartDataPoint) => scaledY0 + AxisHelper.diffScaled(yScale, d.originalPosition, 0),
                    height: (d: ColumnChartDataPoint) => StackedUtil.getSize(yScale, d.originalValueAbsolute)
                },
                zeroShapeLayout: {
                    width: (d: ColumnChartDataPoint) => columnWidth,
                    x: (d: ColumnChartDataPoint) => xScale(isScalar ? d.categoryValue : d.categoryIndex) - xScaleOffset,
                    y: (d: ColumnChartDataPoint) => scaledY0 + AxisHelper.diffScaled(yScale, d.position, 0) + StackedUtil.getSize(yScale, d.valueAbsolute),
                    height: (d: ColumnChartDataPoint) => 0
                },
            };
        }

        private createLabelDataPoints(): LabelDataPoint[]{
            let labelDataPoints: LabelDataPoint[] = [];
            let data = this.data;
            let series = data.series;
            let formattersCache = NewDataLabelUtils.createColumnFormatterCacheManager();
            let shapeLayout = this.layout.shapeLayout;

            for (let currentSeries of series) {
                let labelSettings = currentSeries.labelSettings ? currentSeries.labelSettings : data.labelSettings;
                let axisFormatter: number = NewDataLabelUtils.getDisplayUnitValueFromAxisFormatter(this.yProps.formatter, labelSettings);
                for (let dataPoint of currentSeries.data) {
                    if ((this.interactivityService && this.interactivityService.hasSelection() && !dataPoint.selected) || (data.hasHighlights && !dataPoint.highlight)) {
                        continue;
                    }

                    // Calculate parent rectangle
                    let parentRect: IRect = {
                        left: shapeLayout.x(dataPoint),
                        top: shapeLayout.y(dataPoint),
                        width: shapeLayout.width(dataPoint),
                        height: shapeLayout.height(dataPoint),
                    };

                    // Calculate label text
                    let formatString = "";
                    if (this.graphicsContext.is100Pct) {
                        formatString = NewDataLabelUtils.hundredPercentFormat;
                    }
                    else {
                        formatString = dataPoint.labelFormatString;
                    }
                    let formatter = formattersCache.getOrCreate(formatString, labelSettings, axisFormatter);
                    let text = NewDataLabelUtils.getLabelFormattedText(formatter.format(dataPoint.value));

                    // Calculate text size
                    let properties: TextProperties = {
                        text: text,
                        fontFamily: NewDataLabelUtils.LabelTextProperties.fontFamily,
                        fontSize: NewDataLabelUtils.LabelTextProperties.fontSize,
                        fontWeight: NewDataLabelUtils.LabelTextProperties.fontWeight,
                    };
                    let textWidth = TextMeasurementService.measureSvgTextWidth(properties);
                    let textHeight = TextMeasurementService.estimateSvgTextHeight(properties);

                    labelDataPoints.push({
                        isPreferred: true,
                        text: text,
                        textSize: {
                            width: textWidth,
                            height: textHeight,
                        },
                        outsideFill: labelSettings.labelColor ? labelSettings.labelColor : NewDataLabelUtils.defaultLabelColor,
                        insideFill: labelSettings.labelColor ? labelSettings.labelColor : NewDataLabelUtils.defaultInsideLabelColor,
                        isParentRect: true,
                        parentShape: {
                            rect: parentRect,
                            orientation: dataPoint.value >= 0 ? NewRectOrientation.VerticalBottomBased : NewRectOrientation.VerticalTopBased,
                            validPositions: StackedColumnChartStrategy.validLabelPositions,
                        },
                        identity: dataPoint.identity,
                    });
                }
            }
            
            return labelDataPoints;
        }
    }

    export class StackedBarChartStrategy implements IColumnChartStrategy {
        private static classes = {
            item: {
                class: 'bar',
                selector: '.bar'
            },
            highlightItem: {
                class: 'highlightBar',
                selector: '.highlightBar'
            },
        };

        private data: ColumnChartData;
        private graphicsContext: ColumnChartContext;
        private width: number; height: number;
        private margin: IMargin;
        private xProps: IAxisProperties;
        private yProps: IAxisProperties;
        private categoryLayout: CategoryLayout;
        private barsCenters: number[];
        private columnSelectionLineHandle: D3.Selection;
        private animator: IColumnChartAnimator;
        private interactivityService: IInteractivityService;
        private viewportHeight: number;
        private viewportWidth: number;
        private static validLabelPositions: RectLabelPosition[] = [RectLabelPosition.InsideCenter, RectLabelPosition.InsideEnd, RectLabelPosition.InsideBase];
        private layout: IColumnLayout;

        public setupVisualProps(barChartProps: ColumnChartContext): void {
            this.graphicsContext = barChartProps;
            this.margin = barChartProps.margin;
            this.width = this.graphicsContext.width;
            this.height = this.graphicsContext.height;
            this.categoryLayout = barChartProps.layout;
            this.animator = barChartProps.animator;
            this.interactivityService = barChartProps.interactivityService;
            this.viewportHeight = barChartProps.viewportHeight;
            this.viewportWidth = barChartProps.viewportWidth;
        }

        public setData(data: ColumnChartData) {
            this.data = data;
        }

        public setYScale(is100Pct: boolean, forcedTickCount?: number, forcedYDomain?: any[], axisScaleType?: string): IAxisProperties {
            let height = this.height;

            let forcedYMin, forcedYMax;

            if (forcedYDomain && forcedYDomain.length === 2) {
                forcedYMin = forcedYDomain[0];
                forcedYMax = forcedYDomain[1];
            }

            let props = this.yProps = ColumnUtil.getCategoryAxis(
                this.data,
                height,
                this.categoryLayout,
                true,
                forcedYMin,
                forcedYMax,
                axisScaleType);

            return props;
        }

        public setXScale(is100Pct: boolean, forcedTickCount?: number, forcedXDomain?: any[], axisScaleType?: string): IAxisProperties {
            debug.assert(forcedTickCount === undefined, 'Cannot have stacked bar chart as combo chart.');

            let height = this.viewportHeight;

            let xProps = this.xProps = StackedUtil.getValueAxis(
                this.data,
                is100Pct,
                this.width,
                [0, this.width],
                undefined,
                forcedXDomain,
                axisScaleType);

            xProps.axis.tickSize(-height, 0);

            return xProps;
        }

        public drawColumns(useAnimation: boolean): ColumnChartDrawInfo {
            let data = this.data;
            debug.assertValue(data, 'data should not be null or undefined');

            this.barsCenters = null; // invalidate the barsCenters so that will be calculated again

            let axisOptions: ColumnAxisOptions = {
                columnWidth: this.categoryLayout.categoryThickness * (1 - CartesianChart.InnerPaddingRatio),
                xScale: this.xProps.scale,
                yScale: this.yProps.scale,
                isScalar: this.categoryLayout.isScalar,
                margin: this.margin,
            };
            let stackedBarLayout = this.layout = StackedBarChartStrategy.getLayout(data, axisOptions);
            let dataLabelSettings = data.labelSettings;
            let labelDataPoints: LabelDataPoint[] = [];
            if (dataLabelSettings && dataLabelSettings.show) {
                labelDataPoints = this.createLabelDataPoints();
            }

            let result: ColumnChartAnimationResult;
            let shapes: D3.UpdateSelection;
            let series = ColumnUtil.drawSeries(data, this.graphicsContext.mainGraphicsContext, axisOptions);
            if (this.animator && useAnimation) {
                result = this.animator.animate({
                    viewModel: data,
                    series: series,
                    layout: stackedBarLayout,
                    itemCS: StackedBarChartStrategy.classes.item,
                    interactivityService: this.interactivityService,
                    mainGraphicsContext: this.graphicsContext.mainGraphicsContext,
                    viewPort: { height: this.height, width: this.width },
                });
                shapes = result.shapes;
            }
            if (!this.animator || !useAnimation || result.failed) {
                shapes = ColumnUtil.drawDefaultShapes(data, series, stackedBarLayout, StackedBarChartStrategy.classes.item, !this.animator, this.interactivityService && this.interactivityService.hasSelection());
            }

            ColumnUtil.applyInteractivity(shapes, this.graphicsContext.onDragStart);

            return {
                shapesSelection: shapes,
                viewport: { height: this.height, width: this.width },
                axisOptions: axisOptions,
                labelDataPoints: labelDataPoints,
            };
        }

        public selectColumn(selectedColumnIndex: number, lastInteractiveSelectedColumnIndex: number): void {
            ColumnUtil.setChosenColumnOpacity(this.graphicsContext.mainGraphicsContext, StackedBarChartStrategy.classes.item.selector, selectedColumnIndex, lastInteractiveSelectedColumnIndex);
            this.moveHandle(selectedColumnIndex);
        }

        public getClosestColumnIndex(x: number, y: number): number {
            return ColumnUtil.getClosestColumnIndex(y, this.getBarsCenters());
        }

        /** 
         * Get the chart's columns centers (y value).
         */
        private getBarsCenters(): number[] {
            if (!this.barsCenters) { // lazy creation
                let barWidth: number = this.categoryLayout.categoryThickness * (1 - CartesianChart.InnerPaddingRatio);
                // use the axis scale and first series data to get category centers
                if (this.data.series.length > 0) {
                    let yScaleOffset = 0;
                    if (!this.categoryLayout.isScalar)
                        yScaleOffset = barWidth / 2;
                    let firstSeries = this.data.series[0];
                    this.barsCenters = firstSeries.data.map(d => this.yProps.scale(this.categoryLayout.isScalar ? d.categoryValue : d.categoryIndex) + yScaleOffset);
                }
            }
            return this.barsCenters;
        }

        private moveHandle(selectedColumnIndex: number) {
            let barCenters = this.getBarsCenters();
            let y = barCenters[selectedColumnIndex];

            if (!this.columnSelectionLineHandle) {
                let handle = this.columnSelectionLineHandle = this.graphicsContext.mainGraphicsContext.append('g');
                handle.append('line')
                    .classed('interactive-hover-line', true)
                    .attr({
                        x1: 0,
                        x2: this.width,
                        y1: y,
                        y2: y
                    });
                handle.append('circle')
                    .classed('drag-handle', true)
                    .attr({
                        cx: 0,
                        cy: y,
                        r: '6px',
                    });

            }
            else {
                let handle = this.columnSelectionLineHandle;
                handle.select('line').attr({ y1: y, y2: y });
                handle.select('circle').attr({ cy: y });
            }
        }

        public static getLayout(data: ColumnChartData, axisOptions: ColumnAxisOptions): IColumnLayout {
            let columnWidth = axisOptions.columnWidth;
            let isScalar = axisOptions.isScalar;
            let xScale = axisOptions.xScale;
            let yScale = axisOptions.yScale;
            let scaledX0 = xScale(0);
            let xScaleOffset = 0;

            if (isScalar)
                xScaleOffset = columnWidth / 2;

            return {
                shapeLayout: {
                    width: (d: ColumnChartDataPoint) => -StackedUtil.getSize(xScale, d.valueAbsolute),
                    x: (d: ColumnChartDataPoint) => scaledX0 + AxisHelper.diffScaled(xScale, d.position - d.valueAbsolute, 0),
                    y: (d: ColumnChartDataPoint) => yScale(isScalar ? d.categoryValue : d.categoryIndex) - xScaleOffset,
                    height: (d: ColumnChartDataPoint) => columnWidth,
                },
                shapeLayoutWithoutHighlights: {
                    width: (d: ColumnChartDataPoint) => -StackedUtil.getSize(xScale, d.originalValueAbsolute),
                    x: (d: ColumnChartDataPoint) => scaledX0 + AxisHelper.diffScaled(xScale, d.originalPosition - d.originalValueAbsolute, 0),
                    y: (d: ColumnChartDataPoint) => yScale(isScalar ? d.categoryValue : d.categoryIndex) - xScaleOffset,
                    height: (d: ColumnChartDataPoint) => columnWidth,
                },
                zeroShapeLayout: {
                    width: (d: ColumnChartDataPoint) => 0,
                    x: (d: ColumnChartDataPoint) => scaledX0 + AxisHelper.diffScaled(xScale, d.position - d.valueAbsolute, 0),
                    y: (d: ColumnChartDataPoint) => yScale(isScalar ? d.categoryValue : d.categoryIndex) - xScaleOffset,
                    height: (d: ColumnChartDataPoint) => columnWidth,
                },
            };
        }

        private createLabelDataPoints(): LabelDataPoint[] {
            let labelDataPoints: LabelDataPoint[] = [];
            let data = this.data;
            let series = data.series;
            let formattersCache = NewDataLabelUtils.createColumnFormatterCacheManager();
            let shapeLayout = this.layout.shapeLayout;

            for (let currentSeries of series) {
                let labelSettings = currentSeries.labelSettings ? currentSeries.labelSettings : data.labelSettings;
                let axisFormatter: number = NewDataLabelUtils.getDisplayUnitValueFromAxisFormatter(this.yProps.formatter, labelSettings);
                for (let dataPoint of currentSeries.data) {
                    if ((this.interactivityService && this.interactivityService.hasSelection() && !dataPoint.selected) || (data.hasHighlights && !dataPoint.highlight)) {
                        continue;
                    }

                    // Calculate label text
                    let formatString = undefined;
                    if (this.graphicsContext.is100Pct) {
                        formatString = NewDataLabelUtils.hundredPercentFormat;
                    }
                    else {
                        formatString = dataPoint.labelFormatString;
                    }
                    let formatter = formattersCache.getOrCreate(formatString, labelSettings, axisFormatter);
                    let text = NewDataLabelUtils.getLabelFormattedText(formatter.format(dataPoint.value));

                    // Calculate text size
                    let properties: TextProperties = {
                        text: text,
                        fontFamily: NewDataLabelUtils.LabelTextProperties.fontFamily,
                        fontSize: NewDataLabelUtils.LabelTextProperties.fontSize,
                        fontWeight: NewDataLabelUtils.LabelTextProperties.fontWeight,
                    };
                    let textWidth = TextMeasurementService.measureSvgTextWidth(properties);
                    let textHeight = TextMeasurementService.estimateSvgTextHeight(properties);

                    // Calculate parent rectangle
                    let parentRect: IRect = {
                        left: shapeLayout.x(dataPoint),
                        top: shapeLayout.y(dataPoint),
                        width: shapeLayout.width(dataPoint),
                        height: shapeLayout.height(dataPoint),
                    };

                    labelDataPoints.push({
                        isPreferred: true,
                        text: text,
                        textSize: {
                            width: textWidth,
                            height: textHeight,
                        },
                        outsideFill: labelSettings.labelColor ? labelSettings.labelColor : NewDataLabelUtils.defaultLabelColor,
                        insideFill: labelSettings.labelColor ? labelSettings.labelColor : NewDataLabelUtils.defaultInsideLabelColor,
                        isParentRect: true,
                        parentShape: {
                            rect: parentRect,
                            orientation: dataPoint.value >= 0 ? NewRectOrientation.HorizontalLeftBased : NewRectOrientation.HorizontalRightBased,
                            validPositions: StackedBarChartStrategy.validLabelPositions,
                        },
                        identity: dataPoint.identity,
                    });
                }
            }

            return labelDataPoints;
        }
    }
} 