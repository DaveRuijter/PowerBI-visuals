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

/**
 * IMPORTANT: This chart is not currently enabled in the PBI system and is under development.
 */
module powerbi.visuals {

    export interface IDataDotChartConfiguration {
        xAxisProperties: IAxisProperties;
        yAxisProperties: IAxisProperties;
        margin: any;
    }

    export interface DataDotChartData {
        series: DataDotChartSeries;
        hasHighlights: boolean;
        hasDynamicSeries: boolean;
    }

    export interface DataDotChartSeries extends CartesianSeries {
        xCol: DataViewMetadataColumn;
        yCol: DataViewMetadataColumn;
        data: DataDotChartDataPoint[];   //overridden type of array     
    }

    export interface DataDotChartDataPoint extends CartesianDataPoint, SelectableDataPoint {
        highlight: boolean;
    }

    export interface DataDotChartConstructorOptions extends CartesianVisualConstructorOptions {
    }

    /**
     * The data dot chart shows a set of circles with the data value inside them.
     * The circles are regularly spaced similar to column charts.
     * The radius of all dots is the same across the chart.
     * This is most often combined with a column chart to create the 'chicken pox' chart.
     * If any of the data values do not fit within the circles, then the data values are hidden
     * and the y axis for the dots is displayed instead.
     * This chart only supports a single series of data.
     * This chart does not display a legend.
     */
    export class DataDotChart implements ICartesianVisual {
        public static formatStringProp: DataViewObjectPropertyIdentifier = { objectName: 'general', propertyName: 'formatString' };
        private static ClassName = 'dataDotChart';

        private static DotClassName = 'dot';
        private static DotClassSelector = '.dot';
        private static DotColorKey = 'dataDot';

        private static DotLabelClassName = 'label';
        private static DotLabelClassSelector = '.label';
        private static DotLabelVerticalOffset = '0.4em';
        private static DotLabelTextAnchor = 'middle';

        private options: CartesianVisualInitOptions;

        // Chart properties
        private svg: D3.Selection;
        private element: JQuery;
        private mainGraphicsG: D3.Selection;
        private mainGraphicsContext: D3.Selection;
        private currentViewport: IViewport;
        private hostService: IVisualHostServices;
        private cartesianVisualHost: ICartesianVisualHost;
        private style: IVisualStyle;
        private colors: IDataColorPalette;
        private isScrollable: boolean;

        // Cartesian chart properties
        private xAxisProperties: IAxisProperties;
        private yAxisProperties: IAxisProperties;
        private margin;

        // Data properties
        private data: DataDotChartData;
        private dataViewCategorical: DataViewCategorical;
        private clippedData: DataDotChartData;

        // Interactivity properties
        private interactivityService: IInteractivityService;
        private interactivity: InteractivityOptions;

        constructor(options: DataDotChartConstructorOptions) {
            this.isScrollable = options.isScrollable;
            this.interactivityService = options.interactivityService;
        }

        public init(options: CartesianVisualInitOptions): void {
            this.options = options;

            // Common properties
            this.svg = options.svg;
            this.mainGraphicsG = this.svg.append('g')
                .classed('dataDotChartMainGraphicsContext', true);
            this.mainGraphicsContext = this.mainGraphicsG.append('svg');
            this.currentViewport = options.viewport;
            this.hostService = options.host;
            this.cartesianVisualHost = options.cartesianHost;
            this.style = options.style;
            this.colors = this.style.colorPalette.dataColors;

            // Interactivity properties
            this.interactivity = options.interactivity;

            let element = this.element = options.element;
            element.addClass(DataDotChart.ClassName);
            element.css('overflow', 'visible');
        }

        public setData(dataViews: DataView[]): void {
            this.data = {
                series: <DataDotChartSeries>{
                    data: <DataDotChartDataPoint[]>[]
                },
                hasHighlights: false,
                hasDynamicSeries: false,
            };

            if (dataViews.length > 0) {

                // I only handle a single data view
                let dataView = dataViews[0];
                if (dataView && dataView.categorical) {

                    let dataViewCategorical = this.dataViewCategorical = dataView.categorical;
                    let dvCategories = dataViewCategorical.categories;

                    // I default to text unless there is a category type
                    let categoryType = ValueType.fromDescriptor({ text: true });
                    if (dvCategories && dvCategories.length > 0 && dvCategories[0].source && dvCategories[0].source.type)
                        categoryType = dvCategories[0].source.type;

                    this.data = DataDotChart.converter(dataView, valueFormatter.format(null));
                    if (this.interactivityService) {
                        this.interactivityService.applySelectionStateToData(this.data.series.data);
                    }
                }
            }
        }

        public setFilteredData(startIndex: number, endIndex: number): any {
            let data = this.clippedData = Prototype.inherit(this.data);

            if (data && data.series && data.series.data)
                data.series = { data: data.series.data.slice(startIndex, endIndex), xCol: data.series.xCol, yCol: data.series.yCol };

            return data;
        }

        public calculateAxesProperties(options: CalculateScaleAndDomainOptions): IAxisProperties[] {
            this.currentViewport = options.viewport;
            this.margin = options.margin;

            let data = this.clippedData = this.data;
            let viewport = this.currentViewport;
            let margin = this.margin;
            let series: DataDotChartSeries = data ? data.series : null;
            let seriesArray = series && series.data && series.data.length > 0 ? [series] : [];
            let categoryCount = series && series.data ? series.data.length : 0;

            // If there are highlights, then the series is 2x in length and highlights are interwoven.
            if (data.hasHighlights) {
                categoryCount = categoryCount / 2;
            }

            let width = viewport.width - (margin.left + margin.right);
            let height = viewport.height - (margin.top + margin.bottom);

            let xMetaDataColumn: DataViewMetadataColumn;
            let yMetaDataColumn: DataViewMetadataColumn;

            if (DataDotChart.hasDataPoint(series)) {
                xMetaDataColumn = series.xCol;
                yMetaDataColumn = series.yCol;
            }

            let layout = CartesianChart.getLayout(
                null,
                {
                    availableWidth: width,
                    categoryCount: categoryCount,
                    domain: null,
                    isScalar: false,
                    isScrollable: this.isScrollable
                });
            let outerPadding = layout.categoryThickness * CartesianChart.OuterPaddingRatio;

            // clip data that won't fit
            if (!this.isScrollable) {
                this.clippedData = DataDotChart.createClippedDataIfOverflowed(data, layout.categoryCount);
            }

            let yDomain = AxisHelper.createValueDomain(seriesArray, /*includeZero:*/ true) || fallBackDomain;

            let combinedDomain = AxisHelper.combineDomain(options.forcedYDomain, yDomain);

            this.yAxisProperties = AxisHelper.createAxis({
                pixelSpan: height,
                dataDomain: combinedDomain,
                metaDataColumn: yMetaDataColumn,
                formatStringProp: DataDotChart.formatStringProp,
                outerPadding: 0,
                isScalar: true,
                isVertical: true,
                forcedTickCount: options.forcedTickCount,
                useTickIntervalForDisplayUnits: true,
                isCategoryAxis: true
            });

            let axisType = this.xAxisProperties ? this.xAxisProperties.axisType : ValueType.fromDescriptor({ text: true });
            let xDomain = AxisHelper.createDomain(seriesArray, axisType, /*isScalar:*/ false, options.forcedXDomain);
            this.xAxisProperties = AxisHelper.createAxis({
                pixelSpan: width,
                dataDomain: xDomain,
                metaDataColumn: xMetaDataColumn,
                formatStringProp: DataDotChart.formatStringProp,
                outerPadding: outerPadding,
                isScalar: false,
                isVertical: false,
                forcedTickCount: options.forcedTickCount,
                useTickIntervalForDisplayUnits: true,
                categoryThickness: layout.categoryThickness,
                getValueFn: (index, type) => this.lookupXValue(index, type),
                isCategoryAxis: false
            });

            return [this.xAxisProperties, this.yAxisProperties];
        }

        private static createClippedDataIfOverflowed(data: DataDotChartData, categoryCount: number): DataDotChartData {                                                

            // If there are highlights, then the series is 2x in length and highlights are interwoven.
            let requiredLength = data.hasHighlights ? Math.min(data.series.data.length, categoryCount * 2) : Math.min(data.series.data.length, categoryCount);

            if (requiredLength >= data.series.data.length) {
                return data;
            }

            let clipped: DataDotChartData = Prototype.inherit(data);
            clipped.series = Prototype.inherit(data.series); // This prevents clipped and data from sharing the series object
            clipped.series.data = clipped.series.data.slice(0, requiredLength);
            return clipped;
        }

        private static hasDataPoint(series: DataDotChartSeries): boolean {
            return (series && series.data && series.data.length > 0);
        }

        private lookupXValue(index: number, type: ValueType): any {
            let data = this.data;

            let isDateTime = AxisHelper.isDateTime(type);
            if (isDateTime)
                return new Date(index);

            if (data && data.series) {
                let seriesData = data.series.data;

                if (seriesData) {
                    let dataAtIndex = seriesData[index];
                    if (dataAtIndex) {
                        return dataAtIndex.categoryValue;
                    }
                }
            }

            return index;
        }

        public overrideXScale(xProperties: IAxisProperties): void {
            this.xAxisProperties = xProperties;
        }

        public render(suppressAnimations: boolean): CartesianVisualRenderResult {
            if (!this.clippedData)
                return;
            let data = this.clippedData;
            let dataPoints = data.series.data;
            let hasHighlights = data.hasHighlights;

            let margin = this.margin;
            let viewport = this.currentViewport;
            let width = viewport.width - (margin.left + margin.right);
            let height = viewport.height - (margin.top + margin.bottom);
            let xScale = <D3.Scale.OrdinalScale>this.xAxisProperties.scale;
            let yScale = this.yAxisProperties.scale;
            let dotWidth = this.xAxisProperties.categoryThickness * (1 - CartesianChart.InnerPaddingRatio);
            let dotRadius = dotWidth / 2;
            let dotColor = this.cartesianVisualHost.getSharedColors().getNewColorScale().getColor(DataDotChart.DotColorKey);

            let hasSelection = this.interactivityService ? this.interactivityService.hasSelection() : false;

            this.mainGraphicsContext.attr('width', width)
                .attr('height', height);

            let dots = this.mainGraphicsContext.selectAll(DataDotChart.DotClassSelector).data(dataPoints, d => d.identity.getKey());

            dots.enter()
                .append('circle')
                .classed(DataDotChart.DotClassName, true);

            dots
                .style({ 'fill': dotColor.value })
                .style('fill-opacity', (d: DataDotChartDataPoint) => ColumnUtil.getFillOpacity(d.selected, d.highlight, hasSelection, hasHighlights))
                .classed('null-value', (d: DataDotChartDataPoint) => d.value === null)
                .attr({
                    r: (d: DataDotChartDataPoint) => dotRadius,
                    cx: d => xScale(d.categoryIndex) + dotRadius,
                    cy: d => yScale(d.value)
                });

            dots.exit().remove();

            let dotLabels = this.mainGraphicsContext.selectAll(DataDotChart.DotLabelClassSelector).data(dataPoints, d => d.identity.getKey());

            dotLabels.enter()
                .append('text')
                .classed(DataDotChart.DotLabelClassName, true)
                .attr({
                    'text-anchor': DataDotChart.DotLabelTextAnchor,
                    dy: DataDotChart.DotLabelVerticalOffset
                });

            dotLabels
                .classed('null-value', (d: DataDotChartDataPoint) => d.value === null)
                .classed('overflowed', false)
                .attr({
                    x: d => xScale(d.categoryIndex) + dotRadius,
                    y: d => yScale(d.value)
                })
                .text(d => this.yAxisProperties.formatter.format(d.value));

            let overflowed = false;
            dotLabels
                .each(function () {
                    // jQuery fails to properly inspect SVG class elements, the $('<div>') notation works around it.
                    if (!overflowed && !$("<div>").addClass($(this).attr("class")).hasClass("null-value")) {
                        let width = TextMeasurementService.measureSvgTextElementWidth(this);
                        if (width > dotWidth) {
                            dotLabels.classed('overflowed', true);
                            overflowed = true;
                        }
                    }
                });

            dotLabels.exit().remove();
            let behaviorOptions: DataDotChartBehaviorOptions = undefined;
            if (this.interactivityService) {
                behaviorOptions = {
                    dots: dots,
                    dotLabels: dotLabels,
                    datapoints: dataPoints,
                };
            }        

            // This should always be the last line in the render code.
            SVGUtil.flushAllD3TransitionsIfNeeded(this.options);

            return { dataPoints: dataPoints, behaviorOptions: behaviorOptions, labelDataPoints: [] };
        }

        public calculateLegend(): LegendData {
            return this.createLegendDataPoints(0); // start with index 0
        }

        public hasLegend(): boolean {
            return this.data && this.data.hasDynamicSeries;
        }

        private createLegendDataPoints(columnIndex: number): LegendData {
            let data = this.data;
            if (!data)
                return null;

            let series = data.series;
            let seriesData = series.data;

            let legendDataPoints: LegendDataPoint[] = [];
            let category: any;

            let axisType = this.xAxisProperties ? this.xAxisProperties.axisType : ValueType.fromDescriptor({ text: true });

            // Category will be the same for all series. This is an optimization.
            if (data.series && data.series.data) {
                let firstDataPoint: DataDotChartDataPoint = data.series.data[0];
                category = firstDataPoint && this.lookupXValue(firstDataPoint.categoryValue, axisType);
            }

            // Create a legend data point for the specified column                
            if (series.yCol) {

                let formatStringProp = DataDotChart.formatStringProp;
                let lineDataPoint = seriesData[columnIndex];
                let measure = lineDataPoint && lineDataPoint.value;

                let label = converterHelper.getFormattedLegendLabel(series.yCol, this.dataViewCategorical.values, formatStringProp);

                let dotColor = this.cartesianVisualHost.getSharedColors().getNewColorScale().getColor(DataDotChart.DotColorKey);
                let dataViewCategoricalValues = this.dataViewCategorical.values;
                let identity = dataViewCategoricalValues && dataViewCategoricalValues.length > columnIndex ?
                    SelectionId.createWithIdAndMeasure(dataViewCategoricalValues[columnIndex].identity, dataViewCategoricalValues[columnIndex].source.queryName) :
                    SelectionId.createWithMeasure(dataViewCategoricalValues.source.queryName);
                legendDataPoints.push({
                    color: dotColor.value,
                    icon: LegendIcon.Line,
                    label: label,
                    category: valueFormatter.format(category, valueFormatter.getFormatString(series.xCol, formatStringProp)),
                    measure: valueFormatter.format(measure, valueFormatter.getFormatString(series.yCol, formatStringProp)),
                    identity: identity,
                    selected: false
                });
            }

            return { dataPoints: legendDataPoints };
        }

        public onClearSelection(): void {
            if (this.interactivityService)
                this.interactivityService.clearSelection();

            // cartesianChart handles calling render again.
        }

        public static converter(dataView: DataView, blankCategoryValue: string): DataDotChartData {
            let categorical = dataView.categorical;

            let category: DataViewCategoryColumn = categorical.categories && categorical.categories.length > 0
                ? categorical.categories[0]
                : {
                    source: undefined,
                    values: [blankCategoryValue],
                    identity: undefined
                };

            let categoryType: ValueType = AxisHelper.getCategoryValueType(category.source);
            let isDateTime = AxisHelper.isDateTime(categoryType);
            let categoryValues = category.values;

            // I only handle a single series
            if (categorical.values) {
                let measure = categorical.values[0];

                let hasHighlights: boolean = !!measure.highlights;

                let dataPoints: DataDotChartDataPoint[] = [];
                for (let categoryIndex = 0, len = measure.values.length; categoryIndex < len; categoryIndex++) {

                    debug.assert(!category.identity || categoryIndex < category.identity.length, 'Category identities is smaller than category values.');

                    // I create the identity from the category.  If there is no category, then I use the measure name to create identity
                    let identity = category.identity ?
                        SelectionId.createWithIdAndMeasure(category.identity[categoryIndex], measure.source.queryName) :
                        SelectionId.createWithMeasure(measure.source.queryName);

                    let categoryValue = categoryValues[categoryIndex];

                    dataPoints.push({
                        categoryValue: isDateTime && categoryValue ? categoryValue.getTime() : categoryValue,
                        value: measure.values[categoryIndex],
                        categoryIndex: categoryIndex,
                        seriesIndex: 0,
                        selected: false,
                        identity: identity,
                        highlight: false
                    });

                    if (hasHighlights) {

                        let highlightIdentity = SelectionId.createWithHighlight(identity);
                        let highlightValue = measure.highlights[categoryIndex];

                        dataPoints.push({
                            categoryValue: isDateTime && categoryValue ? categoryValue.getTime() : categoryValue,
                            value: highlightValue,
                            categoryIndex: categoryIndex,
                            seriesIndex: 0,
                            selected: false,
                            identity: highlightIdentity,
                            highlight: true
                        });
                    }
                }

                return {
                    series: {
                        xCol: category.source,
                        yCol: measure.source,
                        data: dataPoints
                    },
                    hasHighlights: hasHighlights,
                    hasDynamicSeries: true,
                };
            }

            return {
                series: <DataDotChartSeries> {
                    data: <DataDotChartDataPoint[]>[]
                },
                hasHighlights: false,
                hasDynamicSeries: false,
            };
        }
    }
} 