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
    import Color = jsCommon.color;

    export interface ScatterChartConstructorOptions extends CartesianVisualConstructorOptions {
    }

    export interface ScatterChartDataPoint extends SelectableDataPoint, TooltipEnabledDataPoint, LabelEnabledDataPoint {
        x: any;
        y: any;
        size: any;
        radius: RadiusData;
        fill: string;        
        category: string;
    }

    export interface RadiusData {
        sizeMeasure: DataViewValueColumn;
        index: number;
    }

    export interface DataRange {
        minRange: number;
        maxRange: number;
        delta: number;
    }

    export interface ScatterChartData {
        xCol: DataViewMetadataColumn;
        yCol: DataViewMetadataColumn;
        dataPoints: ScatterChartDataPoint[];
        legendData: LegendData;
        axesLabels: ChartAxesLabels;
        size?: DataViewMetadataColumn;
        sizeRange: NumberRange;
        dataLabelsSettings: PointDataLabelsSettings;
        defaultDataPointColor?: string;
        showAllDataPoints?: boolean;
        hasDynamicSeries?: boolean;
        fillPoint?: boolean;
        colorBorder?: boolean;
    }

    interface ScatterChartMeasureMetadata {
        idx: {
            x?: number;
            y?: number;
            size?: number;
        };
        cols: {
            x?: DataViewMetadataColumn;
            y?: DataViewMetadataColumn;
            size?: DataViewMetadataColumn;
        };
        axesLabels: ChartAxesLabels;
    }

    interface MouseCoordinates {
        x: number;
        y: number;
    }

    export class ScatterChart implements ICartesianVisual {
        private static ScatterChartCircleTagName = 'circle';
        private static BubbleRadius = 3 * 2;
        public static DefaultBubbleOpacity = 0.85;
        public static DimmedBubbleOpacity = 0.4;
        public static StrokeDarkenColorValue = 255 * 0.25;
        // Chart Area and size range values as defined by PV charts
        private static AreaOf300By300Chart = 90000;
        private static MinSizeRange = 200;
        private static MaxSizeRange = 3000;
        private static ClassName = 'scatterChart';
        private static MainGraphicsContextClassName = 'mainGraphicsContext';
        private static validLabelPositions = [NewPointLabelPosition.Above, NewPointLabelPosition.Below, NewPointLabelPosition.Left, NewPointLabelPosition.Right];

        private static DotClasses: ClassAndSelector = {
            class: 'dot',
            selector: '.dot'
        };

        private svg: D3.Selection;
        private element: JQuery;
        private mainGraphicsContext: D3.Selection;
        private mainGraphicsG: D3.Selection;
        private labelGraphicsContext: D3.Selection;
        private currentViewport: IViewport;
        private style: IVisualStyle;
        private data: ScatterChartData;
        private dataView: DataView;
        private host: IVisualHostServices;
        private margin: IMargin;
        private xAxisProperties: IAxisProperties;
        private yAxisProperties: IAxisProperties;
        private colors: IDataColorPalette;
        private options: VisualInitOptions;
        private interactivity: InteractivityOptions;
        private cartesianVisualHost: ICartesianVisualHost;
        private isInteractiveChart: boolean;
        private interactivityService: IInteractivityService;
        private categoryAxisProperties: DataViewObject;
        private valueAxisProperties: DataViewObject;
        private animator: IGenericAnimator;

        constructor(options: ScatterChartConstructorOptions) {
            if (options) {
                this.interactivityService = options.interactivityService;
                this.animator = options.animator;
            }
        }

        public init(options: CartesianVisualInitOptions) {
            this.options = options;
            let element = this.element = options.element;
            this.currentViewport = options.viewport;
            this.style = options.style;
            this.host = options.host;
            this.colors = this.style.colorPalette.dataColors;
            this.interactivity = options.interactivity;
            this.cartesianVisualHost = options.cartesianHost;
            this.isInteractiveChart = options.interactivity && options.interactivity.isInteractiveLegend;

            element.addClass(ScatterChart.ClassName);
            let svg = this.svg = options.svg;

            this.mainGraphicsG = svg.append('g')
                .classed(ScatterChart.MainGraphicsContextClassName, true);

            this.mainGraphicsContext = this.mainGraphicsG.append('svg');

            this.labelGraphicsContext = svg.append('g')
                .classed(NewDataLabelUtils.labelGraphicsContextClass.class, true);
        }

        public static converter(dataView: DataView, currentViewport: IViewport, colorPalette: IDataColorPalette, interactivityService?: IInteractivityService, categoryAxisProperties?: DataViewObject, valueAxisProperties?: DataViewObject): ScatterChartData {
            let categoryValues: any[],
                categoryFormatter: IValueFormatter,
                categoryObjects: DataViewObjects[],
                categoryIdentities: DataViewScopeIdentity[],
                categoryQueryName: string;

            let dataViewCategorical: DataViewCategorical = dataView.categorical;
            let dataViewMetadata: DataViewMetadata = dataView.metadata;

            if (dataViewCategorical.categories && dataViewCategorical.categories.length > 0) {
                categoryValues = dataViewCategorical.categories[0].values;
                categoryFormatter = valueFormatter.create({ format: valueFormatter.getFormatString(dataViewCategorical.categories[0].source, scatterChartProps.general.formatString), value: categoryValues[0], value2: categoryValues[categoryValues.length - 1] });
                categoryIdentities = dataViewCategorical.categories[0].identity;
                categoryObjects = dataViewCategorical.categories[0].objects;
                categoryQueryName = dataViewCategorical.categories[0].source.queryName;
            }
            else {
                categoryValues = [null];
                // creating default formatter for null value (to get the right string of empty value from the locale)
                categoryFormatter = valueFormatter.createDefaultFormatter(null); 
            }

            let categories = dataViewCategorical.categories;
            let dataValues = dataViewCategorical.values;
            let hasDynamicSeries = !!dataValues.source;
            let grouped = dataValues.grouped();
            let dvSource = dataValues.source;
            let scatterMetadata = ScatterChart.getMetadata(grouped, dvSource);
            let dataLabelsSettings = dataLabelUtils.getDefaultPointLabelSettings();
            let fillPoint = false;
            let defaultDataPointColor = undefined;
            let showAllDataPoints = undefined;
            let colorBorder = false;
            if (dataViewMetadata && dataViewMetadata.objects) {
                let objects = dataViewMetadata.objects;

                defaultDataPointColor = DataViewObjects.getFillColor(objects, columnChartProps.dataPoint.defaultColor);
                showAllDataPoints = DataViewObjects.getValue<boolean>(objects, columnChartProps.dataPoint.showAllDataPoints);

                let labelsObj = objects['categoryLabels'];
                if (labelsObj) {
                    dataLabelsSettings.show = (labelsObj['show'] !== undefined) ? <boolean>labelsObj['show'] : dataLabelsSettings.show;
                    dataLabelsSettings.precision = (labelsObj['labelsPrecision'] !== undefined) ? +<string>labelsObj['labelsPrecision'] : dataLabelsSettings.precision;
                    if (labelsObj['color'] !== undefined) {
                        dataLabelsSettings.labelColor = (<Fill>labelsObj['color']).solid.color;
                    }
                }

                fillPoint = DataViewObjects.getValue(objects, scatterChartProps.fillPoint.show, fillPoint);
                colorBorder = DataViewObjects.getValue(objects, scatterChartProps.colorBorder.show, colorBorder);
            }

            let dataPoints = ScatterChart.createDataPoints(
                dataValues,
                scatterMetadata,
                categories,
                categoryValues,
                categoryFormatter,
                categoryIdentities,
                categoryObjects,
                colorPalette,
                currentViewport,
                hasDynamicSeries,
                dataLabelsSettings,
                defaultDataPointColor,
                categoryQueryName);

            if (interactivityService) {
                interactivityService.applySelectionStateToData(dataPoints);
            }

            let legendItems = hasDynamicSeries
                ? ScatterChart.createSeriesLegend(dataValues, colorPalette, dataValues, valueFormatter.getFormatString(dvSource, scatterChartProps.general.formatString), defaultDataPointColor)
                : [];

            let legendTitle = dataValues && dvSource ? dvSource.displayName : "";
            if (!legendTitle) {
                legendTitle = categories && categories[0].source.displayName ? categories[0].source.displayName : "";
            }

            let legendData = { title: legendTitle, dataPoints: legendItems };

            let sizeRange = ScatterChart.getSizeRangeForGroups(grouped, scatterMetadata.idx.size);

            if (categoryAxisProperties && categoryAxisProperties["showAxisTitle"] !== null && categoryAxisProperties["showAxisTitle"] === false) {
                scatterMetadata.axesLabels.x = null;
            }
            if (valueAxisProperties && valueAxisProperties["showAxisTitle"] !== null && valueAxisProperties["showAxisTitle"] === false) {
                scatterMetadata.axesLabels.y = null;
            }

            return {
                xCol: scatterMetadata.cols.x,
                yCol: scatterMetadata.cols.y,
                dataPoints: dataPoints,
                legendData: legendData,
                axesLabels: scatterMetadata.axesLabels,
                selectedIds: [],
                size: scatterMetadata.cols.size,
                sizeRange: sizeRange,
                dataLabelsSettings: dataLabelsSettings,
                defaultDataPointColor: defaultDataPointColor,
                hasDynamicSeries: hasDynamicSeries,
                showAllDataPoints: showAllDataPoints,
                fillPoint: fillPoint,
                colorBorder: colorBorder,
            };
        }

        private static getSizeRangeForGroups(
            dataViewValueGroups: DataViewValueColumnGroup[],
            sizeColumnIndex: number): NumberRange {

            let result: NumberRange = {};
            if (dataViewValueGroups) {
                dataViewValueGroups.forEach((group) => {
                    let sizeColumn = ScatterChart.getMeasureValue(sizeColumnIndex, group.values);
                    let currentRange: NumberRange = AxisHelper.getRangeForColumn(sizeColumn);
                    if (result.min == null || result.min > currentRange.min) {
                        result.min = currentRange.min;
                    }
                    if (result.max == null || result.max < currentRange.max) {
                        result.max = currentRange.max;
                    }
                });
            }
            return result;
        }

        private static createDataPoints(
            dataValues: DataViewValueColumns,
            metadata: ScatterChartMeasureMetadata,
            categories: DataViewCategoryColumn[],
            categoryValues: any[],
            categoryFormatter: IValueFormatter,
            categoryIdentities: DataViewScopeIdentity[],
            categoryObjects: DataViewObjects[],
            colorPalette: IDataColorPalette,
            viewport: IViewport,
            hasDynamicSeries: boolean,
            labelSettings: PointDataLabelsSettings,
            defaultDataPointColor?: string,
            categoryQueryName?: string): ScatterChartDataPoint[]{

            let dataPoints: ScatterChartDataPoint[] = [],
                indicies = metadata.idx,
                formatStringProp = scatterChartProps.general.formatString,
                dataValueSource = dataValues.source,
                grouped = dataValues.grouped();

            let colorHelper = new ColorHelper(colorPalette, scatterChartProps.dataPoint.fill, defaultDataPointColor);

            for (let categoryIdx = 0, ilen = categoryValues.length; categoryIdx < ilen; categoryIdx++) {
                let categoryValue = categoryValues[categoryIdx];

                for (let seriesIdx = 0, len = grouped.length; seriesIdx < len; seriesIdx++) {
                    let grouping = grouped[seriesIdx];
                    let seriesValues = grouping.values;
                    let measureX = ScatterChart.getMeasureValue(indicies.x, seriesValues);
                    let measureY = ScatterChart.getMeasureValue(indicies.y, seriesValues);
                    let measureSize = ScatterChart.getMeasureValue(indicies.size, seriesValues);

                    let xVal = measureX && measureX.values ? measureX.values[categoryIdx] : null;
                    let yVal = measureY && measureY.values ? measureY.values[categoryIdx] : 0;
                    let size = measureSize && measureSize.values ? measureSize.values[categoryIdx] : null;

                    let hasNullValue = (xVal == null) || (yVal == null);

                    if (hasNullValue)
                        continue;

                    let color: string;
                    if (hasDynamicSeries) {
                        color = colorHelper.getColorForSeriesValue(grouping.objects, dataValues.identityFields, grouping.name);
                    }
                    else {
                        // If we have no Size measure then use a blank query name
                        let measureSource = (measureSize != null)
                            ? measureSize.source.queryName
                            : '';

                        color = colorHelper.getColorForMeasure(categoryObjects && categoryObjects[categoryIdx], measureSource);
                    }

                    let category = categories && categories.length > 0 ? categories[0] : null;
                    let identity = SelectionIdBuilder.builder()
                        .withCategory(category, categoryIdx)
                        .withSeries(dataValues, grouping)
                        .createSelectionId();

                    let seriesData: TooltipSeriesDataItem[] = [];
                    if (dataValueSource) {
                        // Dynamic series
                        seriesData.push({ value: grouping.name, metadata: { source: dataValueSource, values: [] } });
                    }
                    if (measureX) {
                        seriesData.push({ value: xVal, metadata: measureX });
                    }
                    if (measureY) {
                        seriesData.push({ value: yVal, metadata: measureY });
                    }
                    if (measureSize && measureSize.values && measureSize.values.length > 0) {
                        seriesData.push({ value: measureSize.values[categoryIdx], metadata: measureSize });
                    }

                    let tooltipInfo: TooltipDataItem[] = TooltipBuilder.createTooltipInfo(formatStringProp, null, categoryValue, null, categories, seriesData);

                    let dataPoint: ScatterChartDataPoint = {
                        x: xVal,
                        y: yVal,
                        size: size,
                        radius: { sizeMeasure: measureSize, index: categoryIdx },
                        fill: color,
                        category: categoryFormatter.format(categoryValue),
                        selected: false,
                        identity: identity,
                        tooltipInfo: tooltipInfo,
                        labelFill: labelSettings.labelColor,
                    };

                    dataPoints.push(dataPoint);
                }
            }
            return dataPoints;
        }

        private static createSeriesLegend(
            dataValues: DataViewValueColumns,
            colorPalette: IDataColorPalette,
            categorical: DataViewValueColumns,
            formatString: string,
            defaultDataPointColor: string): LegendDataPoint[] {

            let grouped = dataValues.grouped();
            let colorHelper = new ColorHelper(colorPalette, scatterChartProps.dataPoint.fill, defaultDataPointColor);

            let legendItems: LegendDataPoint[] = [];
            for (let i = 0, len = grouped.length; i < len; i++) {
                let grouping = grouped[i];
                let color = colorHelper.getColorForSeriesValue(grouping.objects, dataValues.identityFields, grouping.name);
                legendItems.push({
                    color: color,
                    icon: LegendIcon.Circle,
                    label: valueFormatter.format(grouping.name, formatString),
                    identity: grouping.identity ? SelectionId.createWithId(grouping.identity) : SelectionId.createNull(),
                    selected: false
                });
            }

            return legendItems;
        }

        public static getBubbleRadius(radiusData: RadiusData, sizeRange: NumberRange, viewPort: IViewport): number {
            let actualSizeDataRange = null;
            let bubblePixelAreaSizeRange = null;
            let measureSize = radiusData.sizeMeasure;

            if (!measureSize)
                return ScatterChart.BubbleRadius;

            let minSize = sizeRange.min ? sizeRange.min : 0;
            let maxSize = sizeRange.max ? sizeRange.max : 0;

            let min = Math.min(minSize, 0);
            let max = Math.max(maxSize, 0);
            actualSizeDataRange = {
                minRange: min,
                maxRange: max,
                delta: max - min
            };

            bubblePixelAreaSizeRange = ScatterChart.getBubblePixelAreaSizeRange(viewPort, ScatterChart.MinSizeRange, ScatterChart.MaxSizeRange);

            if (measureSize.values) {
                let sizeValue = measureSize.values[radiusData.index];
                if (sizeValue != null) {
                    return ScatterChart.projectSizeToPixels(sizeValue, actualSizeDataRange, bubblePixelAreaSizeRange) / 2;
                }
            }

            return ScatterChart.BubbleRadius;
        }

        public static getMeasureValue(measureIndex: number, seriesValues: DataViewValueColumn[]): DataViewValueColumn {
            if (measureIndex >= 0)
                return seriesValues[measureIndex];

            return null;
        }

        private static getMetadata(grouped: DataViewValueColumnGroup[], source: DataViewMetadataColumn): ScatterChartMeasureMetadata {
            let xIndex = DataRoleHelper.getMeasureIndexOfRole(grouped, 'X');
            let yIndex = DataRoleHelper.getMeasureIndexOfRole(grouped, 'Y');
            let sizeIndex = DataRoleHelper.getMeasureIndexOfRole(grouped, 'Size');
            let xCol: DataViewMetadataColumn;
            let yCol: DataViewMetadataColumn;
            let sizeCol: DataViewMetadataColumn;
            let xAxisLabel = "";
            let yAxisLabel = "";

            if (grouped && grouped.length) {
                let firstGroup = grouped[0],
                    measureCount = firstGroup.values.length;

                if (!(xIndex >= 0))
                    xIndex = ScatterChart.getDefaultMeasureIndex(measureCount, yIndex, sizeIndex);
                if (!(yIndex >= 0))
                    yIndex = ScatterChart.getDefaultMeasureIndex(measureCount, xIndex, sizeIndex);
                if (!(sizeIndex >= 0))
                    sizeIndex = ScatterChart.getDefaultMeasureIndex(measureCount, xIndex, yIndex);

                if (xIndex >= 0) {
                    xCol = firstGroup.values[xIndex].source;
                    xAxisLabel = firstGroup.values[xIndex].source.displayName;
                }
                if (yIndex >= 0) {
                    yCol = firstGroup.values[yIndex].source;
                    yAxisLabel = firstGroup.values[yIndex].source.displayName;
                }
                if (sizeIndex >= 0) {
                    sizeCol = firstGroup.values[sizeIndex].source;
                }
            }

            return {
                idx: {
                    x: xIndex,
                    y: yIndex,
                    size: sizeIndex,
                },
                cols: {
                    x: xCol,
                    y: yCol,
                    size: sizeCol,
                },
                axesLabels: {
                    x: xAxisLabel,
                    y: yAxisLabel
                }
            };
        }

        private static getDefaultMeasureIndex(count: number, usedIndex: number, usedIndex2: number): number {
            for (let i = 0; i < count; i++) {
                if (i !== usedIndex && i !== usedIndex2)
                    return i;
            }
        }

        public setData(dataViews: DataView[]) {
            this.data = {
                xCol: undefined,
                yCol: undefined,
                dataPoints: [],
                legendData: { dataPoints: [] },
                axesLabels: { x: '', y: '' },
                selectedIds: [],
                sizeRange: [],
                dataLabelsSettings: dataLabelUtils.getDefaultPointLabelSettings(),
                defaultDataPointColor: null,
                hasDynamicSeries: false,
            };

            if (dataViews.length > 0) {
                let dataView = dataViews[0];

                if (dataView) {
                    this.categoryAxisProperties = CartesianHelper.getCategoryAxisProperties(dataView.metadata, true);
                    this.valueAxisProperties = CartesianHelper.getValueAxisProperties(dataView.metadata, true);
                    this.dataView = dataView;

                    if (dataView.categorical && dataView.categorical.values) {
                        this.data = ScatterChart.converter(dataView, this.currentViewport, this.colors, this.interactivityService, this.categoryAxisProperties, this.valueAxisProperties);
                    }

                }
            }
        }

        public calculateLegend(): LegendData {
            return this.data.legendData;
        }

        public hasLegend(): boolean {
            return this.data && this.data.hasDynamicSeries;
        }

        public enumerateObjectInstances(enumeration: ObjectEnumerationBuilder, options: EnumerateVisualObjectInstancesOptions): void {
            switch (options.objectName) {
                case 'dataPoint':
                    let categoricalDataView: DataViewCategorical = this.dataView && this.dataView.categorical ? this.dataView.categorical : null;
                    if (!GradientUtils.hasGradientRole(categoricalDataView))
                        return this.enumerateDataPoints(enumeration);
                case 'categoryAxis':
                    enumeration.pushInstance({
                        selector: null,
                        properties: {
                            showAxisTitle: !this.categoryAxisProperties || this.categoryAxisProperties["showAxisTitle"] == null ? true : this.categoryAxisProperties["showAxisTitle"]
                        },
                        objectName: 'categoryAxis'
                    });
                    break;
                case 'valueAxis':
                    enumeration.pushInstance({
                        selector: null,
                        properties: {
                            showAxisTitle: !this.valueAxisProperties || this.valueAxisProperties["showAxisTitle"] == null ? true : this.valueAxisProperties["showAxisTitle"]
                        },
                        objectName: 'valueAxis'
                    });
                    break;
                case 'categoryLabels':
                    if (this.data)
                        dataLabelUtils.enumerateCategoryLabels(enumeration, this.data.dataLabelsSettings, true);
                    else
                        dataLabelUtils.enumerateCategoryLabels(enumeration, null, true);
                    break;
                case 'fillPoint':
                    // Check if the card should be shown or not based on the existance of size measure
                    if (this.hasSizeMeasure())
                        return;

                    enumeration.pushInstance({
                        objectName: 'fillPoint',
                        selector: null,
                        properties: {
                            show: this.data.fillPoint,
                        },
                    });
                    break;
                case 'colorBorder':
                    // Check if the card should be shown or not based on the existance of size measure
                    if (this.hasSizeMeasure())
                        enumeration.pushInstance({
                            objectName: 'colorBorder',
                            selector: null,
                            properties: {
                                show: this.data.colorBorder,
                            },
                        });
                    break;
            }
        }

        private hasSizeMeasure(): boolean {
            let sizeRange = this.data.sizeRange;
            return sizeRange && sizeRange.min !== undefined;
        }

        private enumerateDataPoints(enumeration: ObjectEnumerationBuilder): void {
            let data = this.data;
            if (!data)
                return;

            let seriesCount = data.dataPoints.length;

            if (!data.hasDynamicSeries) {
                // Add default color and show all slices
                enumeration.pushInstance({
                    objectName: 'dataPoint',
                    selector: null,
                    properties: {
                        defaultColor: { solid: { color: data.defaultDataPointColor || this.colors.getColorByIndex(0).value } }
                    }
                }).pushInstance({
                    objectName: 'dataPoint',
                    selector: null,
                    properties: {
                        showAllDataPoints: !!data.showAllDataPoints
                    }
                });

                for (let i = 0; i < seriesCount; i++) {
                    let seriesDataPoints = data.dataPoints[i];
                    enumeration.pushInstance({
                        objectName: 'dataPoint',
                        displayName: seriesDataPoints.category,
                        selector: ColorHelper.normalizeSelector(seriesDataPoints.identity.getSelector(), /*isSingleSeries*/ true),
                        properties: {
                            fill: { solid: { color: seriesDataPoints.fill } }
                        },
                    });
                }
            }
            else {
                let legendDataPointLength = data.legendData.dataPoints.length;
                for (let i = 0; i < legendDataPointLength; i++) {
                    let series = data.legendData.dataPoints[i];
                    enumeration.pushInstance({
                        objectName: 'dataPoint',
                        displayName: series.label,
                        selector: ColorHelper.normalizeSelector(series.identity.getSelector()),
                        properties: {
                            fill: { solid: { color: series.color } }
                        },
                    });
                }
            }
        }

        public calculateAxesProperties(options: CalculateScaleAndDomainOptions): IAxisProperties[] {
            let data = this.data;
            let dataPoints = data.dataPoints;
            let viewport = this.currentViewport = options.viewport;
            let margin = options.margin;

            this.currentViewport = viewport;
            this.margin = margin;

            let width = viewport.width - (margin.left + margin.right);
            let height = viewport.height - (margin.top + margin.bottom);
            let minY = 0,
                maxY = 10,
                minX = 0,
                maxX = 10;
            if (dataPoints.length > 0) {
                minY = d3.min<ScatterChartDataPoint, number>(dataPoints, d => d.y);
                maxY = d3.max<ScatterChartDataPoint, number>(dataPoints, d => d.y);
                minX = d3.min<ScatterChartDataPoint, number>(dataPoints, d => d.x);
                maxX = d3.max<ScatterChartDataPoint, number>(dataPoints, d => d.x);
            }

            let xDomain = [minX, maxX];
            let combinedXDomain = AxisHelper.combineDomain(options.forcedXDomain, xDomain);

            this.xAxisProperties = AxisHelper.createAxis({
                pixelSpan: width,
                dataDomain: combinedXDomain,
                metaDataColumn: data.xCol,
                formatStringProp: scatterChartProps.general.formatString,
                outerPadding: 0,
                isScalar: true,
                isVertical: false,
                forcedTickCount: options.forcedTickCount,
                useTickIntervalForDisplayUnits: true,
                isCategoryAxis: true, //scatter doesn't have a categorical axis, but this is needed for the pane to react correctly to the x-axis toggle one/off
                scaleType: options.categoryAxisScaleType
            });
            this.xAxisProperties.axis.tickSize(-height, 0);
            this.xAxisProperties.axisLabel = this.data.axesLabels.x;

            let combinedDomain = AxisHelper.combineDomain(options.forcedYDomain, [minY, maxY]);

            this.yAxisProperties = AxisHelper.createAxis({
                pixelSpan: height,
                dataDomain: combinedDomain,
                metaDataColumn: data.yCol,
                formatStringProp: scatterChartProps.general.formatString,
                outerPadding: 0,
                isScalar: true,
                isVertical: true,
                forcedTickCount: options.forcedTickCount,
                useTickIntervalForDisplayUnits: true,
                isCategoryAxis: false,
                scaleType: options.valueAxisScaleType
            });
            this.yAxisProperties.axisLabel = this.data.axesLabels.y;

            return [this.xAxisProperties, this.yAxisProperties];
        }

        public overrideXScale(xProperties: IAxisProperties): void {
            this.xAxisProperties = xProperties;
        }

        public render(suppressAnimations: boolean): CartesianVisualRenderResult {
            if (!this.data)
                return;

            let data = this.data;
            let dataPoints = this.data.dataPoints;

            let margin = this.margin;
            let viewport = this.currentViewport;
            let width = viewport.width - (margin.left + margin.right);
            let height = viewport.height - (margin.top + margin.bottom);

            let hasSelection = this.interactivityService && this.interactivityService.hasSelection();

            this.mainGraphicsContext.attr('width', width)
                .attr('height', height);

            let duration = AnimatorCommon.GetAnimationDuration(this.animator, suppressAnimations);

            let scatterMarkers: D3.UpdateSelection;

            if (this.hasSizeMeasure()) {
                // Bubbles must be drawn from largest to smallest.
                dataPoints = dataPoints.sort(ScatterChart.sortBubbles);
                scatterMarkers = this.drawScatterMarkers(dataPoints, hasSelection, data.sizeRange, duration);
                scatterMarkers.order();
            }
            else {
                scatterMarkers = this.drawScatterMarkers(dataPoints, hasSelection, data.sizeRange, duration);
            }

            let labelDataPoints: LabelDataPoint[] = [];
            if (data.dataLabelsSettings && data.dataLabelsSettings.show || data.dataLabelsSettings.showCategory) {
                labelDataPoints = this.createLabelDataPoints();
            }

            let behaviorOptions: ScatterBehaviorOptions = undefined;
            if (this.interactivityService) {
                behaviorOptions = {
                    host: this.cartesianVisualHost,
                    root: this.svg,
                    dataPointsSelection: scatterMarkers,
                    mainContext: this.mainGraphicsContext,
                    data: this.data,
                    visualInitOptions: this.options,
                    xAxisProperties: this.xAxisProperties,
                    yAxisProperties: this.yAxisProperties,
                    background: d3.select(this.element.get(0)),
                };
            }

            TooltipManager.addTooltip(scatterMarkers, (tooltipEvent: TooltipEvent) => tooltipEvent.data.tooltipInfo);

            SVGUtil.flushAllD3TransitionsIfNeeded(this.options);

            return { dataPoints: data.dataPoints, behaviorOptions, labelDataPoints: labelDataPoints };
        }

        private drawScatterMarkers(scatterData: ScatterChartDataPoint[], hasSelection: boolean, sizeRange: NumberRange, duration: number) {
            let xScale = this.xAxisProperties.scale;
            let yScale = this.yAxisProperties.scale;
            let shouldEnableFill = (!sizeRange || !sizeRange.min) && this.data.fillPoint;
            let colorBorder = this.data.colorBorder;
            
            let markers = this.mainGraphicsContext.selectAll(ScatterChart.DotClasses.selector).data(scatterData, (d: ScatterChartDataPoint) => d.identity.getKey());

            markers.enter().append(ScatterChart.ScatterChartCircleTagName)
                .classed(ScatterChart.DotClasses.class, true);

            markers
                .style({
                'stroke-opacity': (d: ScatterChartDataPoint) => (d.size != null && colorBorder) ? 1 : ScatterChart.getBubbleOpacity(d, hasSelection),
                    'stroke-width': '1px',
                    'stroke': (d: ScatterChartDataPoint) => this.getStrokeFill(d, colorBorder),
                    'fill': (d: ScatterChartDataPoint) => d.fill,
                    'fill-opacity': (d: ScatterChartDataPoint) => (d.size != null || shouldEnableFill) ? ScatterChart.getBubbleOpacity(d, hasSelection) : 0,
                })
                .transition()
                .duration(duration)
                .attr({
                    r: (d: ScatterChartDataPoint) => ScatterChart.getBubbleRadius(d.radius, sizeRange, this.currentViewport),
                    cx: d => xScale(d.x),
                    cy: d => yScale(d.y),
                });

            markers.exit().remove();

            return markers;
        }

        private getStrokeFill(d: ScatterChartDataPoint, colorBorder: boolean): string {
            if (d.size != null && colorBorder) {
                let colorRgb = Color.parseRgb(d.fill);
                return Color.rgbToHexString(Color.darken(colorRgb, ScatterChart.StrokeDarkenColorValue));
            }
            return d.fill;
        }

        public static getBubblePixelAreaSizeRange(viewPort: IViewport, minSizeRange: number, maxSizeRange: number): DataRange {
            let ratio = 1.0;
            if (viewPort.height > 0 && viewPort.width > 0) {
                let minSize = Math.min(viewPort.height, viewPort.width);
                ratio = (minSize * minSize) / ScatterChart.AreaOf300By300Chart;
            }

            let minRange = Math.round(minSizeRange * ratio);
            let maxRange = Math.round(maxSizeRange * ratio);
            return {
                minRange: minRange,
                maxRange: maxRange,
                delta: maxRange - minRange
            };
        }

        public static project(value: number, actualSizeDataRange: DataRange, bubblePixelAreaSizeRange: DataRange): number {
            if (actualSizeDataRange.delta === 0 || bubblePixelAreaSizeRange.delta === 0) {
                return (ScatterChart.rangeContains(actualSizeDataRange, value)) ? bubblePixelAreaSizeRange.minRange : null;
            }

            let relativeX = (value - actualSizeDataRange.minRange) / actualSizeDataRange.delta;
            return bubblePixelAreaSizeRange.minRange + relativeX * bubblePixelAreaSizeRange.delta;
        }

        public static projectSizeToPixels(size: number, actualSizeDataRange: DataRange, bubblePixelAreaSizeRange: DataRange): number {
            let projectedSize = 0;
            if (actualSizeDataRange) {
                // Project value on the required range of bubble area sizes
                projectedSize = bubblePixelAreaSizeRange.maxRange;
                if (actualSizeDataRange.delta !== 0) {
                    let value = Math.min(Math.max(size, actualSizeDataRange.minRange), actualSizeDataRange.maxRange);
                    projectedSize = ScatterChart.project(value, actualSizeDataRange, bubblePixelAreaSizeRange);
                }

                projectedSize = Math.sqrt(projectedSize / Math.PI) * 2;
            }

            return Math.round(projectedSize);
        }

        public static rangeContains(range: DataRange, value: number): boolean {
            return range.minRange <= value && value <= range.maxRange;
        }

        public static getBubbleOpacity(d: ScatterChartDataPoint, hasSelection: boolean): number {
            if (hasSelection && !d.selected) {
                return ScatterChart.DimmedBubbleOpacity;
            }
            return ScatterChart.DefaultBubbleOpacity;
        }

        public onClearSelection(): void {
            if (this.interactivityService)
                this.interactivityService.clearSelection();
        }

        public getSupportedCategoryAxisType(): string {
            return axisType.scalar;
        }

        private createLabelDataPoints(): LabelDataPoint[]{
            let xScale = this.xAxisProperties.scale;
            let yScale = this.yAxisProperties.scale;
            let sizeRange = this.data.sizeRange;
            let labelDataPoints: LabelDataPoint[] = [];
            let dataPoints = this.data.dataPoints;
            let labelSettings = this.data.dataLabelsSettings;

            for (let dataPoint of dataPoints) {
                let text = dataPoint.category;

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
                    insideFill: NewDataLabelUtils.defaultInsideLabelColor,
                    isParentRect: false,
                    parentShape: {
                        point: {
                            x: xScale(dataPoint.x),
                            y: yScale(dataPoint.y),
                        },
                        radius: ScatterChart.getBubbleRadius(dataPoint.radius, sizeRange, this.currentViewport),
                        validPositions: ScatterChart.validLabelPositions,
                    }
                });
            }

            return labelDataPoints;
        }

        public static sortBubbles(a: ScatterChartDataPoint, b: ScatterChartDataPoint): number {
            let diff = (b.radius.sizeMeasure.values[b.radius.index] - a.radius.sizeMeasure.values[a.radius.index]);
            if (diff !== 0)
                return diff;

            // Tie-break equal size bubbles using identity.
            return b.identity.getKey().localeCompare(a.identity.getKey());
        }
    }
}
