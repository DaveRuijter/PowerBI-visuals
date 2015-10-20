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

    export interface PlayChartConstructorOptions extends CartesianVisualConstructorOptions {
        isFrozen?: boolean;
    }

    export interface PlayChartDataPoint extends SelectableDataPoint, TooltipEnabledDataPoint, LabelEnabledDataPoint {
        x: any;
        y: any;
        size: any;
        radius: RadiusData;
        fill: string;
        category: string;
    }

    export interface PlayChartData {
        xCol: DataViewMetadataColumn;
        yCol: DataViewMetadataColumn;
        dataPoints: PlayChartDataPoint[];
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
        frameKeys: any[];
        allDataPoints?: PlayChartDataPoint[][];
        currentFrameIndex?: number;
        lastRenderedFrameIndex?: number;
        colorByCategory?: boolean;
        currentViewport?: IViewport;
    }

    interface PlayChartMeasureMetadata {
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

    interface PlayObjectProperties {
        defaultDataPointColor?: string;
        showAllDataPoints?: boolean;
        colorByCategory: boolean;
        currentFrameIndex?: number;
    }

    export class PlayChart implements ICartesianVisual {
        private static PlayChartCircleTagName = 'circle';
        private static BubbleRadius = 3 * 2;
        public static DefaultBubbleOpacity = 0.85;
        public static DimmedBubbleOpacity = 0.3;
        public static StrokeDarkenColorValue = 255 * 0.25;
        // Chart Area and size range values as defined by PV charts
        private static AreaOf300By300Chart = 90000;
        private static MinSizeRange = 200;
        private static MaxSizeRange = 3000;
        private static ClassName = 'playChart';
        private static MainGraphicsContextClassName = 'mainGraphicsContext';
        private static DataLabelsContextClassName = 'dataLabelsContext';
        private static FrameDuration = 1000;
        private static FrameDurationFudge = 48; // animating the traceLine during play needs this to ensure we finish animating the line before the next frame plays
        private static SliderMarginLeft = 15;
        private static SliderMarginRight = 35;

        private static DotClasses: ClassAndSelector = {
            class: 'dot',
            selector: '.dot'
        };

        private svg: D3.Selection;
        private element: JQuery;
        private mainGraphicsContext: D3.Selection;
        private dataLabelsContext: D3.Selection;
        private clearCatcher: D3.Selection;
        private mainGraphicsG: D3.Selection;
        private currentViewport: IViewport;
        private lastRenderedViewport: IViewport;
        private style: IVisualStyle;
        private data: PlayChartData;
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
        private isFrozen: boolean;
        private frameCount: number;
        private isPlaying: boolean;
        private playAxisContainer: JQuery; //contains the playButton and slider
        private playButton: JQuery;
        private slider: JQuery;
        private callout: JQuery;

        // do not call converter() when we call persistProperties and a new update() happens
        // NOTE: calling persistProperties will still cause a render() call to come from cartesianChart
        private ridiculousFlagForPersistProperties: boolean;

        constructor(options: PlayChartConstructorOptions) {
            this.isFrozen = false;
            if (options) {
                this.interactivityService = options.interactivityService;
                this.animator = options.animator;
                this.isFrozen = !!options.isFrozen;
            }
        }

        public init(options: CartesianVisualInitOptions) {
            this.options = options;
            let element = this.element = options.element;
            this.currentViewport = options.viewport;
            this.lastRenderedViewport = {
                height: this.currentViewport.height,
                width: this.currentViewport.width
            };
            this.style = options.style;
            this.host = options.host;
            this.colors = this.style.colorPalette.dataColors;
            this.interactivity = options.interactivity;
            this.cartesianVisualHost = options.cartesianHost;
            this.isInteractiveChart = options.interactivity && options.interactivity.isInteractiveLegend;

            element.addClass(PlayChart.ClassName);
            let svg = this.svg = options.svg;
            this.clearCatcher = this.svg.select(".clearCatcher");

            this.mainGraphicsG = svg.append('g')
                .classed(PlayChart.MainGraphicsContextClassName, true);

            this.mainGraphicsContext = this.mainGraphicsG.append('svg');
            this.dataLabelsContext = this.mainGraphicsG.append('g')
                .classed(PlayChart.DataLabelsContextClassName, true);
        }

        private static convertMatrixToCategorical(matrix: DataViewMatrix, frame: number): DataViewCategorical {
            var node = matrix.rows.root.children[frame];

            var categorical: DataViewCategorical = {
                categories: [],
                values: powerbi.data.DataViewTransform.createValueColumns()
            };

            let category: DataViewCategoryColumn = {
                source: matrix.rows.levels[1].sources[0],
                values: [],
                objects: undefined,
                identity: []
            };

            if (matrix.columns.levels.length > 1) {
                categorical.values.source = matrix.columns.levels[0].sources[0];
                //let leafCount = 0;
                let leafOfGroup = matrix.columns.root;
                while (leafOfGroup.children && leafOfGroup.children[0].children) {
                    leafOfGroup = leafOfGroup.children[0];
                }

                let nodeQueue = [];
                let columnNode = matrix.columns.root;
                let leafOffset = 0;
                while (columnNode) {
                    if (columnNode.children && columnNode.children[0].children) {
                        for (var j = 0, jlen = columnNode.children.length; j < jlen; j++) {
                            nodeQueue.push(columnNode.children[j]);
                        }
                    } else if (columnNode.children) {
                        let columnLength = columnNode.children.length;
                        for (let j = 0; j < columnLength; j++) {
                            let source = <any>_.create(matrix.valueSources[j], { groupName: columnNode.value });
                            let dataViewColumn: DataViewValueColumn = {
                                identity: columnNode.identity,
                                values: [],
                                source: source,
                                value: columnNode.value,
                            };
                            categorical.values.push(dataViewColumn);
                        }

                        for (let i = 0, len = node.children.length; i < len; i++) {
                            let innerNode = node.children[i];
                            category.identity.push(innerNode.identity);
                            category.values.push(innerNode.value);

                            for (let j = 0; j < columnLength; j++) {
                                categorical.values[j+leafOffset].values.push(innerNode.values[j+leafOffset].value);
                            }
                        }

                        leafOffset += columnLength;
                    }

                    if (nodeQueue.length > 0) {
                        columnNode = nodeQueue[0];
                        nodeQueue = nodeQueue.splice(1);
                    } else
                        columnNode = undefined;
                }
            } else {
                let columnLength = matrix.columns.root.children.length;
                for (let j = 0; j < columnLength; j++) {
                    let dataViewColumn: DataViewValueColumn = {
                        identity: undefined,
                        values: [],
                        source: matrix.valueSources[j]
                    };
                    categorical.values.push(dataViewColumn);
                }

                for (let i = 0, len = node.children.length; i < len; i++) {
                    let innerNode = node.children[i];
                    category.identity.push(innerNode.identity);
                    category.values.push(innerNode.value);

                    // v1, no series.
                    for (let j = 0; j < columnLength; j++) {
                        categorical.values[j].values.push(innerNode.values[j].value);
                    }
                }
            }
            categorical.categories.push(category);

            return categorical;
        }

        private static getObjectProperties(dataViewMetadata: DataViewMetadata, dataLabelsSettings?: PointDataLabelsSettings): PlayObjectProperties {
            let objectProperties: PlayObjectProperties = {
                colorByCategory: true, //default to true
            };
            if (dataViewMetadata && dataViewMetadata.objects) {
                var objects = dataViewMetadata.objects;

                objectProperties.defaultDataPointColor = DataViewObjects.getFillColor(objects, columnChartProps.dataPoint.defaultColor);
                objectProperties.showAllDataPoints = DataViewObjects.getValue<boolean>(objects, columnChartProps.dataPoint.showAllDataPoints);

                var labelsObj = objects['categoryLabels'];
                if (labelsObj && dataLabelsSettings) {
                    dataLabelsSettings.show = (labelsObj['show'] !== undefined) ? <boolean>labelsObj['show'] : dataLabelsSettings.show;
                    dataLabelsSettings.precision = (labelsObj['labelsPrecision'] !== undefined) ? +<string>labelsObj['labelsPrecision'] : dataLabelsSettings.precision;
                    if (labelsObj['color'] !== undefined) {
                        dataLabelsSettings.labelColor = (<Fill>labelsObj['color']).solid.color;
                    }
                }

                var generalObj = objects['general'];
                if (generalObj) {
                    if (generalObj['colorByCategory'] != null)
                        objectProperties.colorByCategory = !!generalObj['colorByCategory'];
                    if (generalObj['frozenIndex'] != null)
                        objectProperties.currentFrameIndex = <number>generalObj['frozenIndex'];
                }
            }
            return objectProperties;
        }

        public static converter(dataView: DataView, currentViewport: IViewport, colorPalette: IDataColorPalette, interactivityService?: IInteractivityService, categoryAxisProperties?: DataViewObject, valueAxisProperties?: DataViewObject): PlayChartData {
            var categoryValues: any[],
                categoryFormatter: IValueFormatter,
                categoryObjects: DataViewObjects[],
                categoryIdentities: DataViewScopeIdentity[];

            let dataViewCategorical: DataViewCategorical = dataView.categorical;
            if (!dataViewCategorical && dataView.matrix) {
                dataViewCategorical = PlayChart.convertMatrixToCategorical(dataView.matrix, dataView.matrix.rows.root.children.length - 1);
            }
            let dataViewMetadata: DataViewMetadata = dataView.metadata;

            if (dataViewCategorical.categories && dataViewCategorical.categories.length > 0) {
                categoryValues = dataViewCategorical.categories[0].values;
                categoryFormatter = valueFormatter.create({ format: valueFormatter.getFormatString(dataViewCategorical.categories[0].source, playChartProps.general.formatString), value: categoryValues[0], value2: categoryValues[categoryValues.length - 1] });
                categoryIdentities = dataViewCategorical.categories[0].identity;
                categoryObjects = dataViewCategorical.categories[0].objects;
            }
            else {
                categoryValues = [null];
                // creating default formatter for null value (to get the right string of empty value from the locale)
                categoryFormatter = valueFormatter.createDefaultFormatter(null); 
            }

            let categories = dataViewCategorical.categories;
            let dataValues = dataViewCategorical.values;

            if (!dataValues)
                return;

            let hasDynamicSeries = !!dataValues.source;
            let grouped = dataValues.grouped();
            let dvSource = dataValues.source;
            let playMetadata = PlayChart.getMetadata(grouped, dvSource);
            let dataLabelsSettings = dataLabelUtils.getDefaultPointLabelSettings();
            let objectProperties = PlayChart.getObjectProperties(dataViewMetadata, dataLabelsSettings);

            let allDataPoints = [];
            let frameKeys = [];
            let dataPoints = undefined;
            let rowChildrenLength = dataView.matrix.rows.root.children.length;
            if (!dataView.categorical && dataView.matrix && rowChildrenLength > 0) {
                let keySourceColumn = dataView.matrix.rows.levels[0].sources[0];
                let keyFormatter = valueFormatter.create({
                    format: valueFormatter.getFormatString(keySourceColumn, playChartProps.general.formatString),
                    value: dataView.matrix.rows.root.children[0],
                    value2: dataView.matrix.rows.root.children[rowChildrenLength-1],
                });
                for (let i = 0, len = rowChildrenLength; i < len; i++) {
                    var key = dataView.matrix.rows.root.children[i];
                    frameKeys.push(keyFormatter.format(key.value));

                    dataViewCategorical = PlayChart.convertMatrixToCategorical(dataView.matrix, i);
                    categoryValues = dataViewCategorical.categories[0].values;
                    categoryFormatter = valueFormatter.create({
                        format: valueFormatter.getFormatString(dataViewCategorical.categories[0].source, playChartProps.general.formatString),
                        value: categoryValues[0],
                        value2: categoryValues[categoryValues.length - 1],
                    });
                    categoryIdentities = dataViewCategorical.categories[0].identity;
                    categoryObjects = dataViewCategorical.categories[0].objects;

                    categories = dataViewCategorical.categories;
                    dataValues = dataViewCategorical.values;

                    dataPoints = PlayChart.createDataPoints(
                        dataValues,
                        playMetadata,
                        categories,
                        categoryValues,
                        categoryFormatter,
                        categoryIdentities,
                        categoryObjects,
                        colorPalette,
                        currentViewport,
                        hasDynamicSeries,
                        dataLabelsSettings,
                        objectProperties.defaultDataPointColor,
                        objectProperties.colorByCategory);
                    allDataPoints.push(dataPoints);
                }
            }
            else {
                dataPoints = PlayChart.createDataPoints(
                    dataValues,
                    playMetadata,
                    categories,
                    categoryValues,
                    categoryFormatter,
                    categoryIdentities,
                    categoryObjects,
                    colorPalette,
                    currentViewport,
                    hasDynamicSeries,
                    dataLabelsSettings,
                    objectProperties.defaultDataPointColor,
                    objectProperties.colorByCategory);
            }

            // use the saved frame index, or default to the last frame
            if (objectProperties.currentFrameIndex == null) {
                objectProperties.currentFrameIndex = frameKeys.length - 1; //default
                // current dataPoints is already the last frame
            }
            else if (objectProperties.currentFrameIndex < frameKeys.length) {
                dataPoints = allDataPoints[objectProperties.currentFrameIndex];
            }

            if (interactivityService) {
                interactivityService.applySelectionStateToData(dataPoints);
            }

            let legendItems = hasDynamicSeries
                ? PlayChart.createSeriesLegend(dataValues, colorPalette, dataValues,
                    valueFormatter.getFormatString(dvSource, playChartProps.general.formatString),
                    objectProperties.defaultDataPointColor)
                : [];

            let legendTitle = dataValues && dvSource ? dvSource.displayName : "";
            if (!legendTitle) {
                legendTitle = categories && categories[0].source.displayName ? categories[0].source.displayName : "";
            }

            let legendData = { title: legendTitle, dataPoints: legendItems };
            let sizeRange = PlayChart.getSizeRangeForGroups(grouped, playMetadata.idx.size);

            if (categoryAxisProperties && categoryAxisProperties["showAxisTitle"] !== null && categoryAxisProperties["showAxisTitle"] === false) {
                playMetadata.axesLabels.x = null;
            }

            if (valueAxisProperties && valueAxisProperties["showAxisTitle"] !== null && valueAxisProperties["showAxisTitle"] === false) {
                playMetadata.axesLabels.y = null;
            }

            return {
                xCol: playMetadata.cols.x,
                yCol: playMetadata.cols.y,
                dataPoints: dataPoints,
                legendData: legendData,
                axesLabels: playMetadata.axesLabels,
                hasSelection: false,
                selectedIds: [],
                size: playMetadata.cols.size,
                sizeRange: sizeRange,
                dataLabelsSettings: dataLabelsSettings,
                defaultDataPointColor: objectProperties.defaultDataPointColor,
                hasDynamicSeries: hasDynamicSeries,
                showAllDataPoints: objectProperties.showAllDataPoints,
                allDataPoints: allDataPoints,
                frameKeys: frameKeys,
                colorByCategory: objectProperties.colorByCategory,
                currentFrameIndex: objectProperties.currentFrameIndex,
            };
        }

        private static getSizeRangeForGroups(
            dataViewValueGroups: DataViewValueColumnGroup[],
            sizeColumnIndex: number): NumberRange {

            let result: NumberRange = {};
            if (dataViewValueGroups) {
                dataViewValueGroups.forEach((group) => {
                    let sizeColumn = PlayChart.getMeasureValue(sizeColumnIndex, group.values);
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
            metadata: PlayChartMeasureMetadata,
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
            colorByCategory?: boolean): PlayChartDataPoint[]{

            let dataPoints: PlayChartDataPoint[] = [],
                indicies = metadata.idx,
                formatStringProp = playChartProps.general.formatString,
                dataValueSource = dataValues.source,
                grouped = dataValues.grouped();

            let colorHelper = new ColorHelper(colorPalette, playChartProps.dataPoint.fill, defaultDataPointColor);

            for (let categoryIdx = 0, ilen = categoryValues.length; categoryIdx < ilen; categoryIdx++) {
                let categoryValue = categoryValues[categoryIdx];

                for (let seriesIdx = 0, len = grouped.length; seriesIdx < len; seriesIdx++) {
                    let grouping = grouped[seriesIdx];
                    let seriesValues = grouping.values;
                    let measureX = PlayChart.getMeasureValue(indicies.x, seriesValues);
                    let measureY = PlayChart.getMeasureValue(indicies.y, seriesValues);
                    let measureSize = PlayChart.getMeasureValue(indicies.size, seriesValues);

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
                    else if (colorByCategory) {
                        color = colorHelper.getColorForSeriesValue(categoryObjects && categoryObjects[categoryIdx], dataValues.identityFields, categoryValue);
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

                    let dataPoint: PlayChartDataPoint = {
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
            // TODO: is it more efficient to calculate min/max for the frame here and pass back {dp, min, max}?
            return dataPoints;
        }

        private static createSeriesLegend(
            dataValues: DataViewValueColumns,
            colorPalette: IDataColorPalette,
            categorical: DataViewValueColumns,
            formatString: string,
            defaultDataPointColor: string): LegendDataPoint[] {

            let grouped = dataValues.grouped();
            let colorHelper = new ColorHelper(colorPalette, playChartProps.dataPoint.fill, defaultDataPointColor);

            let legendItems: LegendDataPoint[] = [];
            for (let i = 0, len = grouped.length; i < len; i++) {
                let grouping = grouped[i];
                let color = colorHelper.getColorForSeriesValue(grouping.objects, dataValues.identityFields, grouping.name);
                legendItems.push({
                    color: color,
                    icon: LegendIcon.Circle,
                    label: valueFormatter.format(grouping.name, formatString),
                    identity: grouping.identity ? SelectionId.createWithId(grouping.identity) : SelectionId.createNull(),
                    selected: false,
                });
            }

            return legendItems;
        }

        public static getBubbleRadius(radiusData: RadiusData, sizeRange: NumberRange, viewPort: IViewport): number {
            let actualSizeDataRange = null;
            let bubblePixelAreaSizeRange = null;
            let measureSize = radiusData.sizeMeasure;

            if (!measureSize)
                return PlayChart.BubbleRadius;

            let minSize = sizeRange.min ? sizeRange.min : 0;
            let maxSize = sizeRange.max ? sizeRange.max : 0;

            let min = Math.min(minSize, 0);
            let max = Math.max(maxSize, 0);
            actualSizeDataRange = {
                minRange: min,
                maxRange: max,
                delta: max - min
            };

            bubblePixelAreaSizeRange = PlayChart.getBubblePixelAreaSizeRange(viewPort, PlayChart.MinSizeRange, PlayChart.MaxSizeRange);

            if (measureSize.values) {
                let sizeValue = measureSize.values[radiusData.index];
                if (sizeValue != null) {
                    return PlayChart.projectSizeToPixels(sizeValue, actualSizeDataRange, bubblePixelAreaSizeRange) / 2;
                }
            }

            return PlayChart.BubbleRadius;
        }

        public static getMeasureValue(measureIndex: number, seriesValues: DataViewValueColumn[]): DataViewValueColumn {
            if (measureIndex >= 0)
                return seriesValues[measureIndex];

            return null;
        }

        private static getMetadata(grouped: DataViewValueColumnGroup[], source: DataViewMetadataColumn): PlayChartMeasureMetadata {
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
                    xIndex = PlayChart.getDefaultMeasureIndex(measureCount, yIndex, sizeIndex);
                if (!(yIndex >= 0))
                    yIndex = PlayChart.getDefaultMeasureIndex(measureCount, xIndex, sizeIndex);
                if (!(sizeIndex >= 0))
                    sizeIndex = PlayChart.getDefaultMeasureIndex(measureCount, xIndex, yIndex);

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
            let defaultData = {
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
                frameKeys: [],
            };

            if (dataViews.length > 0 && dataViews[0] != null) {
                let dataView = dataViews[0];

                this.categoryAxisProperties = CartesianHelper.getCategoryAxisProperties(dataView.metadata, true);
                this.valueAxisProperties = CartesianHelper.getValueAxisProperties(dataView.metadata, true);
                this.dataView = dataView;

                if (this.ridiculousFlagForPersistProperties && dataView.metadata) {
                    let objectProps = PlayChart.getObjectProperties(dataView.metadata);
                    // only copy frameIndex since it is the only property using persistProperties
                    if (this.data) {
                        this.data.currentFrameIndex = objectProps.currentFrameIndex;
                    }
                    // turn off the flag that was set by our persistProperties call
                    this.ridiculousFlagForPersistProperties = false;
                }
                else if (dataView.matrix || dataView.categorical) {
                    let data = PlayChart.converter(dataView, this.currentViewport, this.colors, this.interactivityService, this.categoryAxisProperties, this.valueAxisProperties);
                    if (data) {
                        this.data = data;
                        if (this.interactivityService && this.data.dataPoints) {
                            this.interactivityService.applySelectionStateToData(this.data.dataPoints);
                        }
                    }
                    else
                        this.data = defaultData;
                }
                else
                    this.data = defaultData;
            }
            else {
                this.data = defaultData;
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
                case 'general':
                    enumeration.pushInstance({
                        selector: null,
                        properties: {
                            colorByCategory: this.data ? this.data.colorByCategory : true,
                        },
                        objectName: 'general'
                    });
                    break;
                // NOT WORKING RIGHT NOW
                case 'dataPoint - butNotWorkingAndLeavingHereForTSLINT':
                    let categoricalDataView: DataViewCategorical = this.dataView && this.dataView.categorical ? this.dataView.categorical : null;
                    if (!GradientUtils.hasGradientRole(categoricalDataView))
                        return this.enumerateDataPoints(enumeration);
                    break;
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
                        defaultColor: { solid: { color: data.defaultDataPointColor || this.colors.getColorByIndex(0).value } },
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
                            fill: { solid: { color: series.color } },
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

            if (this.data) {
                this.data.currentViewport = {
                    height: viewport.height,
                    width: viewport.width
                };
            }

            let width = viewport.width - (margin.left + margin.right);
            let height = viewport.height - (margin.top + margin.bottom);
            let minY = 0,
                maxY = 10,
                minX = 0,
                maxX = 10;
            if (dataPoints.length > 0) {
                if (data.allDataPoints && data.allDataPoints.length > 0) {
                    minY = minX = Number.MAX_VALUE;
                    maxY = maxX = Number.MIN_VALUE;
                    for (let i = 0, len = data.allDataPoints.length; i < len; i++) {
                        let dps = data.allDataPoints[i];
                        minY = Math.min(d3.min<PlayChartDataPoint, number>(dps, d => d.y), minY);
                        maxY = Math.max(d3.max<PlayChartDataPoint, number>(dps, d => d.y), maxY);
                        minX = Math.min(d3.min<PlayChartDataPoint, number>(dps, d => d.x), minX);
                        maxX = Math.max(d3.max<PlayChartDataPoint, number>(dps, d => d.x), maxX);
                    }
                } else {
                    minY = d3.min<PlayChartDataPoint, number>(dataPoints, d => d.y);
                    maxY = d3.max<PlayChartDataPoint, number>(dataPoints, d => d.y);
                    minX = d3.min<PlayChartDataPoint, number>(dataPoints, d => d.x);
                    maxX = d3.max<PlayChartDataPoint, number>(dataPoints, d => d.x);
                }
            }

            let xDomain = [minX, maxX];
            let combinedXDomain = AxisHelper.combineDomain(options.forcedXDomain, xDomain);

            this.xAxisProperties = AxisHelper.createAxis({
                pixelSpan: width,
                dataDomain: combinedXDomain,
                metaDataColumn: data.xCol,
                formatStringProp: playChartProps.general.formatString,
                outerPadding: 0,
                isScalar: true,
                isVertical: false,
                forcedTickCount: options.forcedTickCount,
                useTickIntervalForDisplayUnits: true,
                isCategoryAxis: true,//Play doesn't have a categorical axis, but this is needed for the pane to react correctly to the x-axis toggle one/off
                scaleType: options.categoryAxisScaleType,
            });
            this.xAxisProperties.axis.tickSize(-height, 0);
            this.xAxisProperties.axisLabel = this.data.axesLabels.x;

            let combinedDomain = AxisHelper.combineDomain(options.forcedYDomain, [minY, maxY]);

            this.yAxisProperties = AxisHelper.createAxis({
                pixelSpan: height,
                dataDomain: combinedDomain,
                metaDataColumn: data.yCol,
                formatStringProp: playChartProps.general.formatString,
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

        private createSliderDOM(): void {
            // a container to position the button and the slider together
            this.playAxisContainer = $('<div></div>')
                .css('position', 'absolute')
                .css('left', '15px')
                .css('bottom', '35px')
                .appendTo(this.element);

            var playButtonCircle = $('<div class="button-container"></div>')
                .appendTo(this.playAxisContainer)
                .on('click', () => {
                    this.play();
                });
            this.playButton = $('<div class="play"></div>')
                .appendTo(playButtonCircle);

            // the noUiSlider
            this.slider = $('<div class="sliders"></div>')
                .appendTo(this.playAxisContainer);
        }

        private createSliderControl(slider: any, sliderWidth: number): void {
            // re-create the slider
            if ((<any>slider[0]).noUiSlider)
                (<any>slider[0]).noUiSlider.destroy();
            
            var numFrames = this.data.frameKeys.length;
            if (numFrames > 0) {
                var filterPipLabels = this.createPipsFilterFn(sliderWidth);
                var lastIndex = numFrames - 1;
                noUiSlider.create(
                    slider[0],
                    {
                        step: 1,
                        start: [this.data.currentFrameIndex],
                        range: {
                            min: [0],
                            max: [lastIndex],
                        },
                        pips: {
                            mode: 'steps',
                            density: 12,
                            format: {
                                to: (index) => this.data.frameKeys[index],
                                from: (value) => this.data.frameKeys.indexOf(value),
                            },
                            filter: filterPipLabels,
                        },
                    });

                (<any>slider[0]).noUiSlider.on('slide', () => {
                    this.isPlaying = false; //stop the play sequence
                    var indexToShow = Math.round((<any>slider[0]).noUiSlider.get());
                    if (indexToShow >= 0 && indexToShow !== this.data.currentFrameIndex && indexToShow < this.data.allDataPoints.length) {
                        this.data.currentFrameIndex = indexToShow;
                        var data = this.data.allDataPoints[indexToShow];
                        this.data.dataPoints = data;
                        this.persistFrameIndex(indexToShow); //this will cause a render
                    }
                });
            }
            else {
                noUiSlider.create(
                    slider[0],
                    {
                        step: 1,
                        start: [0],
                        range: {
                            min: [0],
                            max: [0],
                        },
                    });
            }
        }

        private createPipsFilterFn(sliderWidth: number): any {
            let textProperties: TextProperties = {
                fontFamily: 'wf_segoe-ui_normal',
                fontSize: jsCommon.PixelConverter.toString(14),
            };
            let maxLabelWidth = 0;
            let anyWillWordBreak = false;
            for (var key of this.data.frameKeys) {
                maxLabelWidth = Math.max(maxLabelWidth, jsCommon.WordBreaker.getMaxWordWidth(key + "", TextMeasurementService.measureSvgTextWidth, textProperties));
                anyWillWordBreak = anyWillWordBreak || jsCommon.WordBreaker.hasBreakers(key);
            }

            var pipSize = 1; //0=hide, 1=large, 2=small
            var skipMod = 1;
            var maxAllowedLabelWidth = this.frameCount > 1 ? sliderWidth / (this.frameCount - 1) : sliderWidth;
            var widthRatio = maxLabelWidth / maxAllowedLabelWidth;

            if (widthRatio > 1.25) {
                skipMod = Math.ceil(widthRatio);
                pipSize = 2;
            }
            else if (widthRatio > 1.0 || anyWillWordBreak) {
                // wordbreak line wrapping is automatic, and we don't reserve enough space to show two lines of text with the larger font
                pipSize = 2;
            }

            var filterPipLabels = (index: any, type: any) => {
                // noUiSlider will word break / wrap to new lines, so max width is the max word length
                if (index % skipMod === 0) {
                    return pipSize;
                }
                return 0; //hide
            };

            return filterPipLabels;
        }

        public render(suppressAnimations: boolean): CartesianVisualRenderResult {
            if (!this.data || !this.dataView)
                return { dataPoints: [], behaviorOptions: null, labelDataPoints: [] };

            let data = this.data;
            let dataPoints = this.data.dataPoints;

            let margin = this.margin;
            let viewport = this.currentViewport;
            let width = viewport.width - (margin.left + margin.right);
            let height = viewport.height - (margin.top + margin.bottom);
            let xScale = this.xAxisProperties.scale;
            let yScale = this.yAxisProperties.scale;

            var hasSelection = false;
            if (this.interactivityService) {
                this.interactivityService.applySelectionStateToData(dataPoints);
                hasSelection = this.interactivityService.hasSelection();
            }

            this.mainGraphicsContext
                .attr('width', width)
                .attr('height', height);

            let sortedData = dataPoints.sort(function (a, b) {
                return b.radius.sizeMeasure ? (b.radius.sizeMeasure.values[b.radius.index] - a.radius.sizeMeasure.values[a.radius.index]) : 0;
            });

            var playMarkers = this.drawPlayMarkers(sortedData, hasSelection, data.sizeRange, suppressAnimations);
            var labelsSelection: D3.Selection;
            if (this.data.dataLabelsSettings.show) {
                var layout = dataLabelUtils.getScatterChartLabelLayout(xScale, yScale, this.data.dataLabelsSettings, viewport, data.sizeRange);
                labelsSelection = dataLabelUtils.drawDefaultLabelsForDataPointChart(
                    dataPoints,
                    this.dataLabelsContext,
                    layout,
                    this.currentViewport,
                    !!this.animator && !suppressAnimations,
                    PlayChart.FrameDuration,
                    hasSelection);
            }
            else {
                dataLabelUtils.cleanDataLabels(this.dataLabelsContext);
            }

            TooltipManager.addTooltip(playMarkers, (tooltipEvent: TooltipEvent) => tooltipEvent.data.tooltipInfo);

            if (this.data) {
                if (!this.playAxisContainer && !this.isFrozen) {
                    this.createSliderDOM();
                }
                let frameKeys = this.data.frameKeys;
                let currentFrameIndex = this.data.currentFrameIndex;

                // callout / currentFrameIndex label
                if (!this.callout) {
                    this.callout = $('<span class="callout"></span>').appendTo(this.element);
                }

                var slider = this.slider;
                var sliderWidth = (width - PlayChart.SliderMarginLeft - PlayChart.SliderMarginRight);
                if (slider)
                    slider.css('width', sliderWidth + 'px');

                // TODO: make this smarter
                let resized = this.currentViewport.width !== this.lastRenderedViewport.width
                    || this.currentViewport.height !== this.lastRenderedViewport.height;
                let changed = !this.playAxisContainer
                    || frameKeys.length !== this.frameCount
                    || resized;

                // default to last frame if frameKeys have changed and it's not the first time we are rendering
                if (this.frameCount && (frameKeys.length !== this.frameCount || currentFrameIndex >= frameKeys.length))
                    currentFrameIndex = frameKeys.length - 1;

                if (changed && slider) {
                    this.frameCount = frameKeys.length;
                    this.createSliderControl(slider, sliderWidth);
                }

                // update callout to current frame index
                var calloutDimension = Math.min(height, width * 1.3); //compensate for tall and narrow-width viewport
                var fontSize = Math.max(12, Math.round(calloutDimension / 7)) + 'px';
                if (currentFrameIndex < frameKeys.length && currentFrameIndex >= 0) {
                    this.callout
                        .text(frameKeys[currentFrameIndex])
                        .css('font-size', fontSize);
                }
                else {
                    this.callout.text('');
                }
                
                if (slider) {
                    // ensure slider position
                    (<any>slider[0]).noUiSlider.set([currentFrameIndex]);
                }

                this.lastRenderedViewport = {
                    height: this.currentViewport.height,
                    width: this.currentViewport.width
                };
            }

            if (this.interactivityService) {
                var behaviorOptions: PlayBehaviorOptions = {
                    host: this.cartesianVisualHost,
                    root: this.svg,
                    dataPointsSelection: playMarkers,
                    labelsSelection: labelsSelection,
                    mainContext: this.mainGraphicsContext,
                    data: this.data,
                    visualInitOptions: this.options,
                    xAxisProperties: this.xAxisProperties,
                    yAxisProperties: this.yAxisProperties,
                    background: d3.select(this.element.get(0)),
                    clearCatcher: this.clearCatcher,
                    dataViewCat: this.dataView.categorical,
                    svg: this.mainGraphicsContext,
                    dataView: this.dataView,
                    renderTraceLine: PlayChart.renderTraceLine,
                };

                if (hasSelection) {
                    let flatAllDataPoints = _.flatten<SelectableDataPoint>(this.data.allDataPoints);
                    let uniqueDataPoints = _.uniq(flatAllDataPoints, (d: PlayChartDataPoint) => d.identity.getKey());
                    PlayChart.renderTraceLine(behaviorOptions, _.filter(uniqueDataPoints, 'selected'), !suppressAnimations);
                }
            }

            SVGUtil.flushAllD3TransitionsIfNeeded(this.options);

            // pass allDataPoints to keep track of current selected bubble even if it drops out for a few frames
            return { dataPoints: _.flatten<SelectableDataPoint>(this.data.allDataPoints), behaviorOptions: behaviorOptions, labelDataPoints: [] };
        }

        public static renderTraceLine(options: PlayBehaviorOptions, selectedPoints: SelectableDataPoint[], shouldAnimate: boolean): void {
            var seriesPoints: PlayChartDataPoint[][] = [];

            if (selectedPoints && selectedPoints.length > 0) {
                let currentFrameIndex = options.data.currentFrameIndex;
                let lastRenderedFrameIndex = options.data.lastRenderedFrameIndex;

                // filter to the selected identity, only up to and including the current frame. Add frames during play.
                let hasBubbleAtCurrentFrame = [];
                for (var selectedIndex = 0, selectedLen = selectedPoints.length; selectedIndex < selectedLen; selectedIndex++) {
                    seriesPoints[selectedIndex] = [];
                    hasBubbleAtCurrentFrame[selectedIndex] = false;
                    for (let frameIndex = 0, frameLen = options.data.allDataPoints.length; frameIndex < frameLen && frameIndex <= currentFrameIndex; frameIndex++) {
                        let values = options.data.allDataPoints[frameIndex].filter((value, index) => {
                            return value.identity.getKey() === selectedPoints[selectedIndex].identity.getKey();
                        });
                        if (values && values.length > 0) {
                            seriesPoints[selectedIndex].push(values[0]);
                            if (frameIndex === currentFrameIndex)
                                hasBubbleAtCurrentFrame[selectedIndex] = true;
                        }
                    }
                }
                if (seriesPoints.length > 0) {
                    let xScale = options.xAxisProperties.scale;
                    let yScale = options.yAxisProperties.scale;

                    let line = d3.svg.line()
                        .x((d: PlayChartDataPoint) => {
                            return xScale(d.x);
                        })
                        .y((d: PlayChartDataPoint) => {
                            return yScale(d.y);
                        })
                        .defined((d: PlayChartDataPoint) => {
                            return d.x !== null && d.y !== null;
                        });

                    // Render Lines
                    let traceLines = options.svg.selectAll('.traceLine').data(selectedPoints, (sp: PlayChartDataPoint) => sp.identity.getKey());
                    traceLines.enter()
                        .append('path')
                        .classed('traceLine', true);
                    // prepare array of new/previous lengths
                    // NOTE: can't use lambda because we need the "this" context to be the DOM Element associated with the .each()
                    let previousLengths = [], newLengths = [];
                    traceLines.each(function (d, i) {
                        let existingPath = (<SVGPathElement>this);
                        let previousLength = existingPath.hasAttribute('d') ? existingPath.getTotalLength() : 0;
                        previousLengths.push(previousLength);
                        // create offline SVG for new path measurement
                        let tempSvgPath = $('<svg><path></path></svg>');
                        let tempPath = $('path', tempSvgPath);
                        tempPath.attr('d', line(seriesPoints[i]));
                        let newLength = (<SVGPathElement>tempPath.get()[0]).getTotalLength();
                        newLengths.push(newLength);
                    });
                    // animate using stroke-dash* trick
                    if (lastRenderedFrameIndex == null || currentFrameIndex >= lastRenderedFrameIndex) {
                        // growing line
                        traceLines
                            .style('stroke', (d: PlayChartDataPoint) => PlayChart.getStrokeFill(d, true))
                            .attr({
                                'd': (d, i: number) => {
                                    return line(seriesPoints[i]);
                                },
                                'stroke-dasharray': (d, i) => newLengths[i] + " " + newLengths[i],
                                'stroke-dashoffset': (d, i) => newLengths[i] - previousLengths[i],
                            });
                        if (shouldAnimate) {
                            traceLines
                                .transition()
                                .duration(PlayChart.FrameDuration - PlayChart.FrameDurationFudge)
                                .attr('stroke-dashoffset', 0);
                        }
                        else {
                            traceLines.attr('stroke-dashoffset', 0);
                        }
                    }
                    else {
                        // shrinking line
                        if (shouldAnimate) {
                            traceLines
                                .transition()
                                .duration(PlayChart.FrameDuration - PlayChart.FrameDurationFudge)
                                .attr('stroke-dashoffset', (d, i) => previousLengths[i] - newLengths[i])
                                .transition()
                                .duration(1) // animate the shrink first, then update with new line properties
                                .delay(PlayChart.FrameDuration - PlayChart.FrameDurationFudge)
                                .style('stroke', (d: PlayChartDataPoint) => PlayChart.getStrokeFill(d, true))
                                .attr({
                                    'd': (d, i) => {
                                        return line(seriesPoints[i]);
                                    },
                                    'stroke-dasharray': (d, i) => newLengths[i] + " " + newLengths[i],
                                    'stroke-dashoffset': 0,
                                });
                        }
                        else {
                            traceLines
                                .style('stroke', (d: PlayChartDataPoint) => PlayChart.getStrokeFill(d, true))
                                .attr({
                                    'd': (d, i) => {
                                        return line(seriesPoints[i]);
                                    },
                                    'stroke-dasharray': (d, i) => newLengths[i] + " " + newLengths[i],
                                    'stroke-dashoffset': 0,
                                });
                        }
                    }
                    traceLines.exit()
                        .remove();

                    // Render circles
                    // slice to length-1 because we draw lines to the current bubble but we don't need to draw the current frame's bubble
                    let circlePoints: PlayChartDataPoint[] = [];
                    for (var selectedIndex = 0; selectedIndex < seriesPoints.length; selectedIndex++) {
                        let points = seriesPoints[selectedIndex];
                        let newPoints = hasBubbleAtCurrentFrame[selectedIndex] ? points.slice(0, points.length - 1) : points;
                        circlePoints = circlePoints.concat(newPoints);
                    }
                    let circles = options.svg.selectAll('.traceBubble').data(circlePoints, (d: PlayChartDataPoint) => d.identity.getKey() + d.x + d.y + d.size);
                    circles.enter()
                        .append('circle')
                        .style('opacity', 0) //fade new bubbles into visibility
                        .classed('traceBubble', true);
                    circles
                        .attr('cx', (d: PlayChartDataPoint) => xScale(d.x))
                        .attr('cy', (d: PlayChartDataPoint) => yScale(d.y))
                        .attr('r', (d: PlayChartDataPoint) => PlayChart.getBubbleRadius(d.radius, options.data.sizeRange, options.data.currentViewport))
                        .style({
                            'stroke-opacity': (d: PlayChartDataPoint) => PlayChart.getBubbleOpacity(d, true),
                            'stroke-width': '1px',
                            'stroke': (d: PlayChartDataPoint) => PlayChart.getStrokeFill(d, true),
                            'fill': (d: PlayChartDataPoint) => d.fill,
                            'fill-opacity': (d: PlayChartDataPoint) => d.size != null ? PlayChart.getBubbleOpacity(d, true) : 0,
                        })
                        .transition()
                        .duration(PlayChart.FrameDuration)
                        .style('opacity',1);
                    circles.exit()
                        .transition()
                        .duration(PlayChart.FrameDuration)
                        .style('opacity', 0) //fade exiting bubbles out
                        .remove();

                    TooltipManager.addTooltip(circles, (tooltipEvent: TooltipEvent) => tooltipEvent.data.tooltipInfo);

                    // sort the z-order, smallest size on top
                    circles.sort((d1: PlayChartDataPoint, d2: PlayChartDataPoint) => { return d2.size - d1.size; });
                }
                else {
                    options.svg.selectAll('.traceLine').remove();
                    options.svg.selectAll('.traceBubble').remove();
                }
            }
            else {
                options.svg.selectAll('.traceLine').remove();
                options.svg.selectAll('.traceBubble').remove();
            }

            options.data.lastRenderedFrameIndex = options.data.currentFrameIndex;
        }

        private play(): void {
            if (this.isPlaying) {
                // Toggle the flag and allow the animation logic to kill it
                this.isPlaying = false;
            }
            else if ((<any>this.slider[0]).noUiSlider != null) {
                this.isPlaying = true;
                this.playButton.removeClass('play').addClass('pause');

                let indexToShow = Math.round((<any>this.slider[0]).noUiSlider.get());
                if (indexToShow >= this.data.allDataPoints.length - 1) {
                    this.data.currentFrameIndex = -1;
                } else {
                    this.data.currentFrameIndex = indexToShow - 1;
                }

                this.playNextFrame();
            }
        }

        private persistFrameIndex(frameIndex: number): void {
            this.ridiculousFlagForPersistProperties = true;

            let properties: { [name: string]: data.SQExpr } = {};
            properties['frozenIndex'] = data.SQExprBuilder.integer(frameIndex);

            this.host.persistProperties([{
                selector: null,
                properties: properties,
                objectName: 'general'
            }]);
        }

        private playNextFrame(startFrame?: number, endFrame?: number): void {
            if (!this.isPlaying) {
                this.playComplete();
                return;
            }

            let nextFrame = this.data.currentFrameIndex + 1;
            if (startFrame != null && endFrame != null) {
                nextFrame = Math.abs(endFrame - startFrame + 1);
                startFrame = nextFrame;
            }

            if (nextFrame < this.data.allDataPoints.length && nextFrame > -1) {
                this.data.currentFrameIndex = nextFrame;
                this.data.dataPoints = this.data.allDataPoints[nextFrame];
                this.persistFrameIndex(nextFrame); //this will cause a render call
                (<any>this.slider[0]).noUiSlider.set([nextFrame]);

                if (nextFrame < this.data.allDataPoints.length) {
                    let timePerFrame = PlayChart.FrameDuration;
                    window.setTimeout(() => {
                        // Update the rangeSlider to show the correct offset
                        (<any>this.slider[0]).noUiSlider.set([nextFrame]);
                        // Play next frame
                        this.playNextFrame(startFrame, endFrame);
                    }, timePerFrame);
                }
            } else {
                this.playComplete();
            }
        }

        private playComplete(): void {
            this.playButton.removeClass('pause').addClass('play');
            this.isPlaying = false;
            this.render(false);
        }

        private drawPlayMarkers(playData: PlayChartDataPoint[], hasSelection: boolean, sizeRange: NumberRange, suppressAnimations: boolean) {
            let duration = PlayChart.FrameDuration;

            let xScale = this.xAxisProperties.scale;
            let yScale = this.yAxisProperties.scale;

            let markerTagName = PlayChart.PlayChartCircleTagName;
            let markers = this.mainGraphicsContext.selectAll(PlayChart.DotClasses.selector).data(playData, (d: PlayChartDataPoint) => d.identity.getKey());

            markers.enter().append(markerTagName)
                .classed(PlayChart.DotClasses.class, true)
                .style('opacity', 0) //fade new bubbles into visibility
                .attr('r', 0);

            markers
                .style({
                    'stroke-opacity': (d: PlayChartDataPoint) => PlayChart.getBubbleOpacity(d, hasSelection),
                    'stroke-width': '1px',
                    'stroke': (d: PlayChartDataPoint) => PlayChart.getStrokeFill(d, true),
                    'fill': (d: PlayChartDataPoint) => d.fill,
                    'fill-opacity': (d: PlayChartDataPoint) => d.size != null ? PlayChart.getBubbleOpacity(d, hasSelection) : 0,
                });

            if (suppressAnimations) {
                markers
                    .style('opacity', 1)
                    .attr({
                        r: (d: PlayChartDataPoint) => PlayChart.getBubbleRadius(d.radius, sizeRange, this.currentViewport),
                        cx: d => xScale(d.x),
                        cy: d => yScale(d.y),
                    });
            }
            else {
                markers
                    .transition()
                    .duration(duration)
                    .style('opacity', 1)
                    .attr({
                        r: (d: PlayChartDataPoint) => PlayChart.getBubbleRadius(d.radius, sizeRange, this.currentViewport),
                        cx: d => xScale(d.x),
                        cy: d => yScale(d.y),
                    });
            }

            markers
                .exit()
                .transition()
                .duration(duration)
                .style('opacity', 0) //fade out bubbles that are removed
                .attr('r', 0)
                .remove();

            return markers;
        }

        public static getStrokeFill(d: ScatterChartDataPoint, colorBorder: boolean): string {
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
                ratio = (minSize * minSize) / PlayChart.AreaOf300By300Chart;
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
                return (PlayChart.rangeContains(actualSizeDataRange, value)) ? bubblePixelAreaSizeRange.minRange : null;
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
                    projectedSize = PlayChart.project(value, actualSizeDataRange, bubblePixelAreaSizeRange);
                }

                projectedSize = Math.sqrt(projectedSize / Math.PI) * 2;
            }

            return Math.round(projectedSize);
        }

        public static rangeContains(range: DataRange, value: number): boolean {
            return range.minRange <= value && value <= range.maxRange;
        }

        public static getBubbleOpacity(d: PlayChartDataPoint, hasSelection: boolean): number {
            if (hasSelection && !d.selected) {
                return PlayChart.DimmedBubbleOpacity;
            }
            return PlayChart.DefaultBubbleOpacity;
        }

        public onClearSelection(): void {
            if (this.interactivityService)
                this.interactivityService.clearSelection();
        }

        public getSupportedCategoryAxisType(): string {
            return axisType.scalar;
        }
    }
}
