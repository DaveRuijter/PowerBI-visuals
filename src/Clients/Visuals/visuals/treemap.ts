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
    import CssConstants = jsCommon.CssConstants;
    import StringExtensions = jsCommon.StringExtensions;

    export interface TreemapConstructorOptions {
        animator: ITreemapAnimator;
        isScrollable: boolean;
        behavior?: TreemapWebBehavior;
    }

    export interface TreemapData {
        root: TreemapNode;
        hasHighlights: boolean;
        legendData: LegendData;
        dataLabelsSettings: VisualDataLabelsSettings;
        legendObjectProperties?: DataViewObject;
        dataWasCulled: boolean;
    }

    /**
     * Treemap node (we extend D3 node (GraphNode) because treemap layout methods rely on the type).
     */
    export interface TreemapNode extends D3.Layout.GraphNode, SelectableDataPoint, TooltipEnabledDataPoint, LabelEnabledDataPoint {
        key: any;
        highlightMultiplier?: number;
        color: string;
        highlightedTooltipInfo?: TooltipDataItem[];
    }

    interface TreemapRawData {
        values: number[][];
        highlights?: number[][];
        highlightsOverflow?: boolean;
        totalValue: number;
    }

    export interface ITreemapLayout {
        shapeClass: (d: TreemapNode) => string;
        shapeLayout: {
            x: (d: TreemapNode) => number;
            y: (d: TreemapNode) => number;
            width: (d: TreemapNode) => number;
            height: (d: TreemapNode) => number;
        };
        highlightShapeClass: (d: TreemapNode) => string;
        highlightShapeLayout: {
            x: (d: TreemapNode) => number;
            y: (d: TreemapNode) => number;
            width: (d: TreemapNode) => number;
            height: (d: TreemapNode) => number;
        };
        zeroShapeLayout: {
            x: (d: TreemapNode) => number;
            y: (d: TreemapNode) => number;
            width: (d: TreemapNode) => number;
            height: (d: TreemapNode) => number;
        };
        majorLabelClass: (d: TreemapNode) => string;
        majorLabelLayout: {
            x: (d: TreemapNode) => number;
            y: (d: TreemapNode) => number;
        };
        majorLabelText: (d: TreemapNode) => string;
        minorLabelClass: (d: TreemapNode) => string;
        minorLabelLayout: {
            x: (d: TreemapNode) => number;
            y: (d: TreemapNode) => number;
        };
        minorLabelText: (d: TreemapNode) => string;
        areMajorLabelsEnabled: () => boolean;
        areMinorLabelsEnabled: () => boolean;
    }

    // Todo: move to shared location
    interface DataPointObject extends DataViewObject {
        fill: Fill;
    }

    /**
     * Renders an interactive treemap visual from categorical data.
     */
    export class Treemap implements IVisual {
        public static DimmedShapeOpacity = 0.4;

        private static ClassName = 'treemap';
        public static LabelsGroupClassName = "labels";
        public static MajorLabelClassName = 'majorLabel';
        public static MinorLabelClassName = 'minorLabel';
        public static ShapesClassName = "shapes";
        public static TreemapNodeClassName = "treemapNode";
        public static RootNodeClassName = 'rootNode';
        public static ParentGroupClassName = 'parentGroup';
        public static NodeGroupClassName = 'nodeGroup';
        public static HighlightNodeClassName = 'treemapNodeHighlight';
        private static TextMargin = 5;
        private static MinorLabelTextSize = 10;
        private static MinTextWidthForMinorLabel = 18;
        private static MajorLabelTextSize = 12;
        private static MinTextWidthForMajorLabel = 22;
        private static MajorLabelTextProperties: TextProperties = {
            fontFamily: 'wf_segoe-ui_normal',
            fontSize: Treemap.MajorLabelTextSize + 'px'
        };

        /**
         * A rect with an area of 9 is a treemap rectangle of only
         * a single pixel in the middle with a 1 pixel stroke on each edge.
         */
        private static CullableArea = 9;

        private svg: D3.Selection;
        private treemap: D3.Layout.TreeMapLayout;
        private shapeGraphicsContext: D3.Selection;
        private labelGraphicsContext: D3.Selection;
        // TODO: Remove this once all visuals have implemented update.
        private currentViewport: IViewport;
        private legend: ILegend;

        private data: TreemapData;
        private style: IVisualStyle;
        private colors: IDataColorPalette;
        private element: JQuery;
        private options: VisualInitOptions;
        private isScrollable: boolean;
        private hostService: IVisualHostServices;

        /**
         * Note: Public for testing.
         */
        public animator: ITreemapAnimator;
        private interactivityService: IInteractivityService;
        private behavior: TreemapWebBehavior;

        // TODO: Remove this once all visuals have implemented update.
        private dataViews: DataView[];

        public static getLayout(labelsSettings: VisualDataLabelsSettings, alternativeScale: number): ITreemapLayout {
            let formattersCache = dataLabelUtils.createColumnFormatterCacheManager();
            let majorLabelsEnabled = labelsSettings.showCategory;
            let minorLabelsEnabled = labelsSettings.show || labelsSettings.showCategory;

            return {
                shapeClass: (d) => Treemap.getNodeClass(d, false),
                shapeLayout: Treemap.createTreemapShapeLayout(false),
                highlightShapeClass: (d) => Treemap.getNodeClass(d, true),
                highlightShapeLayout: Treemap.createTreemapShapeLayout(true),
                zeroShapeLayout: Treemap.createTreemapZeroShapeLayout(),
                majorLabelClass: (d) => Treemap.MajorLabelClassName,
                majorLabelLayout: {
                    x: (d) => d.x + Treemap.TextMargin,
                    y: (d) => d.y + Treemap.TextMargin + Treemap.MajorLabelTextSize,
                },
                majorLabelText: (d) => Treemap.createMajorLabelText(d, labelsSettings, alternativeScale, formattersCache),

                minorLabelClass: (d) => Treemap.MinorLabelClassName,
                minorLabelLayout: {
                    x: (d) => d.x + Treemap.TextMargin,
                    y: (d) => d.y + d.dy - Treemap.TextMargin,
                },
                minorLabelText: (d) => Treemap.createMinorLabelText(d, labelsSettings, alternativeScale, formattersCache),
                areMajorLabelsEnabled: () => majorLabelsEnabled,
                areMinorLabelsEnabled: () => minorLabelsEnabled,
            };
        }

        constructor(options?: TreemapConstructorOptions) {
            if (options && options.animator) {
                this.animator = options.animator;
                this.isScrollable = options.isScrollable ? options.isScrollable : false;
                this.behavior = options.behavior;
            }
        }

        public init(options: VisualInitOptions): void {
            this.options = options;
            let element = options.element;

            // Ensure viewport is empty on init
            element.empty();

            this.svg = d3.select(element.get(0))
                .append('svg')
                .style('position', 'absolute')
                .classed(Treemap.ClassName, true);
            this.shapeGraphicsContext = this.svg
                .append('g')
                .classed(Treemap.ShapesClassName, true);
            this.labelGraphicsContext = this.svg
                .append('g')
                .classed(Treemap.LabelsGroupClassName, true);

            this.element = element;

            // avoid deep copy
            this.currentViewport = {
                height: options.viewport.height,
                width: options.viewport.width,
            };

            this.style = options.style;

            this.treemap = d3.layout.treemap()
                .sticky(false)
                .sort((a, b) => a.size - b.size)
                .value((d) => d.size)
                .round(false);

            if (this.behavior) {
                this.interactivityService = createInteractivityService(options.host);
            }
            this.legend = createLegend(element, options.interactivity && options.interactivity.isInteractiveLegend, this.interactivityService, this.isScrollable);
            this.colors = this.style.colorPalette.dataColors;

            this.hostService = options.host;
        }

        /**
         * Note: Public for testing purposes.
         */
        public static converter(dataView: DataView, colors: IDataColorPalette, labelSettings: VisualDataLabelsSettings, interactivityService: IInteractivityService, viewport: IViewport, legendObjectProperties?: DataViewObject): TreemapData {
            let rootNode: TreemapNode = {
                key: "root",
                name: "root",
                children: [],
                selected: false,
                highlightMultiplier: 0,
                identity: SelectionId.createNull(),
                color: undefined,
            };
            let allNodes: TreemapNode[] = [];
            let hasHighlights: boolean;
            let legendDataPoints: LegendDataPoint[] = [];
            let legendTitle = "";
            let legendData: LegendData = { title: legendTitle, dataPoints: legendDataPoints };
            let colorHelper = new ColorHelper(colors, treemapProps.dataPoint.fill);
            let dataWasCulled = undefined;
            if (dataView && dataView.metadata && dataView.metadata.objects) {
                let objects = dataView.metadata.objects;

                labelSettings.show = DataViewObjects.getValue(objects, treemapProps.labels.show, labelSettings.show);
                labelSettings.labelColor = DataViewObjects.getFillColor(objects, treemapProps.labels.color, labelSettings.labelColor);
                labelSettings.displayUnits = DataViewObjects.getValue(objects, treemapProps.labels.labelDisplayUnits, labelSettings.displayUnits);
                labelSettings.precision = DataViewObjects.getValue(objects, treemapProps.labels.labelPrecision, labelSettings.precision);
                labelSettings.showCategory = DataViewObjects.getValue(objects, treemapProps.categoryLabels.show, labelSettings.showCategory);
            }

            if (dataView && dataView.categorical && dataView.categorical.values) {
                let data = dataView.categorical;
                let valueColumns = data.values;
                hasHighlights = !!(valueColumns.length > 0 && valueColumns[0].highlights);

                let formatStringProp = treemapProps.general.formatString;
                let result = Treemap.getValuesFromCategoricalDataView(data, hasHighlights);
                let values = result.values;
                let highlights = result.highlights;
                let totalValue = result.totalValue;
                if (result.highlightsOverflow) {
                    hasHighlights = false;
                    values = highlights;
                }

                let cullableValue = Treemap.getCullableValue(totalValue, viewport);

                let grouped = valueColumns.grouped();
                let isMultiSeries = grouped && grouped.length > 0 && grouped[0].values && grouped[0].values.length > 1;
                let hasDynamicSeries = !!valueColumns.source;
                dataWasCulled = false;
                let shouldCullValue = undefined;
                let highlight = undefined;
                if ((data.categories == null) && values) {
                    // No categories, sliced by series and measures
                    for (let i = 0, ilen = values[0].length; i < ilen; i++) {
                        let value = values[0][i];
                        if (!Treemap.checkValueForShape(value)) {
                            continue;
                        }
                        if (value < cullableValue) {
                            dataWasCulled = dataWasCulled || shouldCullValue;
                            continue;
                        }
                        let valueColumn = valueColumns[i];
                        let nodeName = converterHelper.getFormattedLegendLabel(valueColumn.source, valueColumns, formatStringProp);

                        let identity = hasDynamicSeries
                            ? SelectionId.createWithId(valueColumns[i].identity)
                            : SelectionId.createWithMeasure(valueColumns[i].source.queryName);

                        let key = identity.getKey();

                        let color = hasDynamicSeries
                            ? colorHelper.getColorForSeriesValue(valueColumn.objects && valueColumn.objects[0], data.values.identityFields, converterHelper.getSeriesName(valueColumn.source))
                            : colorHelper.getColorForMeasure(valueColumn.source.objects, valueColumn.source.queryName);

                        let highlightedValue = hasHighlights && highlight !== 0 ? highlight : undefined;
                        let categorical = dataView.categorical;
                        let valueIndex: number = i;
                        let tooltipInfo: TooltipDataItem[] = TooltipBuilder.createTooltipInfo(formatStringProp, categorical, nodeName, value, null, null, valueIndex , i);
                        let highlightedTooltipInfo: TooltipDataItem[];

                        if (highlightedValue !== undefined) {
                            highlightedTooltipInfo = TooltipBuilder.createTooltipInfo(formatStringProp, categorical, nodeName, value, null, null, valueIndex, i, highlightedValue);
                        }

                        let node: TreemapNode = {
                            key: key,
                            name: nodeName,
                            size: value,
                            color: color,
                            selected: false,
                            identity: identity,
                            tooltipInfo: tooltipInfo,
                            highlightedTooltipInfo: highlightedTooltipInfo,
                            labelFormatString: valueFormatter.getFormatString(valueColumn.source, formatStringProp),
                        };
                        if (hasHighlights && highlights) {
                            node.highlightMultiplier = value ? highlights[0][i] / value : 0;
                        }
                        rootNode.children.push(node);
                        allNodes.push(node);
                        legendDataPoints.push({
                            label: nodeName,
                            color: color,
                            icon: LegendIcon.Box,
                            identity: identity,
                            selected: false
                        });
                    }
                }
                else if (data.categories && data.categories.length > 0) {
                    // Create the first level from categories
                    let categoryColumn = data.categories[0];
                    let valueColumnCount = valueColumns.length;

                    legendTitle = categoryColumn.source ? categoryColumn.source.displayName : "";
                    let categorical = undefined;
                    for (let i = 0, ilen = values.length; i < ilen; i++) {
                        let identity: SelectionId = SelectionIdBuilder.builder()
                            .withCategory(categoryColumn, i)
                            .createSelectionId();

                        let key = identity.getKey();

                        let objects = categoryColumn.objects && categoryColumn.objects[i];

                        let color = colorHelper.getColorForSeriesValue(objects, categoryColumn.identityFields, categoryColumn.values[i]);

                        let labelFormat = valueFormatter.getFormatString(categoryColumn.values[i], formatStringProp);
                        let categoryValue = valueFormatter.format(categoryColumn.values[i], labelFormat);
                        let value = values[i][0];
                        let highlightedValue = hasHighlights && highlights ? highlights[i][0] : undefined;
                        categorical = dataView.categorical;
                        let tooltipInfo: TooltipDataItem[] = TooltipBuilder.createTooltipInfo(formatStringProp, categorical, categoryValue, value);
                        let highlightedTooltipInfo: TooltipDataItem[];

                        if (highlightedValue !== undefined) {
                            highlightedTooltipInfo = TooltipBuilder.createTooltipInfo(formatStringProp, categorical, categoryValue, value, null, null, 0, i, highlightedValue);
                        }

                        let node: TreemapNode = {
                            key: key,
                            name: categoryValue,
                            color: color,
                            selected: false,
                            identity: identity,
                            tooltipInfo: tooltipInfo,
                            highlightedTooltipInfo: highlightedTooltipInfo,
                            labelFormatString: labelFormat,
                        };

                        legendDataPoints.push({
                            label: categoryValue,
                            color: color,
                            icon: LegendIcon.Box,
                            identity: identity,
                            selected: false
                        });

                        let total = 0;
                        let highlightTotal = 0; // Used if omitting second level

                        // Do not add second level if it's one and only one data point per shape and it's not a group value
                        // e.g. Category/Series group plus only one Value field
                        let omitSecondLevel = valueColumnCount === 1 && (valueColumns[0].source.groupName == null);
                        let currentValues = values[i];

                        for (let j = 0, jlen = currentValues.length; j < jlen; j++) {
                            let valueColumn = valueColumns[j];
                            let value = currentValues[j];
                            let highlight: number;

                            shouldCullValue = value < cullableValue;
                            if (!Treemap.checkValueForShape(value) || shouldCullValue) {
                                dataWasCulled = dataWasCulled || shouldCullValue;
                                continue;
                            }

                            total += value;

                            if (hasHighlights) {
                                highlight = highlights[i][j];
                                highlightTotal += highlight;
                            }

                            if (!omitSecondLevel) {
                                let childName: string = null;

                                // Create key to ensure shape uniqueness
                                let childKey = {
                                    parentId: node.key,
                                    nodeId: undefined
                                };
                                if (isMultiSeries) {
                                    // Measure: use name and index
                                    childKey.nodeId = { name: childName, index: j };
                                    childName = valueColumn.source.displayName;
                                }
                                else {
                                    // Series group instance
                                    childKey.nodeId = valueColumn.identity.key;
                                    childName = valueColumn.source.groupName;
                                }

                                let categoricalValues = categorical ? categorical.values : null;
                                let measureId = isMultiSeries ? valueColumn.source.queryName : undefined;
                                let childIdentity = SelectionIdBuilder.builder()
                                    .withCategory(categoryColumn, i)
                                    .withSeries(categoricalValues, valueColumn)
                                    .withMeasure(measureId)
                                    .createSelectionId();

                                let highlightedValue = hasHighlights && highlight !== 0 ? highlight : undefined;
                                categorical = dataView.categorical;
                                let tooltipInfo: TooltipDataItem[] = TooltipBuilder.createTooltipInfo(formatStringProp, categorical, categoryValue, value, null, null, j, i);
                                let highlightedTooltipInfo: TooltipDataItem[];

                                if (highlightedValue !== undefined) {
                                    highlightedTooltipInfo = TooltipBuilder.createTooltipInfo(formatStringProp, categorical, categoryValue, value, null, null, j, i, highlightedValue);
                                }

                                let childNode: TreemapNode = {
                                    key: childIdentity.getKey(),
                                    name: childName,
                                    size: value,
                                    color: color,
                                    selected: false,
                                    identity: childIdentity,
                                    tooltipInfo: tooltipInfo,
                                    highlightedTooltipInfo: highlightedTooltipInfo,
                                    labelFormatString: valueFormatter.getFormatString(valueColumn.source, formatStringProp),
                                };
                                if (hasHighlights)
                                    childNode.highlightMultiplier = value ? highlight / value : 0;
                                if (node.children == null)
                                    node.children = [];

                                node.children.push(childNode);
                                allNodes.push(childNode);
                            }
                        }

                        if (Treemap.checkValueForShape(total)) {
                            node.size = total;
                            rootNode.children.push(node);
                            allNodes.push(node);
                        }
                        if (hasHighlights)
                            node.highlightMultiplier = total ? highlightTotal / total : 0;
                    }
                }
            }

            if (interactivityService) {
                interactivityService.applySelectionStateToData(allNodes);
            }

            legendData = { title: legendTitle, dataPoints: legendDataPoints };

            return { root: rootNode, hasHighlights: hasHighlights, legendData: legendData, dataLabelsSettings: labelSettings, legendObjectProperties: legendObjectProperties, dataWasCulled: dataWasCulled };
        }

        private static getValuesFromCategoricalDataView(data: DataViewCategorical, hasHighlights: boolean): TreemapRawData {
            let valueColumns = data.values;
            let categoryValueCount: number;
            if (valueColumns && (data.categories == null)) {
                categoryValueCount = 1; // We only get the first value out of each valueColumn since we don't have a category
            }
            else if (valueColumns && data.categories && data.categories.length > 0) {
                categoryValueCount = data.categories[0].values.length;
            }

            let values: number[][] = [];
            let highlights: number[][] = [];
            let totalValue = 0;
            for (let i = 0; i < categoryValueCount; i++) {
                values.push([]);
                if (hasHighlights)
                    highlights.push([]);
            }

            let highlightsOverflow: boolean;
            for (let j = 0; j < valueColumns.length; j++) {
                let valueColumn = valueColumns[j];
                for (let i = 0; i < categoryValueCount; i++) {
                    let value = valueColumn.values[i];
                    values[i].push(value);
                    totalValue += isNaN(value) ? 0 : value;
                    if (hasHighlights) {
                        let highlight = valueColumn.highlights[i];
                        if (!highlight)
                            highlight = 0;
                        highlights[i].push(highlight);
                        if (highlight > value)
                            highlightsOverflow = true;
                    }
                }
            }

            return {
                values: values,
                highlights: hasHighlights ? highlights : undefined,
                highlightsOverflow: hasHighlights ? highlightsOverflow : undefined,
                totalValue: totalValue,
            };
        }

        private static getCullableValue(totalValue: number, viewport: IViewport): number {
            let totalArea = viewport.width * viewport.height;
            let culledPercent = Treemap.CullableArea / totalArea;
            return culledPercent * totalValue;
        }

        public update(options: VisualUpdateOptions) {
            debug.assertValue(options, 'options');

            let dataViews = this.dataViews = options.dataViews;
            this.currentViewport = options.viewport;
            var dataViewCategorical = dataViews && dataViews.length > 0 && dataViews[0].categorical ? dataViews[0].categorical : undefined;
            var labelFormatString = (dataViewCategorical && dataViewCategorical.values) ? valueFormatter.getFormatString(dataViewCategorical.values[0].source, treemapProps.general.formatString) : undefined;
            var labelSettings = dataLabelUtils.getDefaultTreemapLabelSettings(labelFormatString);
            var legendObjectProperties = null;

            if (dataViewCategorical) {
                var dataView = dataViews[0];
                var dataViewMetadata = dataView.metadata;
                var objects: DataViewObjects;
                if (dataViewMetadata)
                    objects = dataViewMetadata.objects;

                if (objects) {
                    legendObjectProperties = objects['legend'];
                }

                this.data = Treemap.converter(dataView, this.colors, labelSettings, this.interactivityService, this.currentViewport, legendObjectProperties);

            }
            else {
                let rootNode: TreemapNode = {
                    key: "root",
                    name: "root",
                    children: [],
                    selected: false,
                    highlightMultiplier: 0,
                    identity: SelectionId.createNull(),
                    color: undefined,
                };
                let legendData: LegendData = { title: "", dataPoints: [] };
                let treeMapData: TreemapData = {
                    root: rootNode,
                    hasHighlights: false,
                    legendData: legendData,
                    dataLabelsSettings: labelSettings,
                    dataWasCulled: false,
                };
                this.data = treeMapData;
            }

            this.updateInternal(options.suppressAnimations);

            if (dataViews) {
                let warnings = getInvalidValueWarnings(
                    dataViews,
                    false /*supportsNaN*/,
                    false /*supportsNegativeInfinity*/,
                    false /*supportsPositiveInfinity*/);

                if (this.data.dataWasCulled) {
                    warnings.unshift(new GeometryCulledWarning());
                }

                if (warnings && warnings.length > 0)
                    this.hostService.setWarnings(warnings);
            }
        }

        // TODO: Remove this once all visuals have implemented update.
        public onDataChanged(options: VisualDataChangedOptions): void {
            this.update({
                suppressAnimations: options.suppressAnimations,
                dataViews: options.dataViews,
                viewport: this.currentViewport
            });
        }

        // TODO: Remove this once all visuals have implemented update.
        public onResizing(viewport: IViewport): void {
            this.update({
                suppressAnimations: true,
                dataViews: this.dataViews,
                viewport: viewport
            });
        }

        public onClearSelection(): void {
            if (this.interactivityService)
                this.interactivityService.clearSelection();
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
            let data = this.data;
            if (!data)
                return;

            let objectName = options.objectName,
                enumeration = new ObjectEnumerationBuilder();
            switch (objectName) {
                case 'dataPoint':
                    let dataViewCat: DataViewCategorical = this.dataViews && this.dataViews.length > 0 && this.dataViews[0] && this.dataViews[0].categorical;
                    let hasGradientRole = GradientUtils.hasGradientRole(dataViewCat);
                    if (!hasGradientRole)
                        this.enumerateDataPoints(enumeration, data);
                    break;
                case 'legend':
                    return this.enumerateLegend(data);
                case 'labels':
                    let labelSettingOptions: VisualDataLabelsSettingsOptions = {
                        enumeration: enumeration,
                        dataLabelsSettings: this.data.dataLabelsSettings,
                        show: true,
                        displayUnits: true,
                        precision: true,
                    };
                    dataLabelUtils.enumerateDataLabels(labelSettingOptions);
                    break;
                case 'categoryLabels':
                    if (this.data)
                        dataLabelUtils.enumerateCategoryLabels(enumeration, this.data.dataLabelsSettings, false /*withFill*/, false /*isDonutChart*/, true /*isTreeMap*/);
                    else
                        dataLabelUtils.enumerateCategoryLabels(enumeration, null, false /*withFill*/, false /*isDonutChart*/, true /*isTreeMap*/);
            }

            return enumeration.complete();
        }

        private enumerateDataPoints(enumeration: ObjectEnumerationBuilder, data: TreemapData): void {
            let rootChildren = data.root.children;
            if (_.isEmpty(rootChildren))
                return;

            for (let y = 0; y < rootChildren.length; y++) {
                let treemapNode = <TreemapNode>rootChildren[y];
                enumeration.pushInstance({
                    displayName: treemapNode.name,
                    selector: treemapNode.identity.getSelector(),
                    properties: {
                        fill: { solid: { color: treemapNode.color } }
                    },
                    objectName: 'dataPoint'
                });
            }
        }

        private enumerateLegend(data: TreemapData): VisualObjectInstance[] {
            let legendObjectProperties: DataViewObjects = { legend: data.legendObjectProperties };

            let show = DataViewObjects.getValue(legendObjectProperties, treemapProps.legend.show, this.legend.isVisible());
            let showTitle = DataViewObjects.getValue(legendObjectProperties, treemapProps.legend.showTitle, true);
            let titleText = DataViewObjects.getValue(legendObjectProperties, treemapProps.legend.titleText, this.data.legendData.title);

            return [{
                selector: null,
                objectName: 'legend',
                properties: {
                    show: show,
                    position: LegendPosition[this.legend.getOrientation()],
                    showTitle: showTitle,
                    titleText: titleText
                }
            }];
        }

        private static checkValueForShape(value: any): boolean {
            if (!value)
                return false;

            return value > 0;
        }

        private calculateTreemapSize(): IViewport {
            let legendMargins = this.legend.getMargins();
            return {
                height: this.currentViewport.height - legendMargins.height,
                width: this.currentViewport.width - legendMargins.width
            };
        }

        private initViewportDependantProperties(duration: number = 0): void {

            let viewport = this.calculateTreemapSize();

            this.svg.attr({
                width: viewport.width,
                height: viewport.height
            });

            SVGUtil.flushAllD3TransitionsIfNeeded(this.options);
        }

        private static hasChildrenWithIdentity(node: D3.Layout.GraphNode): boolean {
            let children = node.children;
            if (!children)
                return false;

            let count = children.length;
            if (count === 0)
                return false;

            for (let i = count - 1; i >= 0; i--) {
                if ((<TreemapNode>children[i]).identity.hasIdentity())
                    return true;
            }

            return false;
        }

        private static canDisplayMajorLabel(node: TreemapNode): boolean {
            // Only display major labels for level 1
            if (node.depth !== 1)
                return false;

            if (StringExtensions.isNullOrEmpty(node.name))
                return false;

            // Check if the room is enough for text with or without ellipse
            let availableWidth = node.dx - Treemap.TextMargin * 2;
            if (availableWidth < Treemap.MinTextWidthForMajorLabel)
                return false;
            
            // Check if the shape is high enough for label
            let textHeightWithMargin = Treemap.MajorLabelTextSize + Treemap.TextMargin * 2;
            if (node.dy < textHeightWithMargin)
                return false;

            return true;
        }

        private static canDisplayMinorLabel(node: TreemapNode, labelSettings: VisualDataLabelsSettings): boolean {
            // Only display minor labels for level 1 and 2
            if (node.depth < 1 || node.depth > 2)
                return false;

            // If a depth 1 node has children or is not showing data labels, do not show minor labels
            if (node.depth === 1 && (node.children || !labelSettings.show)) {
                return false;
            }

            if (StringExtensions.isNullOrEmpty(node.name))
                return false;
            
            // Check if the room is enough for text with or without ellipse
            let availableWidth = node.dx - Treemap.TextMargin * 2;
            if (availableWidth < Treemap.MinTextWidthForMinorLabel)
                return false;
            
            // Check if the shape is high enough for label
            let textHeightWithMargin = Treemap.MinorLabelTextSize + Treemap.TextMargin * 2;
            if (node.dy < textHeightWithMargin)
                return false;

            if (node.depth === 2) {
                let parent = node.parent;
                let roomTop = Math.max(parent.y + Treemap.MajorLabelTextSize + Treemap.TextMargin * 2, node.y);

                // Parent's label needs the room
                if (node.y + node.dy - roomTop < textHeightWithMargin)
                    return false;
            }

            return true;
        }

        private static createMajorLabelText(node: TreemapNode, labelsSettings: VisualDataLabelsSettings, alternativeScale: number, formattersCache: IColumnFormatterCacheManager): string {
            let spaceAvaliableForLabels = node.dx - Treemap.TextMargin * 2;
            let baseTextProperties = Treemap.MajorLabelTextProperties;
            let textProperties: powerbi.TextProperties = {
                text: node.name,
                fontFamily: baseTextProperties.fontFamily,
                fontSize: baseTextProperties.fontSize
            };

            return TextMeasurementService.getTailoredTextOrDefault(textProperties, spaceAvaliableForLabels);
        }

        private static createMinorLabelText(node: TreemapNode, labelsSettings: VisualDataLabelsSettings, alternativeScale: number, formattersCache: IColumnFormatterCacheManager): string {
            let spaceAvaliableForLabels = node.dx - Treemap.TextMargin * 2;
            let label = node.name;
            if (labelsSettings.show) {
                let measureFormatter = formattersCache.getOrCreate(node.labelFormatString, labelsSettings, alternativeScale);
                // Create measure label
                label = dataLabelUtils.getLabelFormattedText(node.value, spaceAvaliableForLabels, null /*format*/, measureFormatter);
                // Add category if needed (we're showing category and the node depth is 2)
                label = labelsSettings.showCategory && node.depth === 2 ? dataLabelUtils.getLabelFormattedText(node.name, spaceAvaliableForLabels) + " " + label : label;
            }

            return dataLabelUtils.getLabelFormattedText(label, spaceAvaliableForLabels);
        }

        public static getFill(d: TreemapNode, isHighlightRect: boolean): string {
            // NOTE: only painted shapes will catch click event. We either paint children or their parent but not both.

            // If it's a leaf with no category, parent will be painted instead (and support interactivity)
            if (d.depth > 1 && !d.identity.hasIdentity() && !isHighlightRect)
                return CssConstants.noneValue;

            // If it's not a leaf and it has children with a category, children will be painted
            if (Treemap.hasChildrenWithIdentity(d))
                return CssConstants.noneValue;

            return d.color;
        }

        public static getFillOpacity(d: TreemapNode, hasSelection: boolean, hasHighlights: boolean, isHighlightRect: boolean): string {
            if (hasHighlights) {
                if (isHighlightRect)
                    return null;
                return Treemap.DimmedShapeOpacity.toString();
            }

            if (!hasSelection || d.selected)
                return null;

            // Parent node is selected (as an optimization, we only check below level 1 because root node cannot be selected anyway)
            if (d.depth > 1 && (<TreemapNode>d.parent).selected)
                return null;

            // It's a parent node with interactive children, fall back to default opacity
            if (Treemap.hasChildrenWithIdentity(d))
                return null;

            return Treemap.DimmedShapeOpacity.toString();
        }

        private updateInternal(suppressAnimations: boolean): void {
            let data = this.data;
            let hasHighlights = data && data.hasHighlights;
            let labelSettings = data ? data.dataLabelsSettings : null;
            let duration = AnimatorCommon.GetAnimationDuration(this.animator, suppressAnimations);

            if (!(this.options.interactivity && this.options.interactivity.isInteractiveLegend) && this.data) {
                this.renderLegend();
            }

            this.initViewportDependantProperties(duration);
            let viewport = this.calculateTreemapSize();

            this.treemap.size([viewport.width, viewport.height]);

            // Shapes are drawn for all nodes
            let nodes = (data && data.root) ? this.treemap.nodes(data.root) : [];
            // Highlight shapes are drawn only for nodes with non-null/undefed highlightMultipliers that have no children
            let highlightNodes = nodes.filter((value: TreemapNode) => value.highlightMultiplier != null && (!value.children || value.children.length === 0));
            let majorLabeledNodes = [];
            let minorLabeledNodes = [];
            let alternativeScale: number = null;

            // Only populate major labels if category labels are turned on
            if (labelSettings.showCategory) {
                majorLabeledNodes = nodes.filter((d: TreemapNode) => Treemap.canDisplayMajorLabel(d));
            }

            // Only populate minor labels if category or data labels are turned on
            if (labelSettings.show || labelSettings.showCategory) {
                minorLabeledNodes = nodes.filter((d: TreemapNode) => Treemap.canDisplayMinorLabel(d, labelSettings));

                // If the display unit is 0 we calculate the format scale using the maximum value available
                if (labelSettings.displayUnits === 0)
                    alternativeScale = <number>d3.max(minorLabeledNodes, (d: TreemapNode) => Math.abs(d.value));
            }

            let treemapLayout = Treemap.getLayout(labelSettings, alternativeScale);
            let shapes: D3.UpdateSelection;
            let highlightShapes: D3.UpdateSelection;
            let majorLabels: D3.UpdateSelection;
            let minorLabels: D3.UpdateSelection;
            let result: TreemapAnimationResult;
            if (this.animator && !suppressAnimations) {
                let options: TreemapAnimationOptions = {
                    viewModel: data,
                    nodes: nodes,
                    highlightNodes: highlightNodes,
                    majorLabeledNodes: majorLabeledNodes,
                    minorLabeledNodes: minorLabeledNodes,
                    shapeGraphicsContext: this.shapeGraphicsContext,
                    labelGraphicsContext: this.labelGraphicsContext,
                    interactivityService: this.interactivityService,
                    layout: treemapLayout,
                    labelSettings: labelSettings,
                };
                result = this.animator.animate(options);
                shapes = result.shapes;
                highlightShapes = result.highlightShapes;
                majorLabels = result.majorLabels;
                minorLabels = result.minorLabels;
            }
            if (!this.animator || suppressAnimations || result.failed) {
                let hasSelection = this.interactivityService && this.interactivityService.hasSelection();
                let shapeGraphicsContext = this.shapeGraphicsContext;
                shapes = Treemap.drawDefaultShapes(shapeGraphicsContext, nodes, hasSelection, hasHighlights, treemapLayout);
                highlightShapes = Treemap.drawDefaultHighlightShapes(shapeGraphicsContext, highlightNodes, hasSelection, hasHighlights, treemapLayout);
                let labelGraphicsContext = this.labelGraphicsContext;
                majorLabels = Treemap.drawDefaultMajorLabels(labelGraphicsContext, majorLabeledNodes, labelSettings, treemapLayout);
                minorLabels = Treemap.drawDefaultMinorLabels(labelGraphicsContext, minorLabeledNodes, labelSettings, treemapLayout);
            }

            if (this.interactivityService) {
                let behaviorOptions: TreemapBehaviorOptions = {
                    shapes: shapes,
                    highlightShapes: highlightShapes,
                    majorLabels: majorLabels,
                    minorLabels: minorLabels,
                    nodes: <TreemapNode[]>nodes,
                    hasHighlights: data.hasHighlights,
                };

                this.interactivityService.bind(<TreemapNode[]>nodes, this.behavior, behaviorOptions);
            }

            TooltipManager.addTooltip(shapes, (tooltipEvent: TooltipEvent) => tooltipEvent.data.tooltipInfo);
            TooltipManager.addTooltip(highlightShapes, (tooltipEvent: TooltipEvent) => tooltipEvent.data.highlightedTooltipInfo);

            SVGUtil.flushAllD3TransitionsIfNeeded(this.options);
        }

        private renderLegend(): void {
            let legendObjectProperties = this.data.legendObjectProperties;
            if (legendObjectProperties) {
                let legendData = this.data.legendData;
                LegendData.update(legendData, legendObjectProperties);
                let position = <string>legendObjectProperties[legendProps.position];
                if (position)
                    this.legend.changeOrientation(LegendPosition[position]);

                this.legend.drawLegend(legendData, this.currentViewport);
            } else {
                // TODO: Draw should be the only API. Visuals should only call that with orientation, props, etc 
                // instead of managing state. Will follow up with another change.
                this.legend.changeOrientation(LegendPosition.Top);
                this.legend.drawLegend({ dataPoints: [] }, this.currentViewport);
            }
        }

        private static getNodeClass(d: TreemapNode, highlight?: boolean): string {
            let nodeClass: string;
            switch (d.depth) {
                case 1:
                    nodeClass = Treemap.ParentGroupClassName;
                    break;
                case 2:
                    nodeClass = Treemap.NodeGroupClassName;
                    break;
                case 0:
                    nodeClass = Treemap.RootNodeClassName;
                    break;
                default:
                    debug.assertFail('Treemap only supports 2 levels maxiumum');
            }
            nodeClass += " " + (highlight ? Treemap.HighlightNodeClassName : Treemap.TreemapNodeClassName);
            return nodeClass;
        }

        private static createTreemapShapeLayout(isHighlightRect = false) {
            return {
                x: (d: TreemapNode) => d.x,
                y: (d: TreemapNode) => d.y + (isHighlightRect ? d.dy * (1 - d.highlightMultiplier) : 0),
                width: (d: TreemapNode) => Math.max(0, d.dx),
                height: (d: TreemapNode) => Math.max(0, d.dy * (isHighlightRect ? d.highlightMultiplier : 1)),
            };
        }

        private static createTreemapZeroShapeLayout() {
            return {
                x: (d: TreemapNode) => d.x,
                y: (d: TreemapNode) => d.y + d.dy,
                width: (d: TreemapNode) => Math.max(0, d.dx),
                height: (d: TreemapNode) => 0,
            };
        }

        public static drawDefaultShapes(context: D3.Selection, nodes: D3.Layout.GraphNode[], hasSelection: boolean, hasHighlights: boolean, layout: ITreemapLayout): D3.UpdateSelection {
            let isHighlightShape = false;
            let shapes = context.selectAll('.' + Treemap.TreemapNodeClassName)
                .data(nodes, (d: TreemapNode) => d.key);

            shapes.enter().append('rect')
                .attr('class', layout.shapeClass);

            shapes
                .style("fill", (d: TreemapNode) => Treemap.getFill(d, isHighlightShape))
                .style("fill-opacity", (d: TreemapNode) => Treemap.getFillOpacity(d, hasSelection, hasHighlights, isHighlightShape))
                .attr(layout.shapeLayout);

            shapes.exit().remove();

            return shapes;
        }

        public static drawDefaultHighlightShapes(context: D3.Selection, nodes: D3.Layout.GraphNode[], hasSelection: boolean, hasHighlights: boolean, layout: ITreemapLayout): D3.UpdateSelection {
            let isHighlightShape = true;
            let highlightShapes = context.selectAll('.' + Treemap.HighlightNodeClassName)
                .data(nodes, (d) => d.key + "highlight");

            highlightShapes.enter().append('rect')
                .attr('class', layout.highlightShapeClass);

            highlightShapes
                .style("fill", (d: TreemapNode) => Treemap.getFill(d, isHighlightShape))
                .style("fill-opacity", (d: TreemapNode) => Treemap.getFillOpacity(d, hasSelection, hasHighlights, isHighlightShape))
                .attr(layout.highlightShapeLayout);

            highlightShapes.exit().remove();
            return highlightShapes;
        }

        public static drawDefaultMajorLabels(context: D3.Selection, nodes: D3.Layout.GraphNode[], labelSettings: VisualDataLabelsSettings, layout: ITreemapLayout): D3.UpdateSelection {
            let labels = context
                .selectAll('.' + Treemap.MajorLabelClassName)
                .data(nodes, (d: TreemapNode) => d.key);

            labels.enter().append('text')
                .attr('class', layout.majorLabelClass);

            labels
                .attr(layout.majorLabelLayout)
                .text(layout.majorLabelText)
                .style('fill', () => labelSettings.labelColor);

            labels.exit().remove();

            return labels;
        }

        public static drawDefaultMinorLabels(context: D3.Selection, nodes: D3.Layout.GraphNode[], labelSettings: VisualDataLabelsSettings, layout: ITreemapLayout): D3.UpdateSelection {
            let labels = context
                .selectAll('.' + Treemap.MinorLabelClassName)
                .data(nodes, (d: TreemapNode) => d.key);

            labels.enter().append('text')
                .attr('class', layout.minorLabelClass);

            labels
                .attr(layout.minorLabelLayout)
                .text(layout.minorLabelText)
                .style('fill', () => labelSettings.labelColor);

            labels.exit().remove();

            return labels;
        }

        public static cleanMinorLabels(context: D3.Selection) {
            let empty = [];
            let labels = context
                .selectAll('.' + Treemap.LabelsGroupClassName)
                .selectAll('.' + Treemap.MinorLabelClassName)
                .data(empty);
            labels.exit().remove();
        }
    }
}