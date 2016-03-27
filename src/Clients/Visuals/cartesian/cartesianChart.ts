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
    import EnumExtensions = jsCommon.EnumExtensions;

    var COMBOCHART_DOMAIN_OVERLAP_TRESHOLD_PERCENTAGE = 0.1;

    export const enum CartesianChartType {
        Line,
        Area,
        ClusteredColumn,
        StackedColumn,
        ClusteredBar,
        StackedBar,
        HundredPercentStackedBar,
        HundredPercentStackedColumn,
        Scatter,
        ComboChart,
        DataDot,
        Waterfall,
        LineClusteredColumnCombo,
        LineStackedColumnCombo,
        DataDotClusteredColumnCombo,
        DataDotStackedColumnCombo,
        Play,
    }

    export interface CalculateScaleAndDomainOptions {
        viewport: IViewport;
        margin: IMargin;
        forcedTickCount?: number;
        forcedYDomain?: any[];
        forcedXDomain?: any[];
        showCategoryAxisLabel: boolean;
        showValueAxisLabel: boolean;
        forceMerge: boolean;
        categoryAxisScaleType: string;
        valueAxisScaleType: string;
    }

    export interface MergedValueAxisResult {
        domain: number[];
        merged: boolean;
        tickCount: number;
        forceStartToZero: boolean;
    }

    export interface CartesianSmallViewPortProperties {
        hideLegendOnSmallViewPort: boolean;
        hideAxesOnSmallViewPort: boolean;
        MinHeightLegendVisible: number;
        MinHeightAxesVisible: number;
    }

    export interface AxisRenderingOptions {
        axisLabels: ChartAxesLabels;
        legendMargin: number;
        viewport: IViewport;
        hideXAxisTitle: boolean;
        hideYAxisTitle: boolean;
        hideY2AxisTitle?: boolean;
        xLabelColor?: Fill;
        yLabelColor?: Fill;
        y2LabelColor?: Fill;
    }

    export interface CartesianConstructorOptions {
        chartType: CartesianChartType;
        isScrollable?: boolean;
        animator?: IGenericAnimator;
        cartesianSmallViewPortProperties?: CartesianSmallViewPortProperties;
        behavior?: IInteractiveBehavior;
        seriesLabelFormattingEnabled?: boolean;
    }

    export interface ICartesianVisual {
        init(options: CartesianVisualInitOptions): void;
        setData(dataViews: DataView[]): void;
        calculateAxesProperties(options: CalculateScaleAndDomainOptions): IAxisProperties[];
        overrideXScale(xProperties: IAxisProperties): void;
        render(suppressAnimations: boolean): CartesianVisualRenderResult;
        calculateLegend(): LegendData;
        hasLegend(): boolean;
        onClearSelection(): void;
        enumerateObjectInstances?(enumeration: ObjectEnumerationBuilder, options: EnumerateVisualObjectInstancesOptions): void;
        getVisualCategoryAxisIsScalar?(): boolean;
        getSupportedCategoryAxisType?(): string;
        getPreferredPlotArea?(isScalar: boolean, categoryCount: number, categoryThickness: number): IViewport;
        setFilteredData?(startIndex: number, endIndex: number): CartesianData;
    }

    export interface CartesianVisualConstructorOptions {
        isScrollable: boolean;
        interactivityService?: IInteractivityService;
        animator?: IGenericAnimator;
        seriesLabelFormattingEnabled?: boolean;
    }

    export interface CartesianVisualRenderResult {
        dataPoints: SelectableDataPoint[];
        behaviorOptions: any;
        labelDataPoints: LabelDataPoint[];
    }

    export interface CartesianDataPoint {
        categoryValue: any;
        value: number;
        categoryIndex: number;
        seriesIndex: number;
        highlight?: boolean;
    }

    export interface CartesianSeries {
        data: CartesianDataPoint[];
    }

    export interface CartesianData {
        series: CartesianSeries[];
        categoryMetadata: DataViewMetadataColumn;
        categories: any[];
        hasHighlights?: boolean;
    }

    export interface CartesianVisualInitOptions extends VisualInitOptions {
        svg: D3.Selection;
        cartesianHost: ICartesianVisualHost;
    }

    export interface ICartesianVisualHost {
        updateLegend(data: LegendData): void;
        getSharedColors(): IDataColorPalette;
    }

    export interface ChartAxesLabels {
        x: string;
        y: string;
        y2?: string;
    }

    export const enum AxisLinesVisibility {
        ShowLinesOnXAxis = 1,
        ShowLinesOnYAxis = 2,
        ShowLinesOnBothAxis = ShowLinesOnXAxis | ShowLinesOnYAxis,
    }

    export interface CategoryLayout {
        categoryCount: number;
        categoryThickness: number;
        outerPaddingRatio: number;
        isScalar?: boolean;
    }

    export interface CategoryLayoutOptions {
        availableWidth: number;
        categoryCount: number;
        domain: any;
        isScalar?: boolean;
        isScrollable?: boolean;
    }

    export interface CartesianAxisProperties {
        x: IAxisProperties;
        y1: IAxisProperties;
        y2?: IAxisProperties;
    }

    /** 
     * Renders a data series as a cartestian visual.
     */
    export class CartesianChart implements IVisual {
        public static MinOrdinalRectThickness = 20;
        public static MinScalarRectThickness = 2;
        public static OuterPaddingRatio = 0.4;
        public static InnerPaddingRatio = 0.2;
        public static TickLabelPadding = 2;

        private static ClassName = 'cartesianChart';
        private static AxisGraphicsContextClassName = 'axisGraphicsContext';
        private static MaxMarginFactor = 0.25;
        private static MinBottomMargin = 25;
        private static TopMargin = 8;
        private static LeftPadding = 10;
        private static RightPadding = 15;
        private static BottomPadding = 12;
        private static PlayAxisBottomMargin = 75;
        private static YAxisLabelPadding = 20;
        private static XAxisLabelPadding = 18;
        private static TickPaddingY = 10;
        private static TickPaddingRotatedX = 5;
        private static FontSize = 11;
        private static FontSizeString = jsCommon.PixelConverter.toString(CartesianChart.FontSize);
        private static TextProperties: TextProperties = {
            fontFamily: 'wf_segoe-ui_normal',
            fontSize: CartesianChart.FontSizeString,
        };

        private axisGraphicsContext: D3.Selection;
        private xAxisGraphicsContext: D3.Selection;
        private y1AxisGraphicsContext: D3.Selection;
        private y2AxisGraphicsContext: D3.Selection;
        private element: JQuery;
        private svg: D3.Selection;
        private clearCatcher: D3.Selection;
        private margin: IMargin;
        private type: CartesianChartType;
        private hostServices: IVisualHostServices;
        private layers: ICartesianVisual[];
        private legend: ILegend;
        private legendMargins: IViewport;
        private layerLegendData: LegendData;
        private hasSetData: boolean;
        private visualInitOptions: VisualInitOptions;
        private legendObjectProperties: DataViewObject;
        private categoryAxisProperties: DataViewObject;
        private valueAxisProperties: DataViewObject;
        private cartesianSmallViewPortProperties: CartesianSmallViewPortProperties;
        private interactivityService: IInteractivityService;
        private behavior: IInteractiveBehavior;
        private y2AxisExists: boolean;
        private categoryAxisHasUnitType: boolean;
        private valueAxisHasUnitType: boolean;
        private hasCategoryAxis: boolean;
        private yAxisIsCategorical: boolean;
        private secValueAxisHasUnitType: boolean;
        private axes: CartesianAxisProperties;        
        private yAxisOrientation: string;
        private bottomMarginLimit: number;
        private leftRightMarginLimit: number;
        private sharedColorPalette: SharedColorPalette;
        private seriesLabelFormattingEnabled: boolean;

        public animator: IGenericAnimator;

        // Scrollbar related
        private isScrollable: boolean;
        private scrollY: boolean;
        private scrollX: boolean;
        private isXScrollBarVisible: boolean;
        private isYScrollBarVisible: boolean;
        private svgScrollable: D3.Selection;
        private axisGraphicsContextScrollable: D3.Selection;
        private labelGraphicsContextScrollable: D3.Selection;
        private brushGraphicsContext: D3.Selection;
        private brushContext: D3.Selection;
        private brush: D3.Svg.Brush;
        private static ScrollBarWidth = 10;
        private static fillOpacity = 0.125;
        private brushMinExtent: number;
        private scrollScale: D3.Scale.OrdinalScale;

        // TODO: Remove onDataChanged & onResizing once all visuals have implemented update.
        private dataViews: DataView[];
        private currentViewport: IViewport;

        private static getAxisVisibility(type: CartesianChartType): AxisLinesVisibility {
            switch (type) {
                case CartesianChartType.StackedBar:
                case CartesianChartType.ClusteredBar:
                case CartesianChartType.HundredPercentStackedBar:
                    return AxisLinesVisibility.ShowLinesOnXAxis;
                case CartesianChartType.Scatter:
                case CartesianChartType.Play:
                    return AxisLinesVisibility.ShowLinesOnBothAxis;
                default:
                    return AxisLinesVisibility.ShowLinesOnYAxis;
            }
        }

        constructor(options: CartesianConstructorOptions) {
            this.isScrollable = false;
            if (options) {
                this.type = options.chartType;
                this.seriesLabelFormattingEnabled = options.seriesLabelFormattingEnabled;
                if (options.isScrollable)
                    this.isScrollable = options.isScrollable;
                this.animator = options.animator;
                if (options.cartesianSmallViewPortProperties) {
                    this.cartesianSmallViewPortProperties = options.cartesianSmallViewPortProperties;
                }

                if (options.behavior) {
                    this.behavior = options.behavior;
            }
        }
        }

        public init(options: VisualInitOptions) {
            this.visualInitOptions = options;
            this.layers = [];

            let element = this.element = options.element;
            let viewport = this.currentViewport = options.viewport;
            this.hostServices = options.host;
            this.brush = d3.svg.brush();
            element.addClass(CartesianChart.ClassName);
            this.margin = {
                top: 1,
                right: 1,
                bottom: 1,
                left: 1
            };
            this.yAxisOrientation = yAxisPosition.left;
            this.adjustMargins(viewport);

            this.sharedColorPalette = new SharedColorPalette(options.style.colorPalette.dataColors);

            let axisLinesVisibility = CartesianChart.getAxisVisibility(this.type);

            let showLinesOnX = this.scrollY = EnumExtensions.hasFlag(axisLinesVisibility, AxisLinesVisibility.ShowLinesOnBothAxis) ||
                EnumExtensions.hasFlag(axisLinesVisibility, AxisLinesVisibility.ShowLinesOnXAxis);

            let showLinesOnY = this.scrollX = EnumExtensions.hasFlag(axisLinesVisibility, AxisLinesVisibility.ShowLinesOnBothAxis) ||
                EnumExtensions.hasFlag(axisLinesVisibility, AxisLinesVisibility.ShowLinesOnYAxis);

            /*
                The layout of the visual would look like :
                <svg>
                    <g>
                        <nonscrollable axis/>
                    </g>
                    <svgScrollable>
                        <g>
                            <scrollable axis/>
                        </g>
                    </svgScrollable>
                    <g xbrush/>
                </svg>                    

            */
            let svg = this.svg = d3.select(element.get(0)).append('svg');
            svg.style('position', 'absolute');

            let axisGraphicsContext = this.axisGraphicsContext = svg.append('g')
                .classed(CartesianChart.AxisGraphicsContextClassName, true);

            this.svgScrollable = svg.append('svg')
                .classed('svgScrollable', true)
                .style('overflow', 'hidden');

            let axisGraphicsContextScrollable = this.axisGraphicsContextScrollable = this.svgScrollable.append('g')
                .classed(CartesianChart.AxisGraphicsContextClassName, true);

            this.labelGraphicsContextScrollable = this.svgScrollable.append('g')
                .classed(NewDataLabelUtils.labelGraphicsContextClass.class, true);

            if (this.behavior)
                this.clearCatcher = appendClearCatcher(this.axisGraphicsContextScrollable);

            let axisGroup = showLinesOnX ? axisGraphicsContextScrollable : axisGraphicsContext;

            this.xAxisGraphicsContext = showLinesOnX ? axisGraphicsContext.append('g').attr('class', 'x axis') : axisGraphicsContextScrollable.append('g').attr('class', 'x axis');
            this.y1AxisGraphicsContext = axisGroup.append('g').attr('class', 'y axis');
            this.y2AxisGraphicsContext = axisGroup.append('g').attr('class', 'y axis');

            this.xAxisGraphicsContext.classed('showLinesOnAxis', showLinesOnX);
            this.y1AxisGraphicsContext.classed('showLinesOnAxis', showLinesOnY);
            this.y2AxisGraphicsContext.classed('showLinesOnAxis', showLinesOnY);

            this.xAxisGraphicsContext.classed('hideLinesOnAxis', !showLinesOnX);
            this.y1AxisGraphicsContext.classed('hideLinesOnAxis', !showLinesOnY);
            this.y2AxisGraphicsContext.classed('hideLinesOnAxis', !showLinesOnY);

            if (this.behavior) {
                this.interactivityService = createInteractivityService(this.hostServices);
            }
            this.legend = createLegend(
                element,
                options.interactivity && options.interactivity.isInteractiveLegend,
                this.type !== CartesianChartType.Waterfall ? this.interactivityService : undefined,
                this.isScrollable);
        }

        private renderAxesLabels(options: AxisRenderingOptions): void {      
            debug.assertValue(options, 'options');
            debug.assertValue(options.viewport, 'options.viewport');
            debug.assertValue(options.axisLabels, 'options.axisLabels');
                  
            this.axisGraphicsContext.selectAll('.xAxisLabel').remove();
            this.axisGraphicsContext.selectAll('.yAxisLabel').remove();

            let margin = this.margin;
            let width = options.viewport.width - (margin.left + margin.right);
            let height = options.viewport.height;
            let fontSize = CartesianChart.FontSize;
            let heightOffset = fontSize;
            if (this.type === CartesianChartType.Play && this.animator)
                heightOffset += CartesianChart.PlayAxisBottomMargin;
            let showOnRight = this.yAxisOrientation === yAxisPosition.right;

            if (!options.hideXAxisTitle) {                
                let xAxisLabel = this.axisGraphicsContext.append("text")
                    .style("text-anchor", "middle")
                    .text(options.axisLabels.x)
                    .call((text: D3.Selection) => {
                        text.each(function () {
                            let text = d3.select(this);
                            text.attr({
                                "class": "xAxisLabel",
                                "transform": SVGUtil.translate(width / 2, height - heightOffset)
                            });
                        });
                    });

                xAxisLabel.style("fill", options.xLabelColor ? options.xLabelColor.solid.color : null);

                xAxisLabel.call(AxisHelper.LabelLayoutStrategy.clip,
                    width,
                    TextMeasurementService.svgEllipsis);
            }

            if (!options.hideYAxisTitle) {
                let yAxisLabel = this.axisGraphicsContext.append("text")
                    .style("text-anchor", "middle")
                    .text(options.axisLabels.y)
                    .call((text: D3.Selection) => {
                        text.each(function () {
                            let text = d3.select(this);
                            text.attr({
                                "class": "yAxisLabel",
                                "transform": "rotate(-90)",
                                "y": showOnRight ? width + margin.right - fontSize : -margin.left,
                                "x": -((height - margin.top - options.legendMargin) / 2),
                                "dy": "1em"
                            });
                        });
                    });

                yAxisLabel.style("fill", options.yLabelColor ? options.yLabelColor.solid.color : null);

                yAxisLabel.call(AxisHelper.LabelLayoutStrategy.clip,
                    height - (margin.bottom + margin.top),
                    TextMeasurementService.svgEllipsis);
            }

            if (!options.hideY2AxisTitle && options.axisLabels.y2) {
                let y2AxisLabel = this.axisGraphicsContext.append("text")
                    .style("text-anchor", "middle")
                    .text(options.axisLabels.y2)
                    .call((text: D3.Selection) => {
                        text.each(function () {
                            let text = d3.select(this);
                            text.attr({
                                "class": "yAxisLabel",
                                "transform": "rotate(-90)",
                                "y": showOnRight ? -margin.left : width + margin.right - fontSize,
                                "x": -((height - margin.top - options.legendMargin) / 2),
                                "dy": "1em"
                            });
                        });
                    });

                y2AxisLabel.style("fill", options.y2LabelColor ? options.y2LabelColor.solid.color : null);

                y2AxisLabel.call(AxisHelper.LabelLayoutStrategy.clip,
                    height - (margin.bottom + margin.top),
                    TextMeasurementService.svgEllipsis);
            }
        }

        private adjustMargins(viewport: IViewport): void {
            let margin = this.margin;

            let width = viewport.width - (margin.left + margin.right);
            let height = viewport.height - (margin.top + margin.bottom);

            // Adjust margins if ticks are not going to be shown on either axis
            let xAxis = this.element.find('.x.axis');

            if (AxisHelper.getRecommendedNumberOfTicksForXAxis(width) === 0
                && AxisHelper.getRecommendedNumberOfTicksForYAxis(height) === 0) {
                this.margin = {
                    top: 0,
                    right: 0,
                    bottom: 0,
                    left: 0
                };
                xAxis.hide();
            } else {
                xAxis.show();
            }
        }

        // Margin convention: http://bl.ocks.org/mbostock/3019563
        private translateAxes(viewport: IViewport): void {
            this.adjustMargins(viewport);
            let margin = this.margin;

            let width = viewport.width - (margin.left + margin.right);
            let height = viewport.height - (margin.top + margin.bottom);

            let showY1OnRight = this.yAxisOrientation === yAxisPosition.right;

            this.xAxisGraphicsContext
                .attr('transform', SVGUtil.translate(0, height));

            this.y1AxisGraphicsContext
                .attr('transform', SVGUtil.translate(showY1OnRight ? width : 0, 0));

            this.y2AxisGraphicsContext
                .attr('transform', SVGUtil.translate(showY1OnRight ? 0 : width, 0));

            this.svg.attr({
                'width': viewport.width,
                'height': viewport.height
            });

            this.svgScrollable.attr({
                'width': viewport.width,
                'height': viewport.height
            });

            this.svgScrollable.attr({
                'x': 0
            });

            this.axisGraphicsContext.attr('transform', SVGUtil.translate(margin.left, margin.top));
            this.axisGraphicsContextScrollable.attr('transform', SVGUtil.translate(margin.left, margin.top));
            this.labelGraphicsContextScrollable.attr('transform', SVGUtil.translate(margin.left, margin.top));

            if (this.isXScrollBarVisible) {
                this.svgScrollable.attr({
                    'x': this.margin.left
                });
                this.axisGraphicsContextScrollable.attr('transform', SVGUtil.translate(0, margin.top));
                this.labelGraphicsContextScrollable.attr('transform', SVGUtil.translate(0, margin.top));
                this.svgScrollable.attr('width', width);
                this.svg.attr('width', viewport.width)
                    .attr('height', viewport.height + CartesianChart.ScrollBarWidth);
            }
            else if (this.isYScrollBarVisible) {
                this.svgScrollable.attr('height', height + margin.top);
                this.svg.attr('width', viewport.width + CartesianChart.ScrollBarWidth)
                    .attr('height', viewport.height);
            }
        }

        public static getIsScalar(objects: DataViewObjects, propertyId: DataViewObjectPropertyIdentifier, type: ValueType): boolean {
            let axisTypeValue = DataViewObjects.getValue(objects, propertyId);

            if (!objects || axisTypeValue === undefined) {
                // If we don't have anything set (Auto), show charts as Scalar if the category type is numeric or time. 
                // If we have the property, it will override the type.
                return !AxisHelper.isOrdinal(type);
            }

            // also checking type here to be in sync with AxisHelper, which ignores scalar if the type is non-numeric.
            return (axisTypeValue === axisType.scalar) && !AxisHelper.isOrdinal(type);
        }

        private populateObjectProperties(dataViews: DataView[]) {
            if (dataViews && dataViews.length > 0) {
                let dataViewMetadata = dataViews[0].metadata;

                if (dataViewMetadata) {
                    this.legendObjectProperties = DataViewObjects.getObject(dataViewMetadata.objects, 'legend', {});
                }
                else {
                    this.legendObjectProperties = {};
                }
                this.categoryAxisProperties = CartesianHelper.getCategoryAxisProperties(dataViewMetadata);
                this.valueAxisProperties = CartesianHelper.getValueAxisProperties(dataViewMetadata);
                let axisPosition = this.valueAxisProperties['position'];
                this.yAxisOrientation = axisPosition ? axisPosition.toString() : yAxisPosition.left;
            }
        }

        public update(options: VisualUpdateOptions) {
            debug.assertValue(options, 'options');

            let dataViews = this.dataViews = options.dataViews;
            this.currentViewport = options.viewport;

            if (!dataViews) return;

            if (this.layers.length === 0) {
                // Lazily instantiate the chart layers on the first data load.
                this.layers = this.createAndInitLayers(dataViews);

                debug.assert(this.layers.length > 0, 'createAndInitLayers should update the layers.');
            }
            let layers = this.layers;

            if (dataViews && dataViews.length > 0) {
                let warnings = getInvalidValueWarnings(
                    dataViews,
                    false /*supportsNaN*/,
                    false /*supportsNegativeInfinity*/,
                    false /*supportsPositiveInfinity*/);

                if (warnings && warnings.length > 0)
                    this.hostServices.setWarnings(warnings);

                this.populateObjectProperties(dataViews);
            }

            this.sharedColorPalette.clearPreferredScale();
            for (let i = 0, len = layers.length; i < len; i++) {
                layers[i].setData(getLayerData(dataViews, i, len));

                if (len > 1)
                    this.sharedColorPalette.rotateScale();
            }

            // Note: interactive legend shouldn't be rendered explicitly here
            // The interactive legend is being rendered in the render method of ICartesianVisual
            if (!(this.visualInitOptions.interactivity && this.visualInitOptions.interactivity.isInteractiveLegend)) {
                this.renderLegend();
            }

            this.render(!this.hasSetData || options.suppressAnimations);

            this.hasSetData = this.hasSetData || (dataViews && dataViews.length > 0);
        }

        // TODO: Remove onDataChanged & onResizing once all visuals have implemented update.
        public onDataChanged(options: VisualDataChangedOptions): void {
            this.update({
                dataViews: options.dataViews,
                suppressAnimations: options.suppressAnimations,
                viewport: this.currentViewport
            });
        }

        // TODO: Remove onDataChanged & onResizing once all visuals have implemented update.
        public onResizing(viewport: IViewport): void {
            if (this.currentViewport && (this.currentViewport.height === viewport.height && this.currentViewport.width === viewport.width)) {
                return;
            }

            this.update({
                dataViews: this.dataViews,
                suppressAnimations: true,
                viewport: viewport
            });
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
            let enumeration = new ObjectEnumerationBuilder();
            let layersLength = this.layers ? this.layers.length : 0;

            if (options.objectName === 'legend') {
                if (!this.shouldShowLegendCard())
                    return;
                let show = DataViewObject.getValue(this.legendObjectProperties, legendProps.show, this.legend.isVisible());
                let showTitle = DataViewObject.getValue(this.legendObjectProperties, legendProps.showTitle, true);
                let titleText = DataViewObject.getValue(this.legendObjectProperties, legendProps.titleText, this.layerLegendData ? this.layerLegendData.title : '');

                enumeration.pushInstance({
                    selector: null,
                    properties: {
                        show: show,
                        position: LegendPosition[this.legend.getOrientation()],
                        showTitle: showTitle,
                        titleText: titleText
                    },
                    objectName: options.objectName
                });
            }
            else if (options.objectName === 'categoryAxis' && this.hasCategoryAxis) {
                this.getCategoryAxisValues(enumeration);
            }
            else if (options.objectName === 'valueAxis') {
                this.getValueAxisValues(enumeration);
            }

            for (let i = 0, len = layersLength; i < len; i++) {
                let layer = this.layers[i];
                if (layer.enumerateObjectInstances) {
                    layer.enumerateObjectInstances(enumeration, options);
                                }
                                }

            return enumeration.complete();
                            }

        private shouldShowLegendCard(): boolean {
            let layers = this.layers;
            let dataViews = this.dataViews;

            if (layers && dataViews) {
                let layersLength = layers.length;
                let layersWithValuesCtr = 0;

                for (let i = 0; i < layersLength; i++) {
                    if (layers[i].hasLegend()) {
                        return true;
                    }

                    // if there are at least two layers with values legend card should be shown (even if each of the individual layers don't have legend)
                    let dataView = dataViews[i];
                    if (dataView && dataView.categorical && dataView.categorical.values && dataView.categorical.values.length > 0) {
                        layersWithValuesCtr++;
                        if (layersWithValuesCtr > 1) {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        public scrollTo(position: number): void {
            debug.assert(this.isXScrollBarVisible || this.isYScrollBarVisible, 'scrolling is not available');
            debug.assertValue(this.scrollScale, 'scrollScale');

            let extent = this.brush.extent();
            let extentLength = extent[1] - extent[0];
            extent[0] = this.scrollScale(position);
            extent[1] = extent[0] + extentLength;
            this.brush.extent(extent);

            let scrollSpaceLength = this.scrollScale.rangeExtent()[1];
            this.setMinBrush(scrollSpaceLength, this.brushMinExtent);

            let triggerBrush = this.brush.on('brush');
            triggerBrush(null, 0);  // We don't use the data or index.
        }

        private getCategoryAxisValues(enumeration: ObjectEnumerationBuilder): void {
            let supportedType = axisType.both;
            let isScalar = false;
            let logPossible = !!this.axes.x.isLogScaleAllowed;
            let scaleOptions = [axisScale.log, axisScale.linear];//until options can be update in propPane, show all options

            if (this.layers && this.layers[0].getSupportedCategoryAxisType) {
                supportedType = this.layers[0].getSupportedCategoryAxisType();
                if (supportedType === axisType.scalar) {
                    isScalar = true;
                }
                else {
                    isScalar = CartesianHelper.isScalar(supportedType === axisType.both, this.categoryAxisProperties);
                }
            }

            if (!isScalar) {
                if (this.categoryAxisProperties) {
                    this.categoryAxisProperties['start'] = null;
                    this.categoryAxisProperties['end'] = null;
                }
            }

            let instance: VisualObjectInstance = {
                selector: null,
                properties: {},
                objectName: 'categoryAxis',
                validValues: {
                    axisScale: scaleOptions
                }
            };

            instance.properties['show'] = this.categoryAxisProperties && this.categoryAxisProperties['show'] != null ? this.categoryAxisProperties['show'] : true;
            if (this.yAxisIsCategorical)//in case of e.g. barChart
                instance.properties['position'] = this.valueAxisProperties && this.valueAxisProperties['position'] != null ? this.valueAxisProperties['position'] : yAxisPosition.left;
            if (supportedType === axisType.both) {
                instance.properties['axisType'] = isScalar ? axisType.scalar : axisType.categorical;
            }
            if (isScalar) {
                instance.properties['axisScale'] = (this.categoryAxisProperties && this.categoryAxisProperties['axisScale'] != null && logPossible) ? this.categoryAxisProperties['axisScale'] : axisScale.linear;
                instance.properties['start'] = this.categoryAxisProperties ? this.categoryAxisProperties['start'] : null;
                instance.properties['end'] = this.categoryAxisProperties ? this.categoryAxisProperties['end'] : null;
            }
            instance.properties['showAxisTitle'] = this.categoryAxisProperties && this.categoryAxisProperties['showAxisTitle'] != null ? this.categoryAxisProperties['showAxisTitle'] : false;

            enumeration
                .pushInstance(instance)
                .pushInstance({
                selector: null,
                properties: {
                    axisStyle: this.categoryAxisProperties && this.categoryAxisProperties['axisStyle'] ? this.categoryAxisProperties['axisStyle'] : axisStyle.showTitleOnly,
                    labelColor: this.categoryAxisProperties ? this.categoryAxisProperties['labelColor'] : null
                },
                objectName: 'categoryAxis',
                validValues: {
                    axisStyle: this.categoryAxisHasUnitType ? [axisStyle.showTitleOnly, axisStyle.showUnitOnly, axisStyle.showBoth] : [axisStyle.showTitleOnly]
                }
            });
        }

        //todo: wrap all these object getters and other related stuff into an interface
        private getValueAxisValues(enumeration: ObjectEnumerationBuilder): void {
            let scaleOptions = [axisScale.log, axisScale.linear];  //until options can be update in propPane, show all options
            let logPossible = !!this.axes.y1.isLogScaleAllowed;
            let secLogPossible = this.axes.y2 != null && this.axes.y2.isLogScaleAllowed;       

            let instance: VisualObjectInstance = {
                selector: null,
                properties: {},
                objectName: 'valueAxis',
                validValues: {
                    axisScale: scaleOptions,
                    secAxisScale: scaleOptions
                }
            };

            instance.properties['show'] = this.valueAxisProperties && this.valueAxisProperties['show'] != null ? this.valueAxisProperties['show'] : true;
            
            if (!this.yAxisIsCategorical) {
                instance.properties['position'] = this.valueAxisProperties && this.valueAxisProperties['position'] != null ? this.valueAxisProperties['position'] : yAxisPosition.left;
            }
            instance.properties['axisScale'] = (this.valueAxisProperties && this.valueAxisProperties['axisScale'] != null && logPossible) ? this.valueAxisProperties['axisScale'] : axisScale.linear;
            instance.properties['start'] = this.valueAxisProperties ? this.valueAxisProperties['start'] : null;
            instance.properties['end'] = this.valueAxisProperties ? this.valueAxisProperties['end'] : null;
            instance.properties['showAxisTitle'] = this.valueAxisProperties && this.valueAxisProperties['showAxisTitle'] != null ? this.valueAxisProperties['showAxisTitle'] : false;

            enumeration
                .pushInstance(instance)
                .pushInstance({
                selector: null,
                properties: {
                    axisStyle: this.valueAxisProperties && this.valueAxisProperties['axisStyle'] != null ? this.valueAxisProperties['axisStyle'] : axisStyle.showTitleOnly,
                    labelColor: this.valueAxisProperties ? this.valueAxisProperties['labelColor'] : null
                },
                objectName: 'valueAxis',
                validValues: {
                    axisStyle: this.valueAxisHasUnitType ? [axisStyle.showTitleOnly, axisStyle.showUnitOnly, axisStyle.showBoth] : [axisStyle.showTitleOnly]
                },
            });

            if (this.layers.length === 2) {
                instance.properties['secShow'] = this.valueAxisProperties && this.valueAxisProperties['secShow'] != null ? this.valueAxisProperties['secShow'] : this.y2AxisExists;
                if (instance.properties['secShow']) {
                    instance.properties['axisLabel'] = '';//this.layers[0].getVisualType();//I will keep or remove this, depending on the decision made
                }
            }

            if (this.y2AxisExists && instance.properties['secShow']) {
                enumeration.pushContainer({
                    displayName: data.createDisplayNameGetter('Visual_YAxis_ShowSecondary'),
                    expander: comboChartProps.valueAxis.secShow,
                });

                let secInstance: VisualObjectInstance = {
                    selector: null,
                    properties: {},
                    objectName: 'valueAxis'
                };
                secInstance.properties['secAxisLabel'] = ''; //this.layers[1].getVisualType(); //I will keep or remove this, depending on the decision made                        
                secInstance.properties['secPosition'] = this.valueAxisProperties && this.valueAxisProperties['secPosition'] != null ? this.valueAxisProperties['secPosition'] : yAxisPosition.right;
                secInstance.properties['secAxisScale'] = this.valueAxisProperties && this.valueAxisProperties['secAxisScale'] != null && secLogPossible? this.valueAxisProperties['secAxisScale'] : axisScale.linear;                
                secInstance.properties['secStart'] = this.valueAxisProperties ? this.valueAxisProperties['secStart'] : null;
                secInstance.properties['secEnd'] = this.valueAxisProperties ? this.valueAxisProperties['secEnd'] : null;
                secInstance.properties['secShowAxisTitle'] = this.valueAxisProperties && this.valueAxisProperties['secShowAxisTitle'] != null ? this.valueAxisProperties['secShowAxisTitle'] : false;

                enumeration
                    .pushInstance(secInstance)
                    .pushInstance({
                    selector: null,
                    properties: {
                        secAxisStyle: this.valueAxisProperties && this.valueAxisProperties['secAxisStyle'] ? this.valueAxisProperties['secAxisStyle'] : axisStyle.showTitleOnly,
                        labelColor: this.valueAxisProperties ? this.valueAxisProperties['secLabelColor'] : null
                    },
                    objectName: 'valueAxis',
                    validValues: {
                        secAxisStyle: this.secValueAxisHasUnitType ? [axisStyle.showTitleOnly, axisStyle.showUnitOnly, axisStyle.showBoth] : [axisStyle.showTitleOnly]
                    },
                });

                enumeration.popContainer();
            }
        }

        public onClearSelection(): void {
            if (this.hasSetData) {
                for (let i = 0, len = this.layers.length; i < len; i++) {
                    let layer = this.layers[i];
                    layer.onClearSelection();
                    layer.render(true /* suppressAnimations */);
                }
            }
        }

        private createAndInitLayers(dataViews: DataView[]): ICartesianVisual[] {
            let objects: DataViewObjects;
            if (dataViews && dataViews.length > 0) {
                let dataViewMetadata = dataViews[0].metadata;
                if (dataViewMetadata)
                    objects = dataViewMetadata.objects;
            }
         
            // Create the layers
            let layers = CartesianLayerFactory.createLayers(this.type, objects, this.interactivityService, this.animator, this.isScrollable, this.seriesLabelFormattingEnabled);

            // Initialize the layers
            let cartesianOptions = <CartesianVisualInitOptions>Prototype.inherit(this.visualInitOptions);
            cartesianOptions.svg = this.axisGraphicsContextScrollable;
            cartesianOptions.cartesianHost = {
                updateLegend: data => this.legend.drawLegend(data, this.currentViewport),
                getSharedColors: () => this.sharedColorPalette,
            };

            for (let i = 0, len = layers.length; i < len; i++)
                layers[i].init(cartesianOptions);

            return layers;
        }

        private renderLegend(): void {
            let layers = this.layers;
            let legendData: LegendData = { title: "", dataPoints: [] };

            for (let i = 0, len = layers.length; i < len; i++) {
                this.layerLegendData = layers[i].calculateLegend();
                if (this.layerLegendData) {
                    legendData.title = i === 0 ? this.layerLegendData.title || ""
                        : legendData.title;
                    legendData.dataPoints = legendData.dataPoints.concat(this.layerLegendData.dataPoints || []);
                    if (this.layerLegendData.grouped) {
                        legendData.grouped = true;
                    }
                }
            }

            let legendProperties = this.legendObjectProperties;

            if (legendProperties) {
                LegendData.update(legendData, legendProperties);
                let position = <string>legendProperties[legendProps.position];

                if (position)
                    this.legend.changeOrientation(LegendPosition[position]);
            }
            else {
                this.legend.changeOrientation(LegendPosition.Top);
            }

            if ((legendData.dataPoints.length === 1 && !legendData.grouped) || this.hideLegends()) {
                legendData.dataPoints = [];
            }

            this.legend.drawLegend(legendData, this.currentViewport);
        }

        private hideLegends(): boolean {
            if (this.cartesianSmallViewPortProperties) {
                if (this.cartesianSmallViewPortProperties.hideLegendOnSmallViewPort && (this.currentViewport.height < this.cartesianSmallViewPortProperties.MinHeightLegendVisible)) {
                    return true;
                }
            }
            return false;
        }

        private addUnitTypeToAxisLabel(axes: CartesianAxisProperties): void {
            let unitType = CartesianChart.getUnitType(axes, (axis: CartesianAxisProperties): IAxisProperties => axis.x);
            if (axes.x.isCategoryAxis) {
                this.categoryAxisHasUnitType = unitType !== null;
            }
            else {
                this.valueAxisHasUnitType = unitType !== null;
            }

            if (axes.x.axisLabel && unitType) {
                if (axes.x.isCategoryAxis) {
                    axes.x.axisLabel = AxisHelper.createAxisLabel(this.categoryAxisProperties, axes.x.axisLabel, unitType);
                }
                else {
                    axes.x.axisLabel = AxisHelper.createAxisLabel(this.valueAxisProperties, axes.x.axisLabel, unitType);
                }
            }

            unitType = CartesianChart.getUnitType(axes, (axis: CartesianAxisProperties): IAxisProperties => axis.y1);

            if (!axes.y1.isCategoryAxis) {
                this.valueAxisHasUnitType = unitType !== null;
            }
            else {
                this.categoryAxisHasUnitType = unitType !== null;
            }

            if (axes.y1.axisLabel && unitType) {
                if (!axes.y1.isCategoryAxis) {
                    axes.y1.axisLabel = AxisHelper.createAxisLabel(this.valueAxisProperties, axes.y1.axisLabel, unitType);
                }
                else {
                    axes.y1.axisLabel = AxisHelper.createAxisLabel(this.categoryAxisProperties, axes.y1.axisLabel, unitType);
                }
            }

            if (axes.y2) {
                let unitType = CartesianChart.getUnitType(axes, (axis: CartesianAxisProperties): IAxisProperties => axis.y2);
                this.secValueAxisHasUnitType = unitType !== null;
                if (axes.y2.axisLabel && unitType) {
                    if (this.valueAxisProperties && this.valueAxisProperties['secAxisStyle']) {
                        if (this.valueAxisProperties['secAxisStyle'] === axisStyle.showBoth) {
                            axes.y2.axisLabel = axes.y2.axisLabel + ' (' + unitType + ')';
                        }
                        else if (this.valueAxisProperties['secAxisStyle'] === axisStyle.showUnitOnly) {
                            axes.y2.axisLabel = unitType;
                        }
                    }
                }
            }
        }

        private shouldRenderSecondaryAxis(axisProperties: IAxisProperties): boolean {
            if (!axisProperties) {
                return false;
            }
            if (!this.valueAxisProperties || this.valueAxisProperties["secShow"] == null || this.valueAxisProperties["secShow"]) {
                return axisProperties.values && axisProperties.values.length > 0;
            }

            return false;
        }

        private shouldRenderAxis(axisProperties: IAxisProperties, propertyName: string = "show"): boolean {
            if (!axisProperties) {
                return false;
            }
            else if (axisProperties.isCategoryAxis && (!this.categoryAxisProperties || this.categoryAxisProperties[propertyName] == null || this.categoryAxisProperties[propertyName])) {
                return axisProperties.values && axisProperties.values.length > 0;
            }
            else if (!axisProperties.isCategoryAxis && (!this.valueAxisProperties || this.valueAxisProperties[propertyName] == null || this.valueAxisProperties[propertyName])) {
                return axisProperties.values && axisProperties.values.length > 0;
            }

            return false;
        }

        private render(suppressAnimations: boolean): void {
            let legendMargins = this.legendMargins = this.legend.getMargins();
            let viewport: IViewport = {
                height: this.currentViewport.height - legendMargins.height,
                width: this.currentViewport.width - legendMargins.width
            };

            let maxMarginFactor = this.getMaxMarginFactor();
            let leftRightMarginLimit = this.leftRightMarginLimit = viewport.width * maxMarginFactor;
            let bottomMarginLimit = this.bottomMarginLimit = Math.max(CartesianChart.MinBottomMargin, Math.ceil(viewport.height * maxMarginFactor));

            let margin = this.margin;
            // reset defaults
            margin.top = CartesianChart.TopMargin;
            margin.bottom = CartesianChart.MinBottomMargin;
            margin.right = 0;
            if (this.type === CartesianChartType.Play && this.animator)
                margin.bottom += CartesianChart.PlayAxisBottomMargin;

            let axes = this.axes = calculateAxes(this.layers, viewport, margin, this.categoryAxisProperties, this.valueAxisProperties, CartesianChart.TextProperties, this.isXScrollBarVisible || this.isYScrollBarVisible, null);

            this.y2AxisExists = axes.y2 != null;
            this.yAxisIsCategorical = axes.y1.isCategoryAxis;
            this.hasCategoryAxis = this.yAxisIsCategorical ? axes.y1 && axes.y1.values.length > 0 : axes.x && axes.x.values.length > 0;

            let renderXAxis = this.shouldRenderAxis(axes.x);
            let renderY1Axis = this.shouldRenderAxis(axes.y1);
            let renderY2Axis = this.shouldRenderSecondaryAxis(axes.y2);

            let width = viewport.width - (margin.left + margin.right);
            let isScalar = false;
            let mainAxisScale;
            let preferredViewport: IViewport;
            this.isXScrollBarVisible = false;
            this.isYScrollBarVisible = false;

            let yAxisOrientation = this.yAxisOrientation;
            let showY1OnRight = yAxisOrientation === yAxisPosition.right;

            if (this.layers) {
                if (this.layers[0].getVisualCategoryAxisIsScalar)
                    isScalar = this.layers[0].getVisualCategoryAxisIsScalar();

                if (!isScalar && this.isScrollable && this.layers[0].getPreferredPlotArea) {
                    let categoryThickness = this.scrollX ? axes.x.categoryThickness : axes.y1.categoryThickness;
                    let categoryCount = this.scrollX ? axes.x.values.length : axes.y1.values.length;
                    preferredViewport = this.layers[0].getPreferredPlotArea(isScalar, categoryCount, categoryThickness);
                    if (this.scrollX && preferredViewport && preferredViewport.width > viewport.width) {
                        this.isXScrollBarVisible = true;
                        viewport.height -= CartesianChart.ScrollBarWidth;
                    }

                    if (this.scrollY && preferredViewport && preferredViewport.height > viewport.height) {
                        this.isYScrollBarVisible = true;
                        viewport.width -= CartesianChart.ScrollBarWidth;
                        width = viewport.width - (margin.left + margin.right);
                    }
                }
            }

            // Only create the g tag where there is a scrollbar
            if (this.isXScrollBarVisible || this.isYScrollBarVisible) {
                if (!this.brushGraphicsContext) {
                    this.brushGraphicsContext = this.svg.append("g")
                        .classed('x brush', true);
                }
            }
            else {
                // clear any existing brush if no scrollbar is shown
                this.svg.selectAll('.brush').remove();
                this.brushGraphicsContext = undefined;
            }

            // Recalculate axes now that scrollbar visible variables have been set
            axes = calculateAxes(this.layers, viewport, margin, this.categoryAxisProperties, this.valueAxisProperties, CartesianChart.TextProperties, this.isXScrollBarVisible || this.isYScrollBarVisible, null);

            // we need to make two passes because the margin changes affect the chosen tick values, which then affect the margins again.
            // after the second pass the margins are correct.
            let doneWithMargins = false,
                maxIterations = 2,
                numIterations = 0;
            let tickLabelMargins = undefined;
            let chartHasAxisLabels = undefined;
            let axisLabels: ChartAxesLabels = undefined;
            while (!doneWithMargins && numIterations < maxIterations) {
                numIterations++;
                tickLabelMargins = AxisHelper.getTickLabelMargins(
                    { width: width, height: viewport.height },
                    leftRightMarginLimit,
                    TextMeasurementService.measureSvgTextWidth,
                    TextMeasurementService.estimateSvgTextHeight,
                    axes,
                    bottomMarginLimit,
                    CartesianChart.TextProperties,
                    this.isXScrollBarVisible || this.isYScrollBarVisible,
                    showY1OnRight,
                    renderXAxis,
                    renderY1Axis,
                    renderY2Axis);

                // We look at the y axes as main and second sides, if the y axis orientation is right so the main side represents the right side
                let maxMainYaxisSide = showY1OnRight ? tickLabelMargins.yRight : tickLabelMargins.yLeft,
                    maxSecondYaxisSide = showY1OnRight ? tickLabelMargins.yLeft : tickLabelMargins.yRight,
                    xMax = tickLabelMargins.xMax;
                // TODO: there is a better way, the visual should communicate that it needs space below the x-axis through ICartesianVisual
                if (this.type === CartesianChartType.Play && this.animator)
                    xMax += CartesianChart.PlayAxisBottomMargin;

                maxMainYaxisSide += CartesianChart.LeftPadding;
                if ((renderY2Axis && !showY1OnRight) || (showY1OnRight && renderY1Axis))
                    maxSecondYaxisSide += CartesianChart.RightPadding;
                xMax += CartesianChart.BottomPadding;

                if (this.hideAxisLabels(legendMargins)) {
                    axes.x.axisLabel = null;
                    axes.y1.axisLabel = null;
                    if (axes.y2) {
                        axes.y2.axisLabel = null;
                    }
                }

                this.addUnitTypeToAxisLabel(axes);

                axisLabels = { x: axes.x.axisLabel, y: axes.y1.axisLabel, y2: axes.y2 ? axes.y2.axisLabel : null };
                chartHasAxisLabels = (axisLabels.x != null) || (axisLabels.y != null || axisLabels.y2 != null);

                if (axisLabels.x != null)
                    xMax += CartesianChart.XAxisLabelPadding;
                if (axisLabels.y != null)
                    maxMainYaxisSide += CartesianChart.YAxisLabelPadding;
                if (axisLabels.y2 != null)
                    maxSecondYaxisSide += CartesianChart.YAxisLabelPadding;
                
                margin.left = showY1OnRight ? maxSecondYaxisSide : maxMainYaxisSide;
                margin.right = showY1OnRight ? maxMainYaxisSide : maxSecondYaxisSide;
                margin.bottom = xMax;
                this.margin = margin;

                width = viewport.width - (margin.left + margin.right);

                // re-calculate the axes with the new margins
				let previousTickCountY1 = axes.y1.values.length;
                let previousTickCountY2 = axes.y2 && axes.y2.values.length;
                axes = calculateAxes(this.layers, viewport, margin, this.categoryAxisProperties, this.valueAxisProperties, CartesianChart.TextProperties, this.isXScrollBarVisible || this.isYScrollBarVisible, axes);

                // the minor padding adjustments could have affected the chosen tick values, which would then need to calculate margins again
                // e.g. [0,2,4,6,8] vs. [0,5,10] the 10 is wider and needs more margin.
                if (axes.y1.values.length === previousTickCountY1 && (!axes.y2 || axes.y2.values.length === previousTickCountY2))
                    doneWithMargins = true;
            }

            if (this.isXScrollBarVisible) {
                mainAxisScale = axes.x.scale;
                let brushX = this.margin.left;
                let brushY = viewport.height;
                this.renderChartWithScrollBar(mainAxisScale, brushX, brushY, preferredViewport.width, viewport, axes, width, tickLabelMargins, chartHasAxisLabels, axisLabels, suppressAnimations);
            }
            else if (this.isYScrollBarVisible) {
                mainAxisScale = axes.y1.scale;
                let brushX = viewport.width;
                let brushY = this.margin.top;
                this.renderChartWithScrollBar(mainAxisScale, brushX, brushY, preferredViewport.height, viewport, axes, width, tickLabelMargins, chartHasAxisLabels, axisLabels, suppressAnimations);
            }
            else {
                this.renderChart(mainAxisScale, axes, width, tickLabelMargins, chartHasAxisLabels, axisLabels, viewport, suppressAnimations);
            }
        }

        private hideAxisLabels(legendMargins: IViewport): boolean {
            if (this.cartesianSmallViewPortProperties) {
                if (this.cartesianSmallViewPortProperties.hideAxesOnSmallViewPort && ((this.currentViewport.height + legendMargins.height) < this.cartesianSmallViewPortProperties.MinHeightAxesVisible) && !this.visualInitOptions.interactivity.isInteractiveLegend) {
                    return true;
                }
            }
            return false;
        }

        private renderChartWithScrollBar(
            inputMainAxisScale: D3.Scale.GenericScale<any>,
            brushX: number,
            brushY: number,
            svgLength: number,
            viewport: IViewport,
            axes: CartesianAxisProperties,
            width: number,
            tickLabelMargins: any,
            chartHasAxisLabels: boolean,
            axisLabels: ChartAxesLabels,
            suppressAnimations: boolean): void {

            let mainAxisScale = <D3.Scale.OrdinalScale>inputMainAxisScale;
            let scrollScale = this.scrollScale = <D3.Scale.OrdinalScale>mainAxisScale.copy();
            let brush = this.brush;
            let scrollSpaceLength;
            let marginTop = this.margin.top;
            let marginLeft = this.margin.left;
            let marginRight = this.margin.right;
            let marginBottom = this.margin.bottom;
            let minExtent;

            if (this.isXScrollBarVisible) {
                scrollSpaceLength = viewport.width - (marginLeft + marginRight);
                minExtent = this.getMinExtent(svgLength, scrollSpaceLength);
                scrollScale.rangeBands([0, scrollSpaceLength]);
                brush.x(scrollScale)
                    .extent([0, minExtent]);
            }
            else {
                scrollSpaceLength = viewport.height - (marginTop + marginBottom);
                minExtent = this.getMinExtent(svgLength, scrollSpaceLength);
                scrollScale.rangeBands([0, scrollSpaceLength]);
                brush.y(scrollScale)
                    .extent([0, minExtent]);
            }

            this.brushMinExtent = minExtent;

            brush
                .on("brush", () => window.requestAnimationFrame(() => this.onBrushed(scrollScale, mainAxisScale, axes, width, tickLabelMargins, chartHasAxisLabels, axisLabels, viewport, scrollSpaceLength)))
                .on("brushend", () => this.onBrushEnd(minExtent));

            let brushContext = this.brushContext = this.brushGraphicsContext
                .attr({
                    "transform": SVGUtil.translate(brushX, brushY),
                    "drag-resize-disabled": "true" /*disables resizing of the visual when dragging the scrollbar in edit mode*/
                })
                .call(brush);  /*call the brush function, causing it to create the rectangles   */              
              
            /* Disabling the zooming feature */
            brushContext.selectAll(".resize rect")
                .remove();

            brushContext.select(".background")
                .style('cursor', 'default');

            brushContext.selectAll(".extent")
                .style({
                    "fill-opacity": CartesianChart.fillOpacity,
                    "cursor": "default",
                });

            if (this.isXScrollBarVisible)
                brushContext.selectAll("rect").attr("height", CartesianChart.ScrollBarWidth);
            else
                brushContext.selectAll("rect").attr("width", CartesianChart.ScrollBarWidth);

            if (mainAxisScale && scrollScale) {
                mainAxisScale.rangeBands([0, scrollSpaceLength]);
                this.renderChart(mainAxisScale, axes, width, tickLabelMargins, chartHasAxisLabels, axisLabels, viewport, suppressAnimations, scrollScale, brush.extent());
            }
        }

        private getMinExtent(svgLength: number, scrollSpaceLength: number): number {
            return scrollSpaceLength * scrollSpaceLength / (svgLength);
        }

        private onBrushEnd(minExtent: number): void {
            let brushContext = this.brushContext;
            if (this.isXScrollBarVisible) {
                brushContext.select(".extent").attr("width", minExtent);
            }
            else
                brushContext.select(".extent").attr("height", minExtent);
        }

        private onBrushed(scrollScale: any, mainAxisScale: any, axes: CartesianAxisProperties, width: number, tickLabelMargins: any, chartHasAxisLabels: boolean, axisLabels: ChartAxesLabels, viewport: IViewport, scrollSpaceLength: number): void {
            let brush = this.brush;

            if (mainAxisScale && scrollScale) {
                CartesianChart.clampBrushExtent(this.brush, scrollSpaceLength, this.brushMinExtent);
                let extent = brush.extent();
                this.renderChart(mainAxisScale, axes, width, tickLabelMargins, chartHasAxisLabels, axisLabels, viewport, true /* suppressAnimations */, scrollScale, extent);
            }
        }
        
        /**
         * To show brush every time when mouse is clicked on the empty background.
         */
        private setMinBrush(scrollSpaceLength: number, minExtent: number): void {
            CartesianChart.clampBrushExtent(this.brush, scrollSpaceLength, minExtent);
        }

        private static getUnitType(axis: CartesianAxisProperties, axisPropertiesLookup: (axis: CartesianAxisProperties) => IAxisProperties) {
            if (axisPropertiesLookup(axis).formatter &&
                axisPropertiesLookup(axis).formatter.displayUnit &&
                axisPropertiesLookup(axis).formatter.displayUnit.value > 1)
                return axisPropertiesLookup(axis).formatter.displayUnit.title;
            return null;
        }

        private static clampBrushExtent(brush: D3.Svg.Brush, viewportWidth: number, minExtent: number): void {
            let extent = brush.extent();
            let width = extent[1] - extent[0];

            if (width === minExtent && extent[1] <= viewportWidth && extent[0] >= 0)
                return;

            if (width > minExtent) {
                let padding = (width - minExtent) / 2;
                extent[0] += padding;
                extent[1] -= padding;
            }

            else if (width < minExtent) {
                let padding = (minExtent - width) / 2;
                extent[0] -= padding;
                extent[1] += padding;
            }

            if (extent[0] < 0) {
                extent[0] = 0;
                extent[1] = minExtent;
            }

            else if (extent[0] > viewportWidth - minExtent) {
                extent[0] = viewportWidth - minExtent;
                extent[1] = viewportWidth;
            }

            brush.extent(extent);
        }

        private getMaxMarginFactor(): number {
            return this.visualInitOptions.style.maxMarginFactor || CartesianChart.MaxMarginFactor;
        }

        private static getChartViewport(viewport: IViewport, margin: IMargin): IViewport {
            return {
                width: viewport.width - margin.left - margin.right,
                height: viewport.height - margin.top - margin.bottom,
            };
        }

        private renderChart(
            mainAxisScale: any,
            axes: CartesianAxisProperties,
            width: number,
            tickLabelMargins: any,
            chartHasAxisLabels: boolean,
            axisLabels: ChartAxesLabels,
            viewport: IViewport,
            suppressAnimations: boolean,
            scrollScale?: any,
            extent?: number[]) {

            let bottomMarginLimit = this.bottomMarginLimit;
            let leftRightMarginLimit = this.leftRightMarginLimit;
            let layers = this.layers;
            let duration = AnimatorCommon.GetAnimationDuration(this.animator, suppressAnimations);
            let chartViewport = CartesianChart.getChartViewport(viewport, this.margin);

            debug.assertValue(layers, 'layers');

            // Filter data that fits viewport
            if (scrollScale) {
                let selected: number[];
                let data: CartesianData[] = [];

                let startValue = extent[0];
                let endValue = extent[1];

                let pixelStepSize = scrollScale(1) - scrollScale(0);
                let startIndex = Math.floor(startValue / pixelStepSize);
                let sliceLength = Math.ceil((endValue - startValue) / pixelStepSize);
                let endIndex = startIndex + sliceLength; //intentionally one past the end index for use with slice(start,end)
                let domain = scrollScale.domain();

                mainAxisScale.domain(domain);
                selected = domain.slice(startIndex, endIndex); //up to but not including 'end'
                if (selected && selected.length > 0) {
                    for (let i = 0; i < layers.length; i++) {
                        data[i] = layers[i].setFilteredData(selected[0], selected[selected.length - 1] + 1);
                    }
                    mainAxisScale.domain(selected);

                    let axisPropsToUpdate: IAxisProperties;
                    if (this.isXScrollBarVisible) {
                        axisPropsToUpdate = axes.x;
                    }
                    else {
                        axisPropsToUpdate = axes.y1;
                    }

                    axisPropsToUpdate.axis.scale(mainAxisScale);
                    axisPropsToUpdate.scale(mainAxisScale);

                    // tick values are indices for ordinal axes
                    axisPropsToUpdate.axis.ticks(selected.length);
                    axisPropsToUpdate.axis.tickValues(selected); 

                    // use the original tick format to format the tick values
                    let tickFormat = axisPropsToUpdate.axis.tickFormat();
                    axisPropsToUpdate.values = _.map(selected, (d) => tickFormat(d));
                }
            }

            let xLabelColor:Fill;
            let yLabelColor:Fill;
            let y2LabelColor:Fill;
            //hide show x-axis here
            if (this.shouldRenderAxis(axes.x)) {
                if (axes.x.isCategoryAxis) {
                    xLabelColor = this.categoryAxisProperties && this.categoryAxisProperties['labelColor'] ? this.categoryAxisProperties['labelColor'] : null;
                } else {
                    xLabelColor = this.valueAxisProperties && this.valueAxisProperties['labelColor'] ? this.valueAxisProperties['labelColor'] : null;
                }
                axes.x.axis.orient("bottom");
                if (!axes.x.willLabelsFit)
                    axes.x.axis.tickPadding(CartesianChart.TickPaddingRotatedX);

                let xAxisGraphicsElement = this.xAxisGraphicsContext;
                if (duration) {
                    xAxisGraphicsElement
                        .transition()
                        .duration(duration)
                        .call(axes.x.axis);
                }
                else {
                    xAxisGraphicsElement
                        .call(axes.x.axis);
                }

                xAxisGraphicsElement
                    .call(CartesianChart.darkenZeroLine)
                    .call(CartesianChart.setAxisLabelColor, xLabelColor);

                let xAxisTextNodes = xAxisGraphicsElement.selectAll('text');
                if (axes.x.willLabelsWordBreak) {
                    xAxisTextNodes
                        .call(AxisHelper.LabelLayoutStrategy.wordBreak, axes.x, bottomMarginLimit);
                } else {
                    xAxisTextNodes
                        .call(AxisHelper.LabelLayoutStrategy.rotate,
                            bottomMarginLimit,
                            TextMeasurementService.svgEllipsis,
                            !axes.x.willLabelsFit,
                            bottomMarginLimit === tickLabelMargins.xMax,
                            axes.x,
                            this.margin,
                            this.isXScrollBarVisible || this.isYScrollBarVisible);
                }
            }
            else {
                this.xAxisGraphicsContext.selectAll('*').remove();
            }

            if (this.shouldRenderAxis(axes.y1)) {
                if (axes.y1.isCategoryAxis) {
                    yLabelColor = this.categoryAxisProperties && this.categoryAxisProperties['labelColor'] ? this.categoryAxisProperties['labelColor'] : null;
                } else {
                    yLabelColor = this.valueAxisProperties && this.valueAxisProperties['labelColor'] ? this.valueAxisProperties['labelColor'] : null;
                }
                let yAxisOrientation = this.yAxisOrientation;
                let showY1OnRight = yAxisOrientation === yAxisPosition.right;
                axes.y1.axis
                    .tickSize(-width)
                    .tickPadding(CartesianChart.TickPaddingY)
                    .orient(yAxisOrientation.toLowerCase());

                let y1AxisGraphicsElement = this.y1AxisGraphicsContext;
                if (duration) {
                    y1AxisGraphicsElement
                        .transition()
                        .duration(duration)
                        .call(axes.y1.axis);
                }
                else {
                    y1AxisGraphicsElement
                        .call(axes.y1.axis);
                }

                y1AxisGraphicsElement
                    .call(CartesianChart.darkenZeroLine)
                    .call(CartesianChart.setAxisLabelColor, yLabelColor);

                if (tickLabelMargins.yLeft >= leftRightMarginLimit) {
                    y1AxisGraphicsElement.selectAll('text')
                        .call(AxisHelper.LabelLayoutStrategy.clip,
                            // Can't use padding space to render text, so subtract that from available space for ellipses calculations
                            leftRightMarginLimit - CartesianChart.LeftPadding,
                            TextMeasurementService.svgEllipsis);
                }

                if (axes.y2 && (!this.valueAxisProperties || this.valueAxisProperties['secShow'] == null || this.valueAxisProperties['secShow'])) {
                    y2LabelColor = this.valueAxisProperties && this.valueAxisProperties['secLabelColor'] ? this.valueAxisProperties['secLabelColor'] : null;
                    
                    axes.y2.axis
                        .tickPadding(CartesianChart.TickPaddingY)
                        .orient(showY1OnRight ? yAxisPosition.left.toLowerCase() : yAxisPosition.right.toLowerCase());

                    if (duration) {
                        this.y2AxisGraphicsContext
                            .transition()
                            .duration(duration)
                            .call(axes.y2.axis);
                    }
                    else {
                        this.y2AxisGraphicsContext
                            .call(axes.y2.axis);
                    }

                    this.y2AxisGraphicsContext
                        .call(CartesianChart.darkenZeroLine)
                        .call(CartesianChart.setAxisLabelColor, y2LabelColor);

                    if (tickLabelMargins.yRight >= leftRightMarginLimit) {
                        this.y2AxisGraphicsContext.selectAll('text')
                            .call(AxisHelper.LabelLayoutStrategy.clip,
                                // Can't use padding space to render text, so subtract that from available space for ellipses calculations
                                leftRightMarginLimit - CartesianChart.RightPadding,
                                TextMeasurementService.svgEllipsis);
                    }
                }
                else {
                    this.y2AxisGraphicsContext.selectAll('*').remove();
                    }                    
            }
            else {
                this.y1AxisGraphicsContext.selectAll('*').remove();
                this.y2AxisGraphicsContext.selectAll('*').remove();
            }

            // Axis labels
            //TODO: Add label for second Y axis for combo chart
            if (chartHasAxisLabels) {
                let hideXAxisTitle = !this.shouldRenderAxis(axes.x, "showAxisTitle");
                let hideYAxisTitle = !this.shouldRenderAxis(axes.y1, "showAxisTitle");
                let hideY2AxisTitle = this.valueAxisProperties && this.valueAxisProperties["secShowAxisTitle"] != null && this.valueAxisProperties["secShowAxisTitle"] === false;

                let renderAxisOptions: AxisRenderingOptions = {
                    axisLabels: axisLabels,
                    legendMargin: this.legendMargins.height,
                    viewport: viewport,
                    hideXAxisTitle: hideXAxisTitle,
                    hideYAxisTitle: hideYAxisTitle,
                    hideY2AxisTitle: hideY2AxisTitle,
                    xLabelColor: xLabelColor,
                    yLabelColor: yLabelColor,
                    y2LabelColor: y2LabelColor
                };

                this.renderAxesLabels(renderAxisOptions);
            }
            else {
                this.axisGraphicsContext.selectAll('.xAxisLabel').remove();
                this.axisGraphicsContext.selectAll('.yAxisLabel').remove();
            }

            this.translateAxes(viewport);

            //Render chart columns            
            if (this.behavior) {
                let dataPoints: SelectableDataPoint[] = [];
                let layerBehaviorOptions: any[] = [];
                let labelDataPoints: LabelDataPoint[] = [];
                for (let i = 0, len = layers.length; i < len; i++) {
                    let result = layers[i].render(suppressAnimations);
                    if (result) {
                        dataPoints = dataPoints.concat(result.dataPoints);
                        layerBehaviorOptions.push(result.behaviorOptions);
                        labelDataPoints = labelDataPoints.concat(result.labelDataPoints);
                    }
                }
                labelDataPoints = NewDataLabelUtils.removeDuplicates(labelDataPoints);
                let labelLayout = new LabelLayout({
                    maximumOffset: NewDataLabelUtils.maxLabelOffset,
                    startingOffset: NewDataLabelUtils.startingLabelOffset
                });
                let dataLabels = labelLayout.layout(labelDataPoints, chartViewport);
                if (layers.length > 1) {
                    NewDataLabelUtils.drawLabelBackground(this.labelGraphicsContextScrollable, dataLabels, "#FFFFFF", 0.7);
                }
                if (this.animator && !suppressAnimations) {
                    NewDataLabelUtils.animateDefaultLabels(this.labelGraphicsContextScrollable, dataLabels, this.animator.getDuration());
                }
                else {
                    NewDataLabelUtils.drawDefaultLabels(this.labelGraphicsContextScrollable, dataLabels);
                }
                if (this.interactivityService) {
                    let behaviorOptions: CartesianBehaviorOptions = {
                        layerOptions: layerBehaviorOptions,
                        clearCatcher: this.clearCatcher,
                    };
                    this.interactivityService.bind(dataPoints, this.behavior, behaviorOptions);
                }
            }
            else {
                let labelDataPoints: LabelDataPoint[] = [];
                for (let i = 0, len = layers.length; i < len; i++) {
                    let result = layers[i].render(suppressAnimations);
                    if (result) // Workaround until out of date mobile render path for line chart is removed
                        labelDataPoints = labelDataPoints.concat(result.labelDataPoints);
                }
                labelDataPoints = NewDataLabelUtils.removeDuplicates(labelDataPoints);
                let labelLayout = new LabelLayout({
                    maximumOffset: NewDataLabelUtils.maxLabelOffset,
                    startingOffset: NewDataLabelUtils.startingLabelOffset
                });
                let dataLabels = labelLayout.layout(labelDataPoints, chartViewport);
                if (layers.length > 1) {
                    NewDataLabelUtils.drawLabelBackground(this.labelGraphicsContextScrollable, dataLabels, "#FFFFFF", 0.7);
                }
                NewDataLabelUtils.drawDefaultLabels(this.labelGraphicsContextScrollable, dataLabels);
            }
        }
        
        /**
         * Within the context of the given selection (g), find the offset of
         * the zero tick using the d3 attached datum of g.tick elements.
         * 'Classed' is undefined for transition selections
         */
        private static darkenZeroLine(g: D3.Selection): void {
            let zeroTick = g.selectAll('g.tick').filter((data) => data === 0).node();
            if (zeroTick) {
                d3.select(zeroTick).select('line').classed('zero-line', true);
            }
        }

        private static setAxisLabelColor(g: D3.Selection, fill: Fill): void {
            g.selectAll('g.tick text').style('fill', fill ? fill.solid.color : null);
        }
        
        /**
         * Returns the actual viewportWidth if visual is not scrollable.
         * @return If visual is scrollable, returns the plot area needed to draw all the datapoints.
         */
        public static getPreferredPlotArea(
            categoryCount: number,
            categoryThickness: number,
            viewport: IViewport,
            isScrollable: boolean,
            isScalar: boolean): IViewport {

            let preferredViewport: IViewport = {
                height: viewport.height,
                width: viewport.width
            };
            if (!isScalar && isScrollable) {
                let preferredWidth = CartesianChart.getPreferredCategorySpan(categoryCount, categoryThickness);
                preferredViewport.width = Math.max(preferredWidth, viewport.width);
            }
            return preferredViewport;
        }

        /**
         * Returns preferred Category span if the visual is scrollable.
         */
        public static getPreferredCategorySpan(categoryCount: number, categoryThickness: number): number {
            return categoryThickness * (categoryCount + (CartesianChart.OuterPaddingRatio * 2));
        }
        
        /**
         * Note: Public for testing access.
         */
        public static getLayout(data: ColumnChartData, options: CategoryLayoutOptions): CategoryLayout {
            let categoryCount = options.categoryCount,
                availableWidth = options.availableWidth,
                domain = options.domain,
                isScalar = !!options.isScalar,
                isScrollable = !!options.isScrollable;

            let categoryThickness = CartesianChart.getCategoryThickness(data ? data.series : null, categoryCount, availableWidth, domain, isScalar);

            // Total width of the outer padding, the padding that exist on the far right and far left of the chart.
            let totalOuterPadding = categoryThickness * CartesianChart.OuterPaddingRatio * 2;

            // visibleCategoryCount will be used to discard data that overflows on ordinal-axis charts.
            // Needed for dashboard visuals            
            let calculatedBarCount = Math.round((availableWidth - totalOuterPadding) / categoryThickness);
            let visibleCategoryCount = Math.min(calculatedBarCount, categoryCount);

            let outerPaddingRatio = CartesianChart.OuterPaddingRatio;
            if (!isScalar) {
                // use dynamic outer padding
                let oneOuterPadding = (availableWidth - (categoryThickness * visibleCategoryCount)) / 2;
                outerPaddingRatio = oneOuterPadding / categoryThickness;
            }

            // If scrollable, visibleCategoryCount will be total categories
            if (!isScalar && isScrollable)
                visibleCategoryCount = categoryCount;

            return {
                categoryCount: visibleCategoryCount,
                categoryThickness: categoryThickness,
                outerPaddingRatio: outerPaddingRatio,
                isScalar: isScalar
            };
        }

        /** 
         * Returns the thickness for each category.
         * For clustered charts, you still need to divide by
         * the number of series to get column width after calling this method.
         * For linear or time scales, category thickness accomodates for
         * the minimum interval between consequtive points.
         * For all types, return value has accounted for outer padding,
         * but not inner padding.
         */
        public static getCategoryThickness(seriesList: CartesianSeries[], numCategories: number, plotLength: number, domain: number[], isScalar: boolean): number {
            let thickness;
            if (numCategories < 2)
                thickness = plotLength * (1 - CartesianChart.OuterPaddingRatio);
            else if (isScalar && domain && domain.length > 1) {
                // the smallest interval defines the column width.
                let minInterval = CartesianChart.getMinInterval(seriesList);
                let domainSpan = domain[domain.length - 1] - domain[0];
                // account for outside padding
                let ratio = minInterval / (domainSpan + (minInterval * CartesianChart.OuterPaddingRatio * 2));
                thickness = plotLength * ratio;
                thickness = Math.max(thickness, CartesianChart.MinScalarRectThickness);
            }
            else {
                // Divide the available width up including outer padding (in terms of category thickness) on
                // both sides of the chart, and categoryCount categories. Reverse math:
                // availableWidth = (categoryThickness * categoryCount) + (categoryThickness * (outerPadding * 2)),
                // availableWidth = categoryThickness * (categoryCount + (outerPadding * 2)),
                // categoryThickness = availableWidth / (categoryCount + (outerpadding * 2))
                thickness = plotLength / (numCategories + (CartesianChart.OuterPaddingRatio * 2));
                thickness = Math.max(thickness, CartesianChart.MinOrdinalRectThickness);
            }
            
            // spec calls for using the whole plot area, but the max rectangle thickness is "as if there were three categories"
            // (outerPaddingRatio has the same units as '# of categories' so they can be added)
            let maxRectThickness = plotLength / (3 + (CartesianChart.OuterPaddingRatio * 2));

            if (!isScalar && numCategories >= 3)
                return Math.max(Math.min(thickness, maxRectThickness), CartesianChart.MinOrdinalRectThickness);

            return Math.min(thickness, maxRectThickness);
        }

        private static getMinInterval(seriesList: CartesianSeries[]): number {
            let minInterval = Number.MAX_VALUE;
            if (seriesList.length > 0) {
                let series0data = seriesList[0].data.filter(d => !d.highlight);
                for (let i = 0, ilen = series0data.length - 1; i < ilen; i++) {
                    minInterval = Math.min(minInterval, Math.abs(series0data[i + 1].categoryValue - series0data[i].categoryValue));
                }
            }
            return minInterval;
        }
    }

    function getLayerData(dataViews: DataView[], currentIdx: number, totalLayers: number): DataView[] {
        if (totalLayers > 1) {
            if (dataViews && dataViews.length > currentIdx)
                return [dataViews[currentIdx]];
            return [];
        }

        return dataViews;
    }

    function hasMultipleYAxes(layers: ICartesianVisual[]): boolean {
        debug.assertValue(layers, 'layers');

        return layers.length > 1;
    }

    /**
     * Returns a boolean, that indicates if y axis title should be displayed.
     * @return True if y axis title should be displayed,
     * otherwise false.
     */
    function shouldShowYAxisLabel(layerNumber: number, valueAxisProperties: DataViewObject, yAxisWillMerge: boolean): boolean {
        return ((layerNumber === 0 && !!valueAxisProperties && !!valueAxisProperties['showAxisTitle']) ||
            (layerNumber === 1 && !yAxisWillMerge && !!valueAxisProperties && !!valueAxisProperties['secShowAxisTitle']));
    }

    function tryMergeYDomains(layers: ICartesianVisual[], visualOptions: CalculateScaleAndDomainOptions): MergedValueAxisResult {
        debug.assert(layers.length < 3, 'merging of more than 2 layers is not supported');

        let noMerge: MergedValueAxisResult = {
            domain: undefined,
            merged: false,
            tickCount: undefined,
            forceStartToZero: false
        };

        if (layers.length < 2)
            return noMerge;

        let min: number;
        let max: number;
        let minOfMax: number;
        let maxOfMin: number;

        // TODO: replace full calculateAxesProperties with just a data domain calc
        // we need to be aware of which chart require zero (column/bar) and which don't (line)
        let y1props = layers[0].calculateAxesProperties(visualOptions)[1];
        let y2props = layers[1].calculateAxesProperties(visualOptions)[1];
        let firstYDomain = y1props.scale.domain();
        let secondYDomain = y2props.scale.domain();

        if (firstYDomain[0] >= 0 && secondYDomain[0] >= 0) {
            noMerge.forceStartToZero = true;
        }

        if (y1props.values && y1props.values.length > 0 && y2props.values && y2props.values.length > 0) {
            noMerge.tickCount = Math.max(y1props.values.length, y2props.values.length);
        }

        min = Math.min(firstYDomain[0], secondYDomain[0]);
        max = Math.max(firstYDomain[1], secondYDomain[1]);

        if (visualOptions.forceMerge) {
            return {
                domain: [min, max],
                merged: true,
                tickCount: noMerge.tickCount,
                forceStartToZero: false
            };
        }

        // If domains don't intersect don't merge axis.
        if (firstYDomain[0] > secondYDomain[1] || firstYDomain[1] < secondYDomain[0])
            return noMerge;

        maxOfMin = Math.max(firstYDomain[0], secondYDomain[0]);
        minOfMax = Math.min(firstYDomain[1], secondYDomain[1]);

        let range = (max - min);

        if (range === 0) {
            return noMerge;
        }

        let intersection = Math.abs((minOfMax - maxOfMin) / range);

        // Only merge if intersection of domains greater than 10% of total range.
        if (intersection < COMBOCHART_DOMAIN_OVERLAP_TRESHOLD_PERCENTAGE)
            return noMerge;
        else
            return {
                domain: [min, max],
                merged: true,
                tickCount: noMerge.tickCount,
                forceStartToZero: false
            };
    }
    
    /** 
     * Computes the Cartesian Chart axes from the set of layers.
     */
    function calculateAxes(
        layers: ICartesianVisual[],
        viewport: IViewport,
        margin: IMargin,
        categoryAxisProperties: DataViewObject,
        valueAxisProperties: DataViewObject,
        textProperties: TextProperties,
        scrollbarVisible: boolean,
        existingAxisProperties: CartesianAxisProperties): CartesianAxisProperties {
        debug.assertValue(layers, 'layers');

        let visualOptions: CalculateScaleAndDomainOptions = {
            viewport: viewport,
            margin: margin,
            forcedXDomain: [categoryAxisProperties ? categoryAxisProperties['start'] : null, categoryAxisProperties ? categoryAxisProperties['end'] : null],
            forceMerge: valueAxisProperties && valueAxisProperties['secShow'] === false,
            showCategoryAxisLabel: false,
            showValueAxisLabel: false,
            categoryAxisScaleType: categoryAxisProperties && categoryAxisProperties['axisScale'] != null ? <string>categoryAxisProperties['axisScale'] : axisScale.linear,
            valueAxisScaleType: valueAxisProperties && valueAxisProperties['axisScale'] != null ? <string>valueAxisProperties['axisScale'] : axisScale.linear
        };

        let skipMerge = valueAxisProperties && valueAxisProperties['secShow'] === true;
        let yAxisWillMerge = false;
        let mergeResult: MergedValueAxisResult;
        if (hasMultipleYAxes(layers) && !skipMerge) {
            mergeResult = tryMergeYDomains(layers, visualOptions);
            yAxisWillMerge = mergeResult.merged;
            if (yAxisWillMerge) {
                visualOptions.forcedYDomain = mergeResult.domain;
            }
            else {
                visualOptions.forcedTickCount = mergeResult.tickCount;
            }
        }

        if (valueAxisProperties) {
            visualOptions.forcedYDomain = AxisHelper.applyCustomizedDomain([valueAxisProperties['start'], valueAxisProperties['end']], visualOptions.forcedYDomain);
        }

        let result: CartesianAxisProperties;
        for (let layerNumber = 0, len = layers.length; layerNumber < len; layerNumber++) {
            let currentlayer = layers[layerNumber];

            if (layerNumber === 1 && !yAxisWillMerge) {
                visualOptions.forcedYDomain = valueAxisProperties ? [valueAxisProperties['secStart'], valueAxisProperties['secEnd']] : null;
                visualOptions.valueAxisScaleType = valueAxisProperties && valueAxisProperties['secAxisScale'] != null ? <string>valueAxisProperties['secAxisScale'] : axisScale.linear;
                if (mergeResult && mergeResult.forceStartToZero) {
                    if (!visualOptions.forcedYDomain) {
                        visualOptions.forcedYDomain = [0, undefined];
                    }
                    else if (visualOptions.forcedYDomain[0] == null) {
                        visualOptions.forcedYDomain[0] = 0;//only set when user didn't choose a value
                    }
                }
            }
            visualOptions.showCategoryAxisLabel = (!!categoryAxisProperties && !!categoryAxisProperties['showAxisTitle']);//here

            visualOptions.showValueAxisLabel = shouldShowYAxisLabel(layerNumber, valueAxisProperties, yAxisWillMerge);

            let axes = currentlayer.calculateAxesProperties(visualOptions);

            if (layerNumber === 0) {
                result = {
                    x: axes[0],
                    y1: axes[1]
                };
            }
            else if (axes && !result.y2) {
                if (axes[0].axis.scale().domain().length > result.x.axis.scale().domain().length) {
                    visualOptions.showValueAxisLabel = (!!valueAxisProperties && !!valueAxisProperties['showAxisTitle']);

                    let axes = currentlayer.calculateAxesProperties(visualOptions);
                    // no categories returned for the first layer, use second layer x-axis properties
                    result.x = axes[0];
                    // and 2nd value axis to be the primary
                    result.y1 = axes[1];
                }
                else {
                    // make sure all layers use the same x-axis/scale for drawing
                    currentlayer.overrideXScale(result.x);
                    if (!yAxisWillMerge && !axes[1].usingDefaultDomain)
                        result.y2 = axes[1];
                }
            }

            if (existingAxisProperties && existingAxisProperties.x) {
                result.x.willLabelsFit = existingAxisProperties.x.willLabelsFit;
                result.x.willLabelsWordBreak = existingAxisProperties.x.willLabelsWordBreak;
            } else {
                let width = viewport.width - (margin.left + margin.right);
                result.x.willLabelsFit = AxisHelper.LabelLayoutStrategy.willLabelsFit(
                    result.x,
                    width,
                    TextMeasurementService.measureSvgTextWidth,
                    textProperties);

                // If labels do not fit and we are not scrolling, try word breaking
                result.x.willLabelsWordBreak = (!result.x.willLabelsFit && !scrollbarVisible) && AxisHelper.LabelLayoutStrategy.willLabelsWordBreak(
                    result.x,
                    margin,
                    width,
                    TextMeasurementService.measureSvgTextWidth,
                    TextMeasurementService.estimateSvgTextHeight,
                    TextMeasurementService.getTailoredTextOrDefault,
                    textProperties);
            }
        }

        return result;
    }

    module CartesianLayerFactory {

        export function createLayers(
            type: CartesianChartType,
            objects: DataViewObjects,
            interactivityService: IInteractivityService,
            animator?: any,
            isScrollable: boolean = false,
            seriesLabelFormattingEnabled: boolean = false): ICartesianVisual[]{

            let layers: ICartesianVisual[] = [];

            let cartesianOptions: CartesianVisualConstructorOptions = {
                isScrollable: isScrollable,
                animator: animator,
                interactivityService: interactivityService,
                seriesLabelFormattingEnabled: seriesLabelFormattingEnabled,
            };

            switch (type) {
                case CartesianChartType.Area:
                    layers.push(createLineChartLayer(LineChartType.area, /* inComboChart */ false, cartesianOptions));
                    break;
                case CartesianChartType.Line:
                    layers.push(createLineChartLayer(LineChartType.default, /* inComboChart */ false, cartesianOptions));
                    break;
                case CartesianChartType.Scatter:
                    layers.push(createScatterChartLayer(cartesianOptions));
                    break;
                case CartesianChartType.Play:
                    layers.push(createPlayChartLayer(cartesianOptions, animator == null));
                    break;
                case CartesianChartType.Waterfall:
                    layers.push(createWaterfallChartLayer(cartesianOptions));
                    break;
                case CartesianChartType.DataDot:
                    layers.push(createDataDotChartLayer(cartesianOptions));
                    break;
                case CartesianChartType.StackedColumn:
                    layers.push(createColumnChartLayer(ColumnChartType.stackedColumn, cartesianOptions));
                    break;
                case CartesianChartType.ClusteredColumn:
                    layers.push(createColumnChartLayer(ColumnChartType.clusteredColumn, cartesianOptions));
                    break;
                case CartesianChartType.HundredPercentStackedColumn:
                    layers.push(createColumnChartLayer(ColumnChartType.hundredPercentStackedColumn, cartesianOptions));
                    break;
                case CartesianChartType.StackedBar:
                    layers.push(createColumnChartLayer(ColumnChartType.stackedBar, cartesianOptions));
                    break;
                case CartesianChartType.ClusteredBar:
                    layers.push(createColumnChartLayer(ColumnChartType.clusteredBar, cartesianOptions));
                    break;
                case CartesianChartType.HundredPercentStackedBar:
                    layers.push(createColumnChartLayer(ColumnChartType.hundredPercentStackedBar, cartesianOptions));
                    break;
                case CartesianChartType.ComboChart:
                    let columnType = getComboColumnType();
                    layers.push(createColumnChartLayer(columnType, cartesianOptions));
                    layers.push(createLineChartLayer(LineChartType.default, /* inComboChart */ true, cartesianOptions));
                    break;
                case CartesianChartType.LineClusteredColumnCombo:
                    layers.push(createColumnChartLayer(ColumnChartType.clusteredColumn, cartesianOptions));
                    layers.push(createLineChartLayer(LineChartType.default, /* inComboChart */ true, cartesianOptions));
                    break;
                case CartesianChartType.LineStackedColumnCombo:
                    layers.push(createColumnChartLayer(ColumnChartType.stackedColumn, cartesianOptions));
                    layers.push(createLineChartLayer(LineChartType.default, /* inComboChart */ true, cartesianOptions));
                    break;
                case CartesianChartType.DataDotClusteredColumnCombo:
                    layers.push(createColumnChartLayer(ColumnChartType.clusteredColumn, cartesianOptions));
                    layers.push(createDataDotChartLayer(cartesianOptions));
                    break;
                case CartesianChartType.DataDotStackedColumnCombo:
                    layers.push(createColumnChartLayer(ColumnChartType.stackedColumn, cartesianOptions));
                    layers.push(createDataDotChartLayer(cartesianOptions));
                    break;
            }

            return layers;
        }

        function createLineChartLayer(type: LineChartType, inComboChart: boolean, defaultOptions: CartesianVisualConstructorOptions): LineChart {
            let options: LineChartConstructorOptions = {
                animator: defaultOptions.animator,
                interactivityService: defaultOptions.interactivityService,
                isScrollable: defaultOptions.isScrollable,
                seriesLabelFormattingEnabled: defaultOptions.seriesLabelFormattingEnabled,
                chartType: type
            };

            if (inComboChart) {
                options.chartType = options.chartType | LineChartType.lineShadow;
            }

            return new LineChart(options);
        }

        function createScatterChartLayer(defaultOptions: CartesianVisualConstructorOptions): ScatterChart {
            defaultOptions.isScrollable = false;
            return new ScatterChart(defaultOptions);
        }

        function createPlayChartLayer(defaultOptions: CartesianVisualConstructorOptions, isFrozen: boolean): PlayChart {
            let options: PlayChartConstructorOptions = {
                animator: defaultOptions.animator,
                interactivityService: defaultOptions.interactivityService,
                isScrollable: false,
                isFrozen: isFrozen,
            };

            return new PlayChart(options);
        }

        function createWaterfallChartLayer(defaultOptions: CartesianVisualConstructorOptions): WaterfallChart {
            return new WaterfallChart(defaultOptions);
        }

        function createDataDotChartLayer(defaultOptions: CartesianVisualConstructorOptions): DataDotChart {
            return new DataDotChart(defaultOptions);
        }

        function createColumnChartLayer(type: ColumnChartType, defaultOptions: CartesianVisualConstructorOptions): ColumnChart {
            let options: ColumnChartConstructorOptions = {
                animator: <IColumnChartAnimator>defaultOptions.animator,
                interactivityService: defaultOptions.interactivityService,
                isScrollable: defaultOptions.isScrollable,
                seriesLabelFormattingEnabled: defaultOptions.seriesLabelFormattingEnabled,
                chartType: type
            };
            return new ColumnChart(options);
        }

        function getComboColumnType(objects?: DataViewObjects): ColumnChartType {
            // This supports existing serialized forms of pinned combo-chart visuals
            let columnType: ColumnChartType = ColumnChartType.clusteredColumn;
            if (objects) {
                let comboChartTypes: ComboChartDataViewObject = (<ComboChartDataViewObjects>objects).general;
                if (comboChartTypes) {
                    switch (comboChartTypes.visualType1) {
                        case 'Column':
                            columnType = ColumnChartType.clusteredColumn;
                            break;
                        case 'ColumnStacked':
                            columnType = ColumnChartType.stackedColumn;
                            break;
                        default:
                            debug.assertFail('Unsupported cartesian chart type ' + comboChartTypes.visualType1);
                    }

                    // second visual is always LineChart (for now)
                    if (comboChartTypes.visualType2) {
                        debug.assert(comboChartTypes.visualType2 === 'Line', 'expecting a LineChart for VisualType2');
                    }
                }
            }

            return columnType;
        }
    }

    export class SharedColorPalette implements IDataColorPalette {
        private palette: IDataColorPalette;
        private preferredScale: IColorScale;
        private rotated: boolean;

        constructor(palette: IDataColorPalette) {
            this.palette = palette;
            this.clearPreferredScale();
        }

        public getColorScaleByKey(scaleKey: string): IColorScale {
            this.setPreferredScale(scaleKey);
            return this.preferredScale;
        }

        public getNewColorScale(): IColorScale {
            return this.preferredScale;
        }

        public getColorByIndex(index: number): IColorInfo {
            return this.palette.getColorByIndex(index);
        }

        public getSentimentColors(): IColorInfo[] {
            return this.palette.getSentimentColors();
        }

        public getBasePickerColors(): IColorInfo[] {
            return this.palette.getBasePickerColors();
        }

        public clearPreferredScale(): void {
            this.preferredScale = this.palette.getNewColorScale();
            this.rotated = false;
        }

        public rotateScale(): void {
            // We create a new rotated the scale such that the first color of the new scale is the first
            // free color of the previous scale. Note that the new scale does not have any colors allocated
            // to particular keys.
            this.preferredScale = this.preferredScale.clone();
            this.preferredScale.clearAndRotateScale();
            this.rotated = true;
        }

        private setPreferredScale(scaleKey: string): void {
            if (!this.rotated) {
                // The first layer to express a preference sets the preferred scale.
                this.preferredScale = this.palette.getColorScaleByKey(scaleKey);
            }
        }
    }
}
