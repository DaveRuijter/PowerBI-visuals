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
    export interface GaugeData extends TooltipEnabledDataPoint {
        percent: number;
        adjustedTotal: number;
        total: number;
        metadataColumn: DataViewMetadataColumn;
        targetSettings: GaugeTargetSettings;
    }

    interface KpiArcAttributes {
        start: number;
        end: number;
        fill: string;
    }

    export interface GaugeTargetSettings {
        min: number;
        max: number;
        target: number;
    }

    export interface GaugeTargetData extends GaugeTargetSettings {
        total: number;
        tooltipItems: TooltipSeriesDataItem[];
    }

    interface GaugeStyle {
        transition: {
            ease: string
        };
        arcColors: {
            background: string;
            foreground: string;
        };
        targetLine: {
            show: boolean;
            color: string;
            thickness: number;
        };
        labels: {
            count: number;
            padding: number;
        };
        kpiBands: {
            show: boolean;
            separationRadians: number;
            thickness: number;
        };
    }

    export interface GaugeSmallViewPortProperties {
        hideGaugeSideNumbersOnSmallViewPort: boolean;
        smallGaugeMarginsOnSmallViewPort: boolean;
        MinHeightGaugeSideNumbersVisible: number;
        GaugeMarginsOnSmallViewPort: number;
    }

    export interface GaugeVisualProperties {
        radius: number;
        innerRadiusOfArc: number;
        innerRadiusFactor: number;
        left: number;
        top: number;
        height: number;
        width: number;
        margin: IMargin;
        transformString: string;
    }

    export interface AnimatedNumberProperties {
        transformString: string;
        viewport: IViewport;
    }

    export interface GaugeConstructorOptions {
        gaugeSmallViewPortProperties?: GaugeSmallViewPortProperties;
        animator?: IGenericAnimator;
    }

    export interface GaugeDataViewObjects extends DataViewObjects {
        axis: GaugeDataViewObject;
    }

    export interface GaugeDataViewObject extends DataViewObject {
        min?: number;
        max?: number;
        target?: number;
    }

    /** 
     * Renders a number that can be animate change in value.
     */
    export class Gauge implements IVisual {
        private static MIN_VALUE = -Infinity;
        private static MAX_VALUE = +Infinity;
        private static MinDistanceFromBottom = 10;
        private static MinWidthForTargetLabel = 150;
        private static DefaultTopBottomMargin = 20;
        private static DefaultLeftRightMargin = 45;
        private static ReducedLeftRightMargin = 15;
        private static DEFAULT_MAX = 1;
        private static DEFAULT_MIN = 0;
        private static VisualClassName = 'gauge';
        private static DefaultStyleProperties: GaugeStyle = {
            transition: {
                ease: 'bounce'
            },
            arcColors: {
                background: '#e9e9e9',
                foreground: '#00B8AA'
            },
            targetLine: {
                show: true,
                color: '#666666',
                thickness: 2
            },
            labels: {
                count: 2,
                padding: 5
            },
            kpiBands: {
                show: false,
                separationRadians: Math.PI / 128,
                thickness: 5
            },
        };
        private static DefaultTargetSettings: GaugeTargetSettings = {
            min: 0,
            max: 1,
            target: undefined
        };

        private static InnerRadiusFactor = 0.7;
        private static KpiBandDistanceFromMainArc = 2;

        private static MainGaugeGroupElementName = 'mainGroup';
        private static LabelText: ClassAndSelector = {
            class: 'labelText',
            selector: '.labelText'
        };

        private static TargetConnector: ClassAndSelector = {
            class: 'targetConnector',
            selector: '.targetConnector'
        };

        private static TargetText: ClassAndSelector = {
            class: 'targetText',
            selector: '.targetText'
        };

        /** Note: Public for testability */
        public static formatStringProp: DataViewObjectPropertyIdentifier = {
            objectName: 'general',
            propertyName: 'formatString',
        };

        private svg: D3.Selection;
        private mainGraphicsContext: D3.Selection;
        private currentViewport: IViewport;
        private element: JQuery;
        private style: IVisualStyle;
        private data: GaugeData;
        private color: D3.Scale.OrdinalScale;

        private backgroundArc: D3.Svg.Arc;
        private foregroundArc: D3.Svg.Arc;
        private kpiArcs: D3.Svg.Arc[];

        private kpiArcPaths: D3.Selection[];
        private foregroundArcPath: D3.Selection;
        private backgroundArcPath: D3.Selection;
        private targetLine: D3.Selection;
        private targetConnector: D3.Selection;
        private targetText: D3.Selection;
        private options: VisualInitOptions;

        private lastAngle = -Math.PI / 2;
        private margin: IMargin;
        private animatedNumberGrapicsContext: D3.Selection;
        private animatedNumber: AnimatedNumber;
        private settings: GaugeStyle;
        private targetSettings: GaugeTargetSettings;
        private gaugeVisualProperties: GaugeVisualProperties;
        private gaugeSmallViewPortProperties: GaugeSmallViewPortProperties;
        private showTargetLabel: boolean;

        private hostService: IVisualHostServices;

        // TODO: Remove this once all visuals have implemented update.
        private dataViews: DataView[];

        public animator: IGenericAnimator;

        constructor(options?: GaugeConstructorOptions) {
            if (options) {
                if (options.gaugeSmallViewPortProperties) {
                    this.gaugeSmallViewPortProperties = options.gaugeSmallViewPortProperties;
                }
                this.animator = options.animator;
            }
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] {
            switch (options.objectName) {
                case 'axis':
                    return this.enumerateAxis();
            }
            return null;
        }

        private enumerateAxis(): VisualObjectInstance[] {
            let dataView: DataView = this.dataViews[0];

            if (dataView && dataView.metadata) {
                let properties: GaugeTargetSettings = Gauge.getGaugeObjectsProperties(dataView);

                return [{
                    selector: null,
                    objectName: 'axis',
                    properties: <any>properties,
                }];
            }
            return null;
        }
        
        private static getGaugeObjectsProperties(dataView: DataView): GaugeTargetSettings {
            let properties: any = {};
            let objects: GaugeDataViewObjects = <GaugeDataViewObjects>dataView.metadata.objects;
            let hasAxisObject: boolean = !!objects && !!objects.axis;

            if (!DataRoleHelper.hasRoleInDataView(dataView, gaugeRoleNames.minValue))
                properties.min = hasAxisObject ? objects.axis.min : undefined;

            if (!DataRoleHelper.hasRoleInDataView(dataView, gaugeRoleNames.maxValue))
                properties.max = hasAxisObject ? objects.axis.max : undefined;

            if (!DataRoleHelper.hasRoleInDataView(dataView, gaugeRoleNames.targetValue))
                properties.target = hasAxisObject ? objects.axis.target : undefined;

            return properties;
        }
        

        public init(options: VisualInitOptions) {
            this.element = options.element;
            this.currentViewport = options.viewport;
            this.style = options.style;
            this.options = options;
            this.settings = Gauge.DefaultStyleProperties;
            this.targetSettings = Gauge.DefaultTargetSettings;

            this.setMargins();

            this.color = d3.scale.ordinal().range(
                this.style.colorPalette.dataColors.getSentimentColors().map(
                    color => color.value));

            this.hostService = options.host;
            let svg = this.svg = d3.select(this.element.get(0)).append('svg');
            svg.classed(Gauge.VisualClassName, true);
            let mainGraphicsContext = this.mainGraphicsContext = svg.append('g');
            mainGraphicsContext.attr('class', Gauge.MainGaugeGroupElementName);

            this.initKpiBands();

            let backgroundArc = this.backgroundArc = d3.svg.arc()
                .innerRadius(0)
                .outerRadius(0)
                .startAngle(-Math.PI / 2)
                .endAngle(Math.PI / 2);

            let foregroundArc = this.foregroundArc = d3.svg.arc()
                .innerRadius(0)
                .outerRadius(0)
                .startAngle(-Math.PI / 2);

            this.backgroundArcPath = mainGraphicsContext.append('path')
                .classed('backgroundArc', true)
                .attr('d', backgroundArc);

            this.foregroundArcPath = mainGraphicsContext.append('path')
                .datum({ endAngle: -Math.PI / 2 })
                .classed('foregroundArc', true)
                .attr('d', foregroundArc);

            let g = this.animatedNumberGrapicsContext = svg.append('g');

            this.animatedNumber = new AnimatedNumber(g);
            this.animatedNumber.init(options);

            let gaugeDrawingOptions = this.gaugeVisualProperties = this.getGaugeVisualProperties();
            let animatedNumberProperties = this.getAnimatedNumberProperties(
                gaugeDrawingOptions.radius,
                gaugeDrawingOptions.innerRadiusFactor,
                gaugeDrawingOptions.top,
                gaugeDrawingOptions.left);
            this.animatedNumber.svg.attr('transform', animatedNumberProperties.transformString);
            this.animatedNumber.onResizing(animatedNumberProperties.viewport);
        }

        public update(options: VisualUpdateOptions) {
            debug.assertValue(options, 'options');

            this.currentViewport = options.viewport;
            let dataViews = this.dataViews = options.dataViews;

            if (!dataViews || !dataViews[0]) {
                return;
            }

            this.data = Gauge.converter(dataViews[0]);
            this.targetSettings = this.data.targetSettings;

            if (dataViews[0])
                dataViews[0].single = { value: this.data.total };

            // Only show the target label if:
            //   1. There is a target
            //   2. The viewport width is big enough for a target
            //   3. We're showing label text for side numbers
            this.showTargetLabel =
            this.targetSettings.target != null &&
            (this.currentViewport.width > Gauge.MinWidthForTargetLabel || !this.showMinMaxLabelsOnBottom()) &&
            this.showSideNumbersLabelText();

            this.setMargins();

            let gaugeDrawingOptions = this.gaugeVisualProperties = this.getGaugeVisualProperties();
            let animatedNumberProperties = this.getAnimatedNumberProperties(
                gaugeDrawingOptions.radius,
                gaugeDrawingOptions.innerRadiusFactor,
                gaugeDrawingOptions.top,
                gaugeDrawingOptions.left);

            this.drawViewPort(this.gaugeVisualProperties);
            this.updateInternal(options.suppressAnimations);
            this.animatedNumber.svg.attr('transform', animatedNumberProperties.transformString);
            this.animatedNumber.update({
                viewport: animatedNumberProperties.viewport,
                dataViews: options.dataViews,
                suppressAnimations: options.suppressAnimations,
            });

            let warnings = getInvalidValueWarnings(
                dataViews,
                false /*supportsNaN*/,
                false /*supportsNegativeInfinity*/,
                false /*supportsPositiveInfinity*/);

            if (warnings && warnings.length > 0)
                this.hostService.setWarnings(warnings);
        }

        public onDataChanged(options: VisualDataChangedOptions): void {
            // TODO: Remove onDataChanged & onResizing once all visuals have implemented update.
            this.update({
                dataViews: options.dataViews,
                suppressAnimations: options.suppressAnimations,
                viewport: this.currentViewport
            });
        }

        public onResizing(viewport: IViewport): void {
            // TODO: Remove onDataChanged & onResizing once all visuals have implemented update.
            this.update({
                dataViews: this.dataViews,
                suppressAnimations: true,
                viewMode: ViewMode.View,
                viewport: viewport
            });
        }

        public onStyleChanged(newStyle: IVisualStyle) {
            this.style = newStyle;
            this.color = d3.scale.ordinal().range(
                newStyle.colorPalette.dataColors.getSentimentColors().map(
                    color => color.value));
            this.updateInternal(true /* suppressAnimations */);
        }

        private static getValidSettings(targetData: GaugeTargetData): GaugeTargetSettings {
            let maxVal = (targetData.max === Gauge.MAX_VALUE) ? Gauge.DEFAULT_MAX : targetData.max;
            let minVal = (targetData.min === Gauge.MIN_VALUE) ? Gauge.DEFAULT_MIN : targetData.min;
            let targetVal = targetData.target;

            return {
                min: minVal,
                max: maxVal,
                target: targetVal
            };
        }

        private static getGaugeData(dataView: DataView): GaugeTargetData {
            let settings: GaugeTargetData = {
                max: Gauge.MAX_VALUE,
                min: Gauge.MIN_VALUE,
                target: undefined,
                total: 0,
                tooltipItems: []
            };

            if (dataView && dataView.categorical && dataView.categorical.values && dataView.metadata && dataView.metadata.columns) {
                let values = dataView.categorical.values;
                let metadataColumns = dataView.metadata.columns;

                debug.assert(metadataColumns.length >= values.length, 'length');

                for (let i = 0; i < values.length; i++) {
                    let col = metadataColumns[i],
                        value = values[i].values[0] || 0;
                    if (col && col.roles) {
                        if (col.roles[gaugeRoleNames.y]) {
                            settings.total = value;
                            if (value)
                                settings.tooltipItems.push({ value: value, metadata: values[i] });
                        } else if (col.roles[gaugeRoleNames.minValue]) {
                            settings.min = value;
                        } else if (col.roles[gaugeRoleNames.maxValue]) {
                            settings.max = value;
                        } else if (col.roles[gaugeRoleNames.targetValue]) {
                            settings.target = value;
                            if (value)
                                settings.tooltipItems.push({ value: value, metadata: values[i] });
                        }
                    }
                }

                // Override settings according to property pane axis values
                let gaugeObjectsSettings: GaugeTargetSettings = Gauge.getGaugeObjectsProperties(dataView);
                if (gaugeObjectsSettings && !$.isEmptyObject(gaugeObjectsSettings))
                    Gauge.overrideGaugeSettings(settings, gaugeObjectsSettings);
            }

            return settings;
        }
        
        private static overrideGaugeSettings(settings: GaugeTargetData, gaugeObjectsSettings: GaugeTargetSettings) {
            if ($.isNumeric(gaugeObjectsSettings.min))
                settings.min = gaugeObjectsSettings.min;

            if ($.isNumeric(gaugeObjectsSettings.max))
                settings.max = gaugeObjectsSettings.max;

            if ($.isNumeric(gaugeObjectsSettings.target))
                settings.target = gaugeObjectsSettings.target;
        }
        
        /** Note: Made public for testability */
        public static converter(dataView: DataView): GaugeData {
            let gaugeData = Gauge.getGaugeData(dataView);
            let total = gaugeData.total;
            if (total > 0 && gaugeData.max === Gauge.MAX_VALUE) {
                let hasPercent = false;
                let columns = dataView.metadata.columns;
                if (!_.isEmpty(columns)) {
                    let formatString = valueFormatter.getFormatString(dataView.metadata.columns[0], Gauge.formatStringProp, true);
                    hasPercent = valueFormatter.getFormatMetadata(formatString).hasPercent;
                }

                gaugeData.max = hasPercent ? Gauge.DEFAULT_MAX : total * 2;
            }

            let settings: GaugeTargetSettings = Gauge.getValidSettings(gaugeData);

            //Checking that the value is plotted inside the guage boundries
            let adjustedTotal = Math.max(total, settings.min);
            adjustedTotal = Math.min(adjustedTotal, settings.max);

            let percent: number = (settings.min !== settings.max)
                ? (adjustedTotal - settings.min) / (settings.max - settings.min)
                : 0;

            let tooltipInfo: TooltipDataItem[];

            if (dataView) {
                if (gaugeData.tooltipItems.length > 0) {
                    tooltipInfo = TooltipBuilder.createTooltipInfo(Gauge.formatStringProp, null, null, null, null, gaugeData.tooltipItems);
                }
                else {
                    let dataViewCat = dataView.categorical;

                    if (dataViewCat && dataViewCat.values && dataViewCat.values.length > 0) {
                        let categoryValue: DataViewValueColumn = dataViewCat.values[0];
                        let value = categoryValue.values[0];

                        tooltipInfo = TooltipBuilder.createTooltipInfo(Gauge.formatStringProp, dataViewCat, null, value);
                    }
                }
            }

            return {
                percent: percent,
                adjustedTotal: adjustedTotal,
                total: total,
                metadataColumn: Gauge.getMetaDataColumn(dataView),
                targetSettings: settings,
                tooltipInfo: tooltipInfo,
            };
        }

        public static getMetaDataColumn(dataView: DataView) {
            if (dataView && dataView.metadata && dataView.metadata.columns) {
                for (let i = 0, ilen = dataView.metadata.columns.length; i < ilen; i++) {
                    let column = dataView.metadata.columns[i];
                    if (column.isMeasure) {
                        return column;
                    }
                }
            }
            return null;
        }

        private initKpiBands() {
            if (!this.settings.kpiBands.show)
                return;
            let kpiArcs = this.kpiArcs = [];
            let kpiArcPaths = this.kpiArcPaths = [];
            let mainGraphicsContext = this.mainGraphicsContext;

            for (let i = 0; i < 3; i++) {
                let arc = d3.svg.arc()
                    .innerRadius(0)
                    .outerRadius(0)
                    .startAngle(0)
                    .endAngle(0);

                kpiArcs.push(arc);

                let arcPath = mainGraphicsContext.append('path')
                    .attr("d", arc);

                kpiArcPaths.push(arcPath);
            }
        }

        private updateKpiBands(radius: number, innerRadiusFactor: number, tString: string, kpiAngleAttr: KpiArcAttributes[]) {
            if (!this.settings.kpiBands.show)
                return;

            for (let i = 0; i < kpiAngleAttr.length; i++) {
                this.kpiArcs[i]
                    .innerRadius(radius * innerRadiusFactor - (Gauge.KpiBandDistanceFromMainArc + this.settings.kpiBands.thickness))
                    .outerRadius(radius * innerRadiusFactor - Gauge.KpiBandDistanceFromMainArc)
                    .startAngle(kpiAngleAttr[i].start)
                    .endAngle(kpiAngleAttr[i].end);

                this.kpiArcPaths[i]
                    .attr('fill', kpiAngleAttr[i].fill)
                    .attr('d', this.kpiArcs[i])
                    .attr('transform', tString);
            }
        }

        private removeTargetElements() {
            if (this.targetLine) {
                this.targetLine.remove();
                this.targetText.remove();
                this.targetConnector.remove();
                this.targetLine = this.targetConnector = this.targetText = null;
            }
        }

        private updateTargetLine(radius: number, innerRadius: number, left, top) {
            let targetSettings = this.targetSettings;

            if (!this.targetLine) {
                this.targetLine = this.mainGraphicsContext.append('line');
            }

            let angle = (targetSettings.target - targetSettings.min) / (targetSettings.max - targetSettings.min) * Math.PI;

            let outY = top - radius * Math.sin(angle);
            let outX = left - radius * Math.cos(angle);

            let inY = top - innerRadius * Math.sin(angle);
            let inX = left - innerRadius * Math.cos(angle);

            this.targetLine.attr({
                x1: inX,
                y1: inY,
                x2: outX,
                y2: outY
            });
        }

        /** Note: public for testability */
        public getAnimatedNumberProperties(radius: number,
            innerRadiusFactor: number,
            top: number, left: number): AnimatedNumberProperties {
            let boxAngle = Math.PI / 4;
            let scale = 1;
            let innerRadiusOfArc = radius * innerRadiusFactor;
            let innerRadiusForTextBoundingBox = innerRadiusOfArc - (this.settings.kpiBands.show
                ? (Gauge.KpiBandDistanceFromMainArc + this.settings.kpiBands.thickness)
                : 0);
            let innerRCos = innerRadiusForTextBoundingBox * Math.cos(boxAngle);
            let innerRSin = innerRadiusForTextBoundingBox * Math.sin(boxAngle);
            let innerY = top - innerRSin;
            let innerX = left - innerRCos;
            innerY = innerY * scale;
            innerX = innerX * scale;
            let animatedNumberWidth = innerRCos * 2;

            let properties: AnimatedNumberProperties = {
                transformString: SVGUtil.translate(innerX, innerY),
                viewport: { height: innerRSin, width: animatedNumberWidth }
            };
            return properties;
        }

        /** Note: public for testability */
        public getGaugeVisualProperties(): GaugeVisualProperties {
            let viewport = this.currentViewport;
            let margin: IMargin = this.margin;
            let width = viewport.width - margin.right - margin.left;
            let halfWidth = width / 2;
            let height = viewport.height - margin.top - margin.bottom;
            let radius = Math.min(halfWidth, height);
            let innerRadiusFactor = Gauge.InnerRadiusFactor;
            let left = margin.left + halfWidth;
            let top = radius + (height - radius) / 2 + margin.top;
            let tString = SVGUtil.translate(left, top);
            let innerRadiusOfArc = radius * innerRadiusFactor;

            let gaugeData: GaugeVisualProperties = {
                radius: radius,
                innerRadiusOfArc: innerRadiusOfArc,
                left: left,
                top: top,
                height: height,
                width: width,
                margin: margin,
                transformString: tString,
                innerRadiusFactor: innerRadiusFactor
            };

            return gaugeData;
        }

        /** Note: public for testability */
        public drawViewPort(drawOptions: GaugeVisualProperties): void {
            debug.assertAnyValue(drawOptions, "Gauge options");

            let separation = this.settings.kpiBands.separationRadians;
            let innerRadiusFactor = Gauge.InnerRadiusFactor;

            let backgroudArc = this.backgroundArc;
            let color = this.color;

            let attrs: KpiArcAttributes[] = [{
                fill: color(0),
                start: -Math.PI / 2,
                end: -Math.PI / 2 + Math.PI / 4 - separation
            }, {
                    fill: color(1),
                    start: -Math.PI / 2 + Math.PI * 1 / 4 + separation,
                    end: -Math.PI / 2 + Math.PI * 3 / 4 - separation
                }, {
                    fill: color(2),
                    start: -Math.PI / 2 + Math.PI * 3 / 4 + separation,
                    end: Math.PI / 2
                }];

            let radius = drawOptions.radius;
            let transformString = drawOptions.transformString;
            this.updateKpiBands(radius, innerRadiusFactor, transformString, attrs);

            backgroudArc
                .innerRadius(radius * innerRadiusFactor)
                .outerRadius(radius)
                .startAngle(-Math.PI / 2)
                .endAngle(Math.PI / 2);

            this.backgroundArcPath
                .attr("d", backgroudArc)
                .attr("transform", transformString);

            let foregroundArc = this.foregroundArc;

            foregroundArc
                .innerRadius(radius * innerRadiusFactor)
                .outerRadius(radius)
                .startAngle(-Math.PI / 2);

            this.foregroundArcPath
                .datum({ endAngle: this.lastAngle })
                .attr("transform", transformString)
                .attr("d", foregroundArc);

            let innerRadiusOfArc = drawOptions.innerRadiusOfArc;
            let left = drawOptions.left;
            let top = drawOptions.top;
            let margin = drawOptions.margin;
            let height = drawOptions.height;
            let targetSettings = this.targetSettings;
            if (!this.settings.targetLine.show || targetSettings.target == null) {
                this.removeTargetElements();
            } else {
                if (targetSettings.min > targetSettings.target || targetSettings.max < targetSettings.target) {
                    this.removeTargetElements();
                } else {
                    this.updateTargetLine(radius, innerRadiusOfArc, left, top);
                    this.appendTargetTextAlongArc(radius, height, drawOptions.width, margin);
                }
            }
            this.svg.attr('height', this.currentViewport.height).attr('width', this.currentViewport.width);
        }

        private createTicks(total: number): string[] {
            let settings = this.settings;
            let targetSettings = this.targetSettings;
            total = targetSettings.max - targetSettings.min;
            let numberOfLabels = settings.labels.count;
            let step = total / numberOfLabels;
            let arr: string[] = [];
            let data = this.data;

            let formatter = valueFormatter.create({
                format: valueFormatter.getFormatString(data.metadataColumn, Gauge.formatStringProp),
                value: targetSettings.min,
                value2: targetSettings.max,
                formatSingleValues: true,
                allowFormatBeautification: true,
                columnType: (data && data.metadataColumn) ? data.metadataColumn.type : undefined,
            });

            for (let i = 0; i < numberOfLabels + 1; i++) {
                arr.push(formatter.format(targetSettings.min + (i * step)));
            }

            return arr;
        }

        private updateInternal(suppressAnimations: boolean) {
            let height = this.gaugeVisualProperties.height;
            let width = this.gaugeVisualProperties.width;
            let radius = this.gaugeVisualProperties.radius;
            let margin: IMargin = this.margin;
            let duration = AnimatorCommon.GetAnimationDuration(this.animator, suppressAnimations);

            let data = this.data;
            let lastAngle = this.lastAngle = -Math.PI / 2 + Math.PI * data.percent;

            let ticks = this.createTicks(data.adjustedTotal);

            this.foregroundArcPath
                .transition()
                .ease(this.settings.transition.ease)
                .duration(duration)
                .call(this.arcTween, [lastAngle, this.foregroundArc]);

            this.appendTextAlongArc(ticks, radius, height, width, margin);
            this.updateVisualConfigurations();
            this.updateVisualStyles();

            TooltipManager.addTooltip(this.foregroundArcPath, (tooltipEvent: TooltipEvent) => data.tooltipInfo);
        }

        private updateVisualStyles() {
            this.mainGraphicsContext.selectAll('text')
                .style({
                    'fill': this.style.labelText.color.value
                });
        }

        private updateVisualConfigurations() {
            let configOptions = this.settings;

            this.mainGraphicsContext
                .select('line')
                .attr({
                    stroke: configOptions.targetLine.color,
                    'stroke-width': configOptions.targetLine.thickness
                });

            this.backgroundArcPath.style('fill', configOptions.arcColors.background);
            this.foregroundArcPath.style('fill', configOptions.arcColors.foreground);
        }

        private appendTextAlongArc(ticks: string[], radius: number, height: number, width: number, margin: IMargin) {
            this.svg.selectAll(Gauge.LabelText.selector).remove();

            let total = ticks.length;
            let divisor = total - 1;
            let top = (radius + (height - radius) / 2 + margin.top);
            let showMinMaxLabelsOnBottom = this.showMinMaxLabelsOnBottom();
            let fontSize = this.style.labelText.fontSize;
            let padding = this.settings.labels.padding;

            for (let count = 0; count < total; count++) {
                if (Math.floor(total / 2) === count)
                    continue; // Skip Middle label, by design

                if (this.showSideNumbersLabelText()) {

                    let x = (margin.left + width / 2) - (radius * Math.cos(Math.PI * count / divisor));
                    let y = top - (radius * Math.sin(Math.PI * count / divisor));
                    let anchor: string;
                    let onRight = count * 2 > total;
                    let onBottom = false;

                    if (showMinMaxLabelsOnBottom && (count === 0 || count === total - 1)) {
                        // If this is a min or max label and we're showing them on the bottom rather than the sides
                        // Adjust the label display properties to appear under the arc
                        onBottom = true;
                        y += padding / 2;

                        // Align the labels with the outer edge of the arc
                        anchor = onRight ? 'end' : 'start';
                    }
                    else {
                        // For all other labels, display around the arc
                        anchor = onRight ? 'start' : 'end';
                        x += padding * (onRight ? 1 : -1);
                    }

                    let text = this.mainGraphicsContext
                        .append('text')
                        .attr({
                            'x': x,
                            'y': y,
                            'dy': onBottom ? fontSize : 0,
                            'class': Gauge.LabelText.class
                        })
                        .style({
                            'text-anchor': anchor,
                            'font-size': fontSize
                        })
                        .text(ticks[count]);

                    if (!onBottom)
                        this.truncateTextIfNeeded(text, x, onRight);
                }
            }
        }

        private truncateTextIfNeeded(text: D3.Selection, positionX: number, onRight: boolean) {
            let availableSpace = (onRight ? this.currentViewport.width - positionX : positionX);
            text.call(AxisHelper.LabelLayoutStrategy.clip,
                availableSpace,
                TextMeasurementService.svgEllipsis);
        }

        private appendTargetTextAlongArc(radius: number, height: number, width: number, margin: IMargin) {
            let targetSettings = this.targetSettings;

            let target = targetSettings.target;
            let tRatio = (target - targetSettings.min) / (targetSettings.max - targetSettings.min);
            let top = (radius + (height - radius) / 2 + margin.top);
            let flag = tRatio > 0.5;
            let padding = this.settings.labels.padding;

            let anchor = flag ? 'start' : 'end';
            let formatter = valueFormatter.create({
                format: valueFormatter.getFormatString(this.data.metadataColumn, Gauge.formatStringProp),
                value: targetSettings.min,
                value2: targetSettings.max,
                formatSingleValues: true,
                allowFormatBeautification: true,
            });

            let maxRatio = Math.asin(Gauge.MinDistanceFromBottom / radius) / Math.PI;

            let finalRatio = tRatio < maxRatio || tRatio > (1 - maxRatio)
                ? flag
                    ? 1 - maxRatio
                    : maxRatio
                : tRatio;

            let targetX = (margin.left + width / 2) - ((radius + padding) * Math.cos(Math.PI * finalRatio));
            let targetY = top - ((radius + padding) * Math.sin(Math.PI * finalRatio));

            if (!this.targetText) {
                this.targetText = this.mainGraphicsContext
                    .append('text')
                    .classed(Gauge.TargetText.class, true);
            }

            this.targetText
                .attr({
                    'x': targetX,
                    'y': targetY,
                })
                .style({
                    'text-anchor': anchor,
                    'display': this.showTargetLabel ? '' : 'none',
                    'font-size': this.style.labelText.fontSize
                })
                .text(formatter.format(target));

            this.truncateTextIfNeeded(this.targetText, targetX, flag);

            if (!this.targetConnector) {
                this.targetConnector = this.mainGraphicsContext
                    .append('line')
                    .classed(Gauge.TargetConnector.class, true);
            }

            // Hide the target connector if the text is going to align with the target line in the arc
            // It should only be shown if the target text is displaced (ex. when the target is very close to min/max)
            if (tRatio === finalRatio) {
                this.targetConnector.style('display', 'none');
            }
            else {
                this.targetConnector
                    .attr({
                        'x1': (margin.left + width / 2) - (radius * Math.cos(Math.PI * tRatio)),
                        'y1': top - (radius * Math.sin(Math.PI * tRatio)),
                        'x2': targetX,
                        'y2': targetY
                    })
                    .style({
                        'stroke-width': this.settings.targetLine.thickness,
                        'stroke': this.settings.targetLine.color,
                        'display': ''
                    });
            }
        }

        private arcTween(transition, arr): void {
            transition.attrTween('d', (d) => {
                let interpolate = d3.interpolate(d.endAngle, arr[0]);
                return (t) => {
                    d.endAngle = interpolate(t);
                    return arr[1](d);
                };
            });
        }

        private showMinMaxLabelsOnBottom(): boolean {
            // We want to show the start/end ticks on the bottom when there is more vertical space than horizontal
            // and we aren't displaying other ticks on the side that will use the horizontal space anyway
            return this.currentViewport.height > this.currentViewport.width && this.settings.labels.count <= 3;
        }

        private setMargins(): void {
            if (this.gaugeSmallViewPortProperties) {
                if (this.gaugeSmallViewPortProperties.smallGaugeMarginsOnSmallViewPort && (this.currentViewport.height < this.gaugeSmallViewPortProperties.MinHeightGaugeSideNumbersVisible)) {
                    let margins = this.gaugeSmallViewPortProperties.GaugeMarginsOnSmallViewPort;
                    this.margin = { top: margins, bottom: margins, left: margins, right: margins };
                    return;
                }
            }

            this.margin = {
                top: Gauge.DefaultTopBottomMargin,
                bottom: Gauge.DefaultTopBottomMargin,
                left: Gauge.DefaultLeftRightMargin,
                right: Gauge.DefaultLeftRightMargin
            };

            // If we're not showing side labels, reduce the margin so that the gauge has more room to display
            if (!this.showSideNumbersLabelText() || this.showMinMaxLabelsOnBottom()) {
                let targetSettings = this.targetSettings;

                if (this.showTargetLabel) {
                    // If we're showing the target label, only reduce the margin on the side that doesn't have a target label
                    let tRatio = (targetSettings.target - targetSettings.min) / (targetSettings.max - targetSettings.min);

                    if (tRatio > 0.5)
                        this.margin.left = Gauge.ReducedLeftRightMargin;
                    else
                        this.margin.right = Gauge.ReducedLeftRightMargin;
                }
                else {
                    // Otherwise, reduce both margins
                    this.margin.left = this.margin.right = Gauge.ReducedLeftRightMargin;
                }
            }
        }

        private showSideNumbersLabelText(): boolean {
            if (this.gaugeSmallViewPortProperties) {
                if (this.gaugeSmallViewPortProperties.hideGaugeSideNumbersOnSmallViewPort) {
                    if (this.currentViewport.height < this.gaugeSmallViewPortProperties.MinHeightGaugeSideNumbersVisible) {
                        return false;
                    }
                }
            }
            return true;
        }
    }
}
