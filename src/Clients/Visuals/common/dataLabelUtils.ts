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

    export const enum PointLabelPosition {
        Above,
        Below,
    }

    export interface PointDataLabelsSettings extends VisualDataLabelsSettings {
        position: PointLabelPosition;
    }

    export interface VisualDataLabelsSettings {
        show: boolean;
        displayUnits?: number;
        showCategory?: boolean;
        position?: any;
        precision?: number;
        labelColor: string;
    }

    export interface VisualDataLabelsSettingsOptions {
        show: boolean;
        enumeration: ObjectEnumerationBuilder;
        dataLabelsSettings: VisualDataLabelsSettings;
        displayUnits?: boolean;
        precision?: boolean;
        position?: boolean;
        positionObject?: string[];
        selector?: powerbi.data.Selector;
    }

    export interface LabelEnabledDataPoint {
        //for collistion detection use
        labelX?: number;
        labelY?: number;
        //for overriding color from label settings
        labelFill?: string;
        //for display units and precision
        labeltext?: string;
        //taken from column metadata
        labelFormatString?: string;
        isLabelInside?: boolean;
    }

    export interface IColumnFormatterCache {
        [column: string]: IValueFormatter;
        defaultFormatter?: IValueFormatter;
    }

    export interface IColumnFormatterCacheManager {
        cache: IColumnFormatterCache;
        getOrCreate: (formatString: string, labelSetting: VisualDataLabelsSettings, value2?: number) => IValueFormatter;
    }

    export interface LabelPosition {
        y: (d: any, i: number) => number;
        x: (d: any, i: number) => number;
    }

    export interface ILabelLayout {
        labelText: (d: any) => string;
        labelLayout: LabelPosition;
        filter: (d: any) => boolean;
        style: {};
    }

    export interface DataLabelObject extends DataViewObject {
        show: boolean;
        color: Fill;
        labelDisplayUnits: number;
        labelPrecision?: number;
        labelPosition: any;
    }

    export module dataLabelUtils {

        export var labelMargin: number = 8;
        export var maxLabelWidth: number = 50;
        export var defaultColumnLabelMargin: number = 5;
        export var defaultColumnHalfLabelHeight: number = 4;
        export var LabelTextProperties: TextProperties = {
            fontFamily: 'wf_standard-font',
            fontSize: '12px',
            fontWeight: 'normal',
        };
        export var defaultLabelColor = "#777777";
        export var defaultInsideLabelColor = "#ffffff"; //white
        export var hundredPercentFormat = "0.00 %;-0.00 %;0.00 %";

        const defaultDecimalLabelPrecision: number = 2;        
        const defaultCountLabelPrecision: number = 0;   

        let labelGraphicsContextClass: ClassAndSelector = {
            class: 'labels',
            selector: '.labels',
        };

        let linesGraphicsContextClass: ClassAndSelector = {
            class: 'lines',
            selector: '.lines',
        };

        let labelsClass: ClassAndSelector = {
            class: 'data-labels',
            selector: '.data-labels',
        };

        let lineClass: ClassAndSelector = {
            class: 'line-label',
            selector: '.line-label',
        };

        export function updateLabelSettingsFromLabelsObject(labelsObj: DataLabelObject, labelSettings: VisualDataLabelsSettings): void {
            if (labelsObj) {
                if (labelsObj.show !== undefined)
                    labelSettings.show = labelsObj.show;
                if (labelsObj.color !== undefined) {
                    labelSettings.labelColor = labelsObj.color.solid.color;
                }
                if (labelsObj.labelDisplayUnits !== undefined) {
                    labelSettings.displayUnits = labelsObj.labelDisplayUnits;
                }
                if (labelsObj.labelPrecision !== undefined) {
                    labelSettings.precision = (labelsObj.labelPrecision >= 0) ? labelsObj.labelPrecision : 0;
                }
            }
        }
        
        export function getDefaultLabelSettings(show: boolean = false, labelColor?: string, format?: string): VisualDataLabelsSettings {
            return {
                show: show,
                position: PointLabelPosition.Above,
                displayUnits: 0,
                precision: getPrecision(format),
                labelColor: labelColor || defaultLabelColor,
                formatterOptions: null,
            };
        }

        export function getDefaultTreemapLabelSettings(format?: string): VisualDataLabelsSettings {
            return {
                show: false,
                displayUnits: 0,
                precision: getPrecision(format),
                labelColor: defaultInsideLabelColor,
                showCategory: true,
                formatterOptions: null,
            };
        }

        export function getDefaultColumnLabelSettings(isLabelPositionInside: boolean, format?: string): VisualDataLabelsSettings {
            var labelSettings = getDefaultLabelSettings(false, undefined, format);
            labelSettings.position = null;
            labelSettings.labelColor = isLabelPositionInside ? defaultInsideLabelColor : defaultLabelColor;
            return labelSettings;
        }

        export function getDefaultPointLabelSettings(format?: string): PointDataLabelsSettings {
            return {
                show: false,
                position: PointLabelPosition.Above,
                displayUnits: 0,
                precision: getPrecision(format),
                labelColor: defaultLabelColor,
                formatterOptions: null,
            };
        }

        export function getDefaultMapLabelSettings(format?: string): PointDataLabelsSettings {
            return {
                show: false,
                position: PointLabelPosition.Above,
                displayUnits: 0,
                precision: getPrecision(format),
                labelColor: defaultInsideLabelColor,
                formatterOptions: null,
            };
        }

        export function getDefaultDonutLabelSettings(format?: string): VisualDataLabelsSettings {
            return {
                show: false,
                displayUnits: 0,
                precision: getPrecision(format),
                labelColor: defaultLabelColor,
                position: null,
                showCategory: true,
                formatterOptions: null,
            };
        }

        export function getDefaultFunnelLabelSettings(format?: string): VisualDataLabelsSettings {
            return {
                show: true,
                position: powerbi.labelPosition.insideCenter,
                displayUnits: 0,
                precision: getPrecision(format),
                labelColor: defaultLabelColor,
                formatterOptions: null,
            };
        }

        function getPrecision(format: string): number {
            debug.assertAnyValue(format, 'format');

            if (format) {
                let formatMetadata = NumberFormat.getCustomFormatMetadata(format);
                if (formatMetadata.hasDots) {
                    return defaultDecimalLabelPrecision;
                }
            }

            return defaultCountLabelPrecision;
        }

        export function drawDefaultLabelsForDataPointChart(data: any[], context: D3.Selection, layout: ILabelLayout,
            viewport: IViewport, isAnimator: boolean = false, animationDuration?: number, hasSelection?: boolean): D3.UpdateSelection {
            debug.assertValue(data, 'data cannot be null or undefined');

            // Hide and reposition labels that overlap
            let dataLabelManager = new DataLabelManager();
            let filteredData = dataLabelManager.hideCollidedLabels(viewport, data, layout);

            let hasAnimation = isAnimator && !!animationDuration;
            let labels = selectLabels(filteredData, context, false, hasAnimation);

            if (!labels)
                return;
            
            if (hasAnimation) {
                labels
                    .text((d: LabelEnabledDataPoint) => d.labeltext)
                    .transition()
                    .duration(animationDuration)
                    .style(layout.style)
                    .style('opacity', hasSelection ? (d: SelectableDataPoint) => ColumnUtil.getFillOpacity(d.selected, false, hasSelection, false) : 1)
                    .attr({ x: (d: LabelEnabledDataPoint) => d.labelX, y: (d: LabelEnabledDataPoint) => d.labelY });

                labels
                    .exit()
                    .transition()
                    .duration(animationDuration)
                    .style('opacity', 0) //fade out labels that are removed
                    .remove();
            }
            else {
                labels
                    .attr({ x: (d: LabelEnabledDataPoint) => d.labelX, y: (d: LabelEnabledDataPoint) => d.labelY })
                    .text((d: LabelEnabledDataPoint) => d.labeltext)
                    .style(layout.style);

                labels
                    .exit()
                    .remove();
            }

            return labels;
        }
        
        /**
         * Note: Funnel chart uses animation and does not use collision detection.
         */
        export function drawDefaultLabelsForFunnelChart(data: FunnelSlice[], context: D3.Selection, layout: ILabelLayout, isAnimator: boolean = false, animationDuration?: number): D3.UpdateSelection {
            debug.assertValue(data, 'data could not be null or undefined');

            let filteredData = data.filter(layout.filter);

            let labels = selectLabels(filteredData, context);

            if (!labels)
                return;

            labels
                .attr(layout.labelLayout)
                .text(layout.labelText)
                .style(layout.style);

            if (isAnimator && animationDuration) {
                labels.transition().duration(animationDuration);
            }

            labels
                .exit()
                .remove();

            return labels;
        }

        export function drawDefaultLabelsForDonutChart(data: any[], context: D3.Selection, layout: ILabelLayout, viewport: IViewport, radius: number, arc: D3.Svg.Arc, outerArc: D3.Svg.Arc) {
            debug.assertValue(data, 'data could not be null or undefined');

            // Hide and reposition labels that overlap
            let dataLabelManager = new DataLabelManager();
            let filteredData = dataLabelManager.hideCollidedLabels(viewport, data, layout,/* addTransform */ true);

            let labels = selectLabels(filteredData, context, true);

            if (!labels)
                return;
            
            labels
                .attr({ x: (d: LabelEnabledDataPoint) => d.labelX, y: (d: LabelEnabledDataPoint) => d.labelY, dy: '.35em' })
                .text((d: LabelEnabledDataPoint) => d.labeltext)
                .style(layout.style);

            labels
                .exit()
                .remove();

            if (context.select(linesGraphicsContextClass.selector).empty())
                context.append('g').classed(linesGraphicsContextClass.class, true);

            let lines = context.select(linesGraphicsContextClass.selector).selectAll('polyline')
                .data(filteredData, (d: DonutArcDescriptor) => d.data.identity.getKey());
            let innerLinePointMultiplier = 2.05;

            let midAngle = function (d: DonutArcDescriptor) { return d.startAngle + (d.endAngle - d.startAngle) / 2; };

            lines.enter()
                .append('polyline')
                .classed(lineClass.class, true);

            lines
                .attr('points', function (d) {
                    let textPoint = outerArc.centroid(d);
                    textPoint[0] = radius * 0.95 * (midAngle(d) < Math.PI ? 1 : -1);
                    let midPoint = outerArc.centroid(d);
                    let chartPoint = arc.centroid(d);
                    chartPoint[0] *= innerLinePointMultiplier;
                    chartPoint[1] *= innerLinePointMultiplier;
                    return [chartPoint, midPoint, textPoint];
                }).
                style({
                    'opacity': (d: DonutArcDescriptor) => DonutChart.PolylineOpacity,
                    'stroke': (d: DonutArcDescriptor) => d.data.labelColor,
                });

                lines
                    .exit()
                    .remove();
        }

        function selectLabels(filteredData: LabelEnabledDataPoint[], context: D3.Selection, isDonut: boolean = false, forAnimation: boolean = false): D3.UpdateSelection {

            // Check for a case where resizing leaves no labels - then we need to remove the labels 'g'
            if (filteredData.length === 0) {
                cleanDataLabels(context, true);
                return null;
            }

            if (context.select(labelGraphicsContextClass.selector).empty())
                context.append('g').classed(labelGraphicsContextClass.class, true);

            // line chart ViewModel has a special 'key' property for point identification since the 'identity' field is set to the series identity
            let hasKey: boolean = (<any>filteredData)[0].key != null;
            let hasDataPointIdentity: boolean = (<any>filteredData)[0].identity != null;
            let getIdentifier = hasKey ?
                (d: any) => d.key
                : hasDataPointIdentity ?
                    (d: SelectableDataPoint) => d.identity.getKey()
                    : undefined;

            let labels = isDonut ?
                context.select(labelGraphicsContextClass.selector).selectAll(labelsClass.selector).data(filteredData, (d: DonutArcDescriptor) => d.data.identity.getKey())
                : getIdentifier != null ?
                    context.select(labelGraphicsContextClass.selector).selectAll(labelsClass.selector).data(filteredData, getIdentifier)
                    : context.select(labelGraphicsContextClass.selector).selectAll(labelsClass.selector).data(filteredData);

            let newLabels = labels.enter()
                .append('text')
                .classed(labelsClass.class, true);
            if (forAnimation)
                newLabels.style('opacity', 0);

            return labels;
        }

        export function cleanDataLabels(context: D3.Selection, removeLines: boolean = false) {
            let empty = [];
            let labels = context.selectAll(labelsClass.selector).data(empty);
            labels.exit().remove();
            context.selectAll(labelGraphicsContextClass.selector).remove();
            if (removeLines) {
                let lines = context.selectAll(lineClass.selector).data(empty);
                lines.exit().remove();
                context.selectAll(linesGraphicsContextClass.selector).remove();
            }
        }

        export function setHighlightedLabelsOpacity(context: D3.Selection, hasSelection: boolean, hasHighlights: boolean) {
            context.selectAll(labelsClass.selector).style("fill-opacity", (d: ColumnChartDataPoint) => {
                let labelOpacity = ColumnUtil.getFillOpacity(d.selected, d.highlight, !d.highlight && hasSelection, !d.selected && hasHighlights) < 1 ? 0 : 1;
                return labelOpacity;
            });
        }

        export function getLabelFormattedText(label: string | number, maxWidth?: number, format?: string, formatter?: IValueFormatter): string {
            let properties: TextProperties = {
                text: formatter
                    ? formatter.format(label)
                    : formattingService.formatValue(label, format),
                fontFamily: LabelTextProperties.fontFamily,
                fontSize: LabelTextProperties.fontSize,
                fontWeight: LabelTextProperties.fontWeight,
            };
            maxWidth = maxWidth ? maxWidth : maxLabelWidth;

            return TextMeasurementService.getTailoredTextOrDefault(properties, maxWidth);
        }

        export function getLabelLayoutXYForWaterfall(xAxisProperties: IAxisProperties, categoryWidth: number, yAxisProperties: IAxisProperties, dataDomain: number[]): LabelPosition {
            return {
                x: (d: WaterfallChartDataPoint) => xAxisProperties.scale(d.categoryIndex) + (categoryWidth / 2),
                y: (d: WaterfallChartDataPoint) => getWaterfallLabelYPosition(yAxisProperties.scale, d, dataDomain)
            };
        }

        function getWaterfallLabelYPosition(scale: D3.Scale.GenericScale<any>, d: WaterfallChartDataPoint, dataDomain: number[]): number {

            let yValue = scale(0) - scale(Math.abs(d.value));
            let yPos = scale(d.position);
            let scaleMinDomain = scale(dataDomain[0]);
            let endPosition = scale(d.position + d.value);

            if (d.value < 0) {
                let properties: TextProperties = {
                    text: d.labeltext,
                    fontFamily: dataLabelUtils.LabelTextProperties.fontFamily,
                    fontSize: dataLabelUtils.LabelTextProperties.fontSize,
                    fontWeight: dataLabelUtils.LabelTextProperties.fontWeight,
                };
                let outsideBelowPosition = yPos + yValue + TextMeasurementService.estimateSvgTextHeight(properties);
                // Try to honor the position, but if the label doesn't fit where specified, then swap the position.
                if (scaleMinDomain > outsideBelowPosition) {
                    return outsideBelowPosition;
                }
            }
            else {
                let outsideAbovePosition = yPos - yValue - dataLabelUtils.labelMargin;
                // Try to honor the position, but if the label doesn't fit where specified, then swap the position.
                if (outsideAbovePosition > 0) {
                    return outsideAbovePosition;
                }
            }
            d.isLabelInside = true;
            return getWaterfallInsideLabelYPosition(yPos, endPosition, scaleMinDomain);
        }

        function getWaterfallInsideLabelYPosition(startPosition: number, endPosition: number, scaleMinDomain: number): number {
            // Get the start and end position of the column
            // If the start or end is outside of the visual because of clipping - adjust the position
            startPosition = startPosition < 0 ? 0 : startPosition;
            startPosition = startPosition > scaleMinDomain ? scaleMinDomain : startPosition;

            endPosition = endPosition < 0 ? 0 : endPosition;
            endPosition = endPosition > scaleMinDomain ? scaleMinDomain : endPosition;

            return (Math.abs(endPosition - startPosition) / 2) + Math.min(startPosition, endPosition);
        }

        export function doesDataLabelFitInShape(d: WaterfallChartDataPoint, yAxisProperties: IAxisProperties, layout: WaterfallLayout): boolean {

            if (d == null || d.value === null)
                return false;

            let properties: TextProperties = {
                text: layout.labelText(d),
                fontFamily: dataLabelUtils.LabelTextProperties.fontFamily,
                fontSize: dataLabelUtils.LabelTextProperties.fontSize,
                fontWeight: dataLabelUtils.LabelTextProperties.fontWeight,
            };

            let outsidePosition = WaterfallChart.getRectTop(yAxisProperties.scale, d.position, d.value) - dataLabelUtils.labelMargin;

            // The shape is fit to be outside
            if (outsidePosition > 0)
                return true;

            let textWidth = TextMeasurementService.measureSvgTextWidth(properties);
            let textHeight = TextMeasurementService.estimateSvgTextHeight(properties);

            let shapeWidth = layout.categoryWidth;
            let shapeHeight = Math.abs(AxisHelper.diffScaled(yAxisProperties.scale, Math.max(0, Math.abs(d.value)), 0));

            //checking that labels aren't greater than shape
            if ((textWidth > shapeWidth) || (textHeight > shapeHeight))
                return false;
            return true;
        }

        export function getMapLabelLayout(labelSettings: PointDataLabelsSettings): ILabelLayout {
            
            return {
                labelText: (d: MapVisualDataPoint) => {
                    return getLabelFormattedText(d.labeltext);
                },
                labelLayout: {
                    x: (d: MapVisualDataPoint) => d.x,
                    y: (d: MapVisualDataPoint) => {
                        let margin = d.radius + labelMargin;
                        return labelSettings.position === PointLabelPosition.Above ? d.y - margin : d.y + margin;
                    },
                },
                filter: (d: MapVisualDataPoint) => {
                    return (d != null && d.labeltext != null);
                },
                style: {
                    'fill': (d: MapVisualDataPoint) => d.labelFill,
                },
            };
        }

        export function getColumnChartLabelLayout(
            data: ColumnChartData,
            labelLayoutXY: any,
            isColumn: boolean,
            isHundredPercent: boolean,
            axisFormatter: IValueFormatter,
            axisOptions: ColumnAxisOptions,
            interactivityService: IInteractivityService,
            visualWidth?: number): ILabelLayout {

            let formatOverride: string = (isHundredPercent) ? hundredPercentFormat : null;
            let formattersCache = createColumnFormatterCacheManager();
            let hasSelection = interactivityService ? interactivityService.hasSelection() : false;

            return {
                labelText: (d: ColumnChartDataPoint) => {
                    let formatString = (formatOverride != null) ? formatOverride : d.labelFormatString;
                    let value2: number = getDisplayUnitValueFromAxisFormatter(axisFormatter, d.labelSettings);
                    let formatter = formattersCache.getOrCreate(formatString, d.labelSettings, value2);
                    return getLabelFormattedText(formatter.format(d.value), maxLabelWidth);
                },
                labelLayout: labelLayoutXY,
                filter: (d: ColumnChartDataPoint) => dataLabelUtils.getColumnChartLabelFilter(d, hasSelection, data.hasHighlights, axisOptions, visualWidth),
                style: {
                    'fill': (d: ColumnChartDataPoint) => d.labelFill,
                    'text-anchor': isColumn ? 'middle' : 'start',
                },
            };
        }
        
        /**
         * Valide for stacked column/bar chart and 100% stacked column/bar chart,
         * that labels that should to be inside the shape aren't bigger then shapes.
         */
        function validateLabelsSize(d: ColumnChartDataPoint, axisOptions: ColumnAxisOptions, visualWidth?: number): boolean {
            let xScale = axisOptions.xScale;
            let yScale = axisOptions.yScale;
            let columnWidth = axisOptions.columnWidth;
            let properties: TextProperties = {
                text: d.labeltext,
                fontFamily: dataLabelUtils.LabelTextProperties.fontFamily,
                fontSize: dataLabelUtils.LabelTextProperties.fontSize,
                fontWeight: dataLabelUtils.LabelTextProperties.fontWeight,
            };
            let textWidth = TextMeasurementService.measureSvgTextWidth(properties);
            let textHeight = TextMeasurementService.estimateSvgTextHeight(properties);
            let shapeWidth, shapeHeight;
            let inside = false;
            let outsidePosition: number = ColumnUtil.calculatePosition(d, axisOptions);
            switch (d.chartType) {
                case ColumnChartType.stackedBar:
                case ColumnChartType.hundredPercentStackedBar:
                    // if the series isn't last or the label doesn't fit where specified, then it should be inside 
                    if (!d.lastSeries || (outsidePosition + textWidth > visualWidth) || d.chartType === ColumnChartType.hundredPercentStackedBar) {
                        shapeWidth = -StackedUtil.getSize(xScale, d.valueAbsolute);
                        shapeHeight = columnWidth;
                        inside = true;
                    }
                    break;
                case ColumnChartType.clusteredBar:
                   
                    // if the label doesn't fit where specified, then it should be inside 
                    if ((outsidePosition + textWidth) > visualWidth) {
                        shapeWidth = Math.abs(AxisHelper.diffScaled(xScale, 0, d.value));
                        shapeHeight = columnWidth;
                        inside = true;
                    }
                    break;
                case ColumnChartType.stackedColumn:
                case ColumnChartType.hundredPercentStackedColumn:
                    // if the series isn't last or the label doesn't fit where specified, then it should be inside 
                    if (!d.lastSeries || outsidePosition <= 0 || d.chartType === ColumnChartType.hundredPercentStackedColumn) {
                        shapeWidth = columnWidth;
                        shapeHeight = StackedUtil.getSize(yScale, d.valueAbsolute);
                        inside = true;
                    }
                    break;
                case ColumnChartType.clusteredColumn:
                    // if the label doesn't fit where specified, then it should be inside 
                    if (outsidePosition <= 0) {
                        shapeWidth = columnWidth;
                        shapeHeight = Math.abs(AxisHelper.diffScaled(yScale, 0, d.value));
                        inside = true;
                    }
                    break;
                default:
                    return true;
            }

            //checking that labels aren't greater than shape
            if (inside && ((textWidth > shapeWidth) || textHeight > shapeHeight)) return false;
            return true;
        }

        export function getColumnChartLabelFilter(d: ColumnChartDataPoint, hasSelection: boolean, hasHighlights: boolean, axisOptions: ColumnAxisOptions, visualWidth?: number): any {
                //labels of dimmed are hidden
                let shapesOpacity = hasSelection ? ColumnUtil.getFillOpacity(d.selected, d.highlight, !d.highlight && hasSelection, !d.selected && hasHighlights) :
                    ColumnUtil.getFillOpacity(d.selected, d.highlight, hasSelection, hasHighlights);
                return (d != null && d.value != null && validateLabelsSize(d, axisOptions, visualWidth) && shapesOpacity === 1);
        }

        export function getScatterChartLabelLayout(xScale: D3.Scale.GenericScale<any>, yScale: D3.Scale.GenericScale<any>, labelSettings: PointDataLabelsSettings, viewport: IViewport, sizeRange: NumberRange): ILabelLayout {

            return {
                labelText: (d: ScatterChartDataPoint) => {
                    return getLabelFormattedText(d.category, maxLabelWidth * 2.0);
                },
                labelLayout: {
                    x: (d: ScatterChartDataPoint) => xScale(d.x),
                    y: (d: ScatterChartDataPoint) => {
                        let margin = ScatterChart.getBubbleRadius(d.radius, sizeRange, viewport) + labelMargin;
                        return labelSettings.position === PointLabelPosition.Above ? yScale(d.y) - margin : yScale(d.y) + margin;
                },
                },
                filter: (d: ScatterChartDataPoint) => {
                    return (d != null && d.category != null);
                },
                style: {
                    'fill': (d: ScatterChartDataPoint) => d.labelFill,
                },
            };
        }

        export function getLineChartLabelLayout(xScale: D3.Scale.GenericScale<any>, yScale: D3.Scale.GenericScale<any>, labelSettings: PointDataLabelsSettings, isScalar: boolean, axisFormatter: IValueFormatter): ILabelLayout {
            let formattersCache = createColumnFormatterCacheManager();

            return {
                labelText: (d: LineChartDataPoint) => {
                    let value2: number = getDisplayUnitValueFromAxisFormatter(axisFormatter, d.labelSettings);
                    let formatter = formattersCache.getOrCreate(d.labelFormatString, d.labelSettings, value2);
                    return getLabelFormattedText(formatter.format(d.value));
                },
                labelLayout: {
                    x: (d: LineChartDataPoint) => xScale(isScalar ? d.categoryValue : d.categoryIndex),
                    y: (d: LineChartDataPoint) => { return labelSettings.position === PointLabelPosition.Above ? yScale(d.value) - labelMargin : yScale(d.value) + labelMargin; },
                },
                filter: (d: LineChartDataPoint) => {
                    return (d != null && d.value != null);
                },
                style: {
                    'fill': (d: LineChartDataPoint) => d.labelFill,
                },
            };
        }

        export function getDonutChartLabelLayout(labelSettings: VisualDataLabelsSettings, radius: number, outerArc: D3.Svg.Arc, viewport: IViewport, value2: number): ILabelLayout {

            let midAngle = function (d: DonutArcDescriptor) { return d.startAngle + (d.endAngle - d.startAngle) / 2; };
            let spaceAvaliableForLabels = viewport.width / 2 - radius;
            let minAvailableSpace = Math.min(spaceAvaliableForLabels, maxLabelWidth);
            let measureFormattersCache = dataLabelUtils.createColumnFormatterCacheManager();

            return {
                labelText: (d: DonutArcDescriptor) => {
                    if (labelSettings.show) {
                        let measureFormatter = measureFormattersCache.getOrCreate(d.data.labelFormatString, labelSettings, value2);
                        return labelSettings.showCategory
                            ? getLabelFormattedText(getLabelFormattedText(d.data.label, minAvailableSpace) + " (" + measureFormatter.format(d.data.measure) + ")", spaceAvaliableForLabels)
                            : getLabelFormattedText(d.data.measure, minAvailableSpace,/* format */ null, measureFormatter);
                    }
                    // show only category label
                    return getLabelFormattedText(d.data.label, minAvailableSpace);
                },
                labelLayout: {
                    x: (d: DonutArcDescriptor) => {
                        return radius * (midAngle(d) < Math.PI ? 1 : -1);
                    },
                    y: (d: DonutArcDescriptor) => {
                        let pos = outerArc.centroid(d);
                        return pos[1];
                    },
                },
                filter: (d: DonutArcDescriptor) => (d != null && d.data != null && d.data.label != null),
                style: {
                    'fill': (d: DonutArcDescriptor) => d.data.labelColor,
                    'text-anchor': (d: DonutArcDescriptor) => midAngle(d) < Math.PI ? 'start' : 'end',
                },
            };
        }

        export function getFunnelChartLabelLayout(
            data: FunnelData,
            axisOptions: FunnelAxisOptions, innerTextHeightDelta: number,
            textMinimumPadding: number,
            labelSettings: VisualDataLabelsSettings,
            currentViewport: IViewport): ILabelLayout {

            let yScale = axisOptions.yScale;
            let xScale = axisOptions.xScale;
            let marginLeft = axisOptions.margin.left;

            //the bars are tranform, verticalRange mean horizontal range, xScale is y, yscale is x
            let halfRangeBandPlusDelta = axisOptions.xScale.rangeBand() / 2 + innerTextHeightDelta;
            let pixelSpan = axisOptions.verticalRange / 2;
            let formatString = valueFormatter.getFormatString(data.valuesMetadata[0], funnelChartProps.general.formatString);
            let textMeasurer: ITextAsSVGMeasurer = TextMeasurementService.measureSvgTextWidth;

            let value2: number = null;
            if (labelSettings.displayUnits === 0) {
                let minY = <number>d3.min(data.slices, (d) => { return d.value; });
                let maxY = <number>d3.max(data.slices, (d) => { return d.value; });
                value2 = Math.max(Math.abs(minY), Math.abs(maxY));
            }

            let formattersCache = createColumnFormatterCacheManager();

            return {
                labelText: (d: FunnelSlice) => {
                    let barWidth = Math.abs(yScale(d.value) - yScale(0));
                    let insideAvailableSpace = Math.abs(yScale(d.value) - yScale(0)) - (textMinimumPadding * 2);
                    let outsideAvailableSpace = pixelSpan - (barWidth / 2) - textMinimumPadding;
                    let labelFormatString = (formatString != null) ? formatString : d.labelFormatString;

                    let maximumTextSize = Math.max(insideAvailableSpace, outsideAvailableSpace);
                    let formatter = formattersCache.getOrCreate(labelFormatString, labelSettings, value2);
                    let labelText = formatter.format(FunnelChart.getFunnelSliceValue(d, true /* asOriginal */));
                    return getLabelFormattedText(labelText, maximumTextSize);
                },
                labelLayout: {
                    y: (d, i) => {
                        return xScale(i) + halfRangeBandPlusDelta;
                    },
                    x: (d: FunnelSlice) => {
                        let barWidth = Math.abs(yScale(d.value) - yScale(0));
                        let insideAvailableSpace = Math.abs(yScale(d.value) - yScale(0)) - (textMinimumPadding * 2);
                        let outsideAvailableSpace = pixelSpan - (barWidth / 2) - textMinimumPadding;
                        let maximumTextSize = Math.max(insideAvailableSpace, outsideAvailableSpace);
                        let labelFormatString = (formatString != null) ? formatString : d.labelFormatString;

                        let formatter = formattersCache.getOrCreate(labelFormatString, labelSettings, value2);
                        let labelText = formatter.format(FunnelChart.getFunnelSliceValue(d, true /* asOriginal */));
                        let properties: TextProperties = {
                            text: getLabelFormattedText(labelText, maximumTextSize),
                            fontFamily: LabelTextProperties.fontFamily,
                            fontSize: LabelTextProperties.fontSize,
                            fontWeight: LabelTextProperties.fontWeight,
                        };

                        let textLength = textMeasurer(properties);

                        // Try to honor the position, but if the label doesn't fit where specified, then swap the position.
                        let labelPosition = labelSettings.position;
                        if ((labelPosition === powerbi.labelPosition.outsideEnd && outsideAvailableSpace < textLength) || d.value === 0)
                            labelPosition = powerbi.labelPosition.insideCenter;
                        else if (labelPosition === powerbi.labelPosition.insideCenter && insideAvailableSpace < textLength) {
                            labelPosition = powerbi.labelPosition.outsideEnd;
                        }

                        switch (labelPosition) {
                            case powerbi.labelPosition.outsideEnd:
                                return marginLeft + pixelSpan + (barWidth / 2) + textMinimumPadding + (textLength / 2);
                            default:
                                // Inside position, change color to white unless value is 0
                                d.labelFill = d.value !== 0 ? defaultInsideLabelColor : d.labelFill;
                                return marginLeft + pixelSpan;
                        }
                    },
                },
                filter: (d: FunnelSlice) => {
                    return d != null && d.value != null && data.hasHighlights === !!d.highlight;
                },
                style: {
                    'fill': (d: FunnelSlice) => d.labelFill,
                    'fill-opacity': (d: FunnelSlice) => ColumnUtil.getFillOpacity(d.selected, false, false, false),
                },
            };
        }

        export function enumerateDataLabels(
            options: VisualDataLabelsSettingsOptions): ObjectEnumerationBuilder {

            debug.assertValue(options, 'options');
            debug.assertValue(options.enumeration, 'enumeration');

            if (!options.dataLabelsSettings)
                return;

            let instance: VisualObjectInstance = {
                objectName: 'labels',
                selector: options.selector,
                properties: {},
            };

            if (options.show) {
                instance.properties['show'] = options.dataLabelsSettings.show;
            }

            instance.properties['color'] = options.dataLabelsSettings.labelColor;

            if (options.displayUnits) {
                instance.properties['labelDisplayUnits'] = options.dataLabelsSettings.displayUnits;
            }
            if (options.precision) {
                instance.properties['labelPrecision'] = options.dataLabelsSettings.precision;
            }
            if (options.position) {
                instance.properties['labelPosition'] = options.dataLabelsSettings.position;
                if (options.positionObject) {
                    debug.assert(!instance.validValues, '!instance.validValues');

                    instance.validValues = { 'labelPosition': options.positionObject };
                }
            }

            return options.enumeration.pushInstance(instance);
        }

        export function enumerateCategoryLabels(enumeration: ObjectEnumerationBuilder, dataLabelsSettings: VisualDataLabelsSettings, withFill: boolean, isDonutChart: boolean = false, isTreeMap: boolean = false): void {
            let labelSettings = (dataLabelsSettings)
                ? dataLabelsSettings
                : (isDonutChart)
                ? getDefaultDonutLabelSettings()
                : (isTreeMap)
                ? getDefaultTreemapLabelSettings()
                : getDefaultPointLabelSettings();

            let instance: VisualObjectInstance = {
                objectName: 'categoryLabels',
                selector: null,
                properties: {
                    show: isDonutChart || isTreeMap
                    ? labelSettings.showCategory
                    : labelSettings.show,
                },
            };
            
            if (withFill) {
                instance.properties['color'] = labelSettings.labelColor;
            }
            
            enumeration.pushInstance(instance);
        }

        function getDisplayUnitValueFromAxisFormatter(axisFormatter: IValueFormatter, labelSettings: VisualDataLabelsSettings): number {
            if (axisFormatter && axisFormatter.displayUnit && labelSettings.displayUnits === 0)
                return axisFormatter.displayUnit.value;
            return null;
        }

        export function createColumnFormatterCacheManager(): IColumnFormatterCacheManager {
            return <IColumnFormatterCacheManager> {

                cache: { defaultFormatter: null, },
                getOrCreate(formatString: string, labelSetting: VisualDataLabelsSettings, value2?: number) {
                    if (formatString) {
                        var cacheKeyObject = {
                            formatString: formatString,
                            displayUnits: labelSetting.displayUnits,
                            precision: labelSetting.precision,
                            value2: value2
                        };
                        var cacheKey = JSON.stringify(cacheKeyObject);
                        if (!this.cache[cacheKey])
                            this.cache[cacheKey] = valueFormatter.create(getOptionsForLabelFormatter(labelSetting, formatString, value2));
                        return this.cache[cacheKey];
                    }
                    if (!this.cache.defaultFormatter) {
                        this.cache.defaultFormatter = valueFormatter.create(getOptionsForLabelFormatter(labelSetting, formatString, value2));
                    }
                    return this.cache.defaultFormatter;
                }
            };
        }

        function getOptionsForLabelFormatter(labelSetting: VisualDataLabelsSettings, formatString: string, value2?: number): ValueFormatterOptions {
            return {
                displayUnitSystemType: DisplayUnitSystemType.DataLabels,
                format: formatString,
                precision: labelSetting.precision,
                value: labelSetting.displayUnits,
                value2: value2,
                allowFormatBeautification: true,
            };
        }
    }
}
