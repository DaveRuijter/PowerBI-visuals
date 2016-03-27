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

    export module NewDataLabelUtils {
        export let startingLabelOffset = 8;
        export let maxLabelOffset = 8;
        export let maxLabelWidth: number = 50;
        export let hundredPercentFormat = '0.00 %;-0.00 %;0.00 %';
        export let LabelTextProperties: TextProperties = {
            fontFamily: 'wf_standard-font',
            fontSize: '12px',
            fontWeight: 'normal',
        };
        export let defaultLabelColor = "#777777";
        export let defaultInsideLabelColor = "#ffffff"; //white
        let horizontalLabelBackgroundMargin = 4;
        let verticalLabelBackgroundMargin = 2;
        let labelBackgroundRounding = 4;

        export let labelGraphicsContextClass: ClassAndSelector = {
            class: 'labelGraphicsContext',
            selector: '.labelGraphicsContext'
        };

        let labelsClass: ClassAndSelector = {
            class: 'label',
            selector: '.label'
        };

        export function drawDefaultLabels(context: D3.Selection, dataLabels: Label[]): D3.UpdateSelection {
            let labels = context.selectAll(labelsClass.selector)
                .data(_.filter(dataLabels, (d: Label) => d.isVisible), (d: Label, index: number) => { return d.identity ? d.identity.getKeyWithoutHighlight() : index; });

            labels.enter()
                .append("text")
                .classed(labelsClass.class, true);

            labels
                .text((d: Label) => d.text)
                .attr({
                    x: (d: Label) => {
                        return (d.boundingBox.left + (d.boundingBox.width / 2));
                    },
                    y: (d: Label) => {
                        return d.boundingBox.top + d.boundingBox.height;
                    },
                    dy: "-0.15em",
                })
                .style('fill', (d: Label) => d.fill);

            labels.exit()
                .remove();

            return labels;
        }

        export function animateDefaultLabels(context: D3.Selection, dataLabels: Label[], duration: number): D3.UpdateSelection {
            let labels = context.selectAll(labelsClass.selector)
                .data(_.filter(dataLabels, (d: Label) => d.isVisible), (d: Label, index: number) => { return d.identity ? d.identity.getKeyWithoutHighlight() : index; });

            labels.enter()
                .append("text")
                .classed(labelsClass.class, true);

            labels
                .text((d: Label) => d.text)
                .style('fill', (d: Label) => d.fill)
                .transition()
                .duration(duration)
                .attr({
                    x: (d: Label) => {
                        return (d.boundingBox.left + (d.boundingBox.width / 2));
                    },
                    y: (d: Label) => {
                        return d.boundingBox.top + d.boundingBox.height;
                    },
                    dy: "-0.15em",
                });

            labels.exit()
                .remove();

            return labels;
        }

        /** Draws black rectangles based on the bounding bx of labels, to be used in debugging */
        export function drawLabelBackground(context: D3.Selection, dataLabels: Label[], fill?: string, fillOpacity?: number): D3.UpdateSelection {
            let labelRects = context.selectAll("rect")
                .data(_.filter(dataLabels, (d: Label) => d.isVisible));

            labelRects.enter()
                .append("rect");

            labelRects
                .attr({
                    x: (d: Label) => {
                        return d.boundingBox.left - horizontalLabelBackgroundMargin;
                    },
                    y: (d: Label) => {
                        return d.boundingBox.top - verticalLabelBackgroundMargin;
                    },
                    rx: labelBackgroundRounding,
                    ry: labelBackgroundRounding,
                    width: (d: Label) => {
                        return d.boundingBox.width + 2 * horizontalLabelBackgroundMargin;
                    },
                    height: (d: Label) => {
                        return d.boundingBox.height + 2 * verticalLabelBackgroundMargin;
                    },
                })
                .style("fill", fill ? fill : "#000000")
                .style("fill-opacity", fillOpacity != null ? fillOpacity : 1);

            labelRects.exit()
                .remove();

            return labelRects;
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

        export function getDisplayUnitValueFromAxisFormatter(axisFormatter: IValueFormatter, labelSettings: VisualDataLabelsSettings): number {
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

        export function removeDuplicates(labelDataPoints: LabelDataPoint[]): LabelDataPoint[]{
            let uniqueLabelDataPoints: LabelDataPoint[] = [];
            let labelDataPointMap = {};
            let sameParentIsInArray = (newValue: any, array: any[], parentIsRect: boolean) => {
                return array.some((arrayValue) => {
                    if (parentIsRect) {
                        return shapes.Rect.equals(newValue.parentShape.rect, arrayValue.rect);
                    }
                    else {
                        return shapes.Point.equals(newValue.parentShape.point, arrayValue.point);
                    }
                });
            };
            for (let dataPoint of labelDataPoints) {
                let parentIsRect = dataPoint.isParentRect;
                let resultsFromMap = labelDataPointMap[dataPoint.text];
                if (!resultsFromMap) {
                    uniqueLabelDataPoints.push(dataPoint);
                    labelDataPointMap[dataPoint.text] = [dataPoint.parentShape];
                }
                else {
                    if (!sameParentIsInArray(dataPoint, resultsFromMap, parentIsRect)) {
                        uniqueLabelDataPoints.push(dataPoint);
                        resultsFromMap.push(dataPoint.parentShape);
                    }
                }
            }
            return uniqueLabelDataPoints;
        }
    }
}
