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

/// <reference path="../../_references.ts"/>

module powerbi.visuals.samples {
    import SelectionManager = utility.SelectionManager;
    import ClassAndSelector = jsCommon.CssConstants.ClassAndSelector;

    export interface DotPlotConstructorOptions {
        animator?: IGenericAnimator;
        svg?: D3.Selection;
        margin?: IMargin;
        radius?: number;
        strokeWidth?: number;
    }
    
    export interface DotPlotDatapoint {
        x: number;
        y: number;
        color?: string;
        value?: number;
        label?: string;
        identity: SelectionId;
        tooltipInfo?: TooltipDataItem[];
    }
    
    export interface DotPlotData {
        dataPoints: DotPlotDatapoint[];
        legendData: LegendData;
    }

    export class DotPlot implements IVisual {
        public static capabilities: VisualCapabilities = {
            dataRoles: [
                {
                    name: 'Category',
                    kind: powerbi.VisualDataRoleKind.Grouping,
                    displayName: data.createDisplayNameGetter("Role_DisplayName_Category")
                },
                {
                    name: 'Y',
                    kind: powerbi.VisualDataRoleKind.Measure,
                    displayName: data.createDisplayNameGetter("Role_DisplayName_Value")
                }
            ],
            dataViewMappings: [
                {
                    conditions: [
                        { 'Category': { max: 1 }, 'Y': { max: 1 } },
                    ],
                    categorical: {
                        categories: {
                            for: { in: 'Category' },
                            dataReductionAlgorithm: { top: {} }
                        },
                        values: {
                            select: [{ bind: { to: 'Y' } }]
                        }
                    }
                }
            ],
            objects: {
                general: {
                    displayName: data.createDisplayNameGetter('Visual_General'),
                    properties: {
                        formatString: {
                            type: { formatting: { formatString: true } },
                        },
                    },
                },
                label: {
                    displayName: 'Label',
                    properties: {
                        fill: {
                            displayName: 'Fill',
                            type: { fill: { solid: { color: true } } }
                        }
                    }
                }
            }
        };

        private static VisualClassName = 'dotPlot';

        private static DotPlot: ClassAndSelector = {
            class: 'dotPlot',
            selector: '.dotPlot'
        };

        private static Axis: ClassAndSelector = {
            class: 'axis',
            selector: '.axis'
        };

        private static Dot: ClassAndSelector = {
            class: 'dot',
            selector: '.dot'
        };

        private svg: D3.Selection;
        private axis: D3.Selection;
        private dotPlot: D3.Selection;

        private colors: IDataColorPalette;
        private selectionManager: SelectionManager;
        private dataView: DataView;
        private viewport: IViewport;

        private animator: IGenericAnimator;
        private margin: IMargin;
        private legend: ILegend;

        private static DefaultMargin: IMargin = {
            top: 50,
            bottom: 50,
            right: 50,
            left: 50
        };

        private radius: number;
        private strokeWidth: number;

        private static DefaultRadius: number = 5;
        private static DefaultStrokeWidth: number = 1;
        private static FrequencyText: string = "Frequency";

        private static round10(value: number, digits: number = 2) {
            const scale = Math.pow(10, digits);
            return (Math.round(scale * value) / scale);
        }

        private static getTooltipData(value: number): TooltipDataItem[] {
            return [{
                displayName: DotPlot.FrequencyText,
                value: value.toString()
            }];
        }

        public static converter(dataView: DataView, maxDots: number, colors: IDataColorPalette): DotPlotData {
            let dataPoints: DotPlotDatapoint[] = [];
            let legendData: LegendData = {
                dataPoints: [],
            };

            if (!dataView.categorical || 
                !dataView.categorical.values || 
                dataView.categorical.values.length < 1) {
                return {
                    dataPoints: dataPoints,
                    legendData: legendData,
                };
            }

            let catDv: DataViewCategorical = dataView.categorical;
            let series = catDv.values;
            for (let i = 0, iLen = series.length; i < iLen; i++) {
                let counts = [];
                let values = series[i].values;
                for (let j = 0, jLen = values.length; j < jLen; j++) {
                    let idx = values[j];
                    counts[idx] ? counts[idx]++ : counts[idx] = 1;
                }

                let legendText = series[i].source.displayName;
                let color = colors.getColorByIndex(i).value;

                let data = d3.entries(counts);
                let min = d3.min(data, d => d.value);
                let max = d3.max(data, d => d.value);

                let dotsScale: D3.Scale.LinearScale = d3.scale.linear().domain([min, max]);

                if (max > maxDots) {
                    dotsScale.rangeRound([0, maxDots]);
                    let scale = DotPlot.round10(max / maxDots);
                    legendText += ` (1 dot = x${scale})`;
                } else {
                    dotsScale.rangeRound([min, max]);
                }

                for (let k = 0, kLen = data.length; k < kLen; k++) {
                    let y = dotsScale(data[k].value);

                    for (let level = 0; level < y; level++) {
                        let id = SelectionIdBuilder
                            .builder()
                            .withSeries(dataView.categorical.values, dataView.categorical.values[i])
                            .createSelectionId();
                        dataPoints.push({
                            x: data[k].key,
                            y: level,
                            color: color,
                            identity: id,
                            tooltipInfo: DotPlot.getTooltipData(data[k].value)
                        });
                    }
                }

                legendData.dataPoints.push({
                    label: legendText,
                    color: color,
                    icon: LegendIcon.Box,
                    selected: false,
                    identity: null
                });
            }

            return {
                dataPoints: dataPoints,
                legendData: legendData
            };
        }

        public constructor(options?: DotPlotConstructorOptions) {
            if (options) {
                if (options.svg) {
                    this.svg = options.svg;
                }
                if (options.animator) {
                    this.animator = options.animator;
                }
                this.margin = options.margin || DotPlot.DefaultMargin;
                this.radius = options.radius || DotPlot.DefaultRadius;
                this.strokeWidth = options.strokeWidth || DotPlot.DefaultStrokeWidth;
            }
        }

        public init(options: VisualInitOptions): void {
            let element = options.element;
            this.selectionManager = new SelectionManager({ hostServices: options.host });

            if (!this.svg) {
                this.svg = d3.select(element.get(0)).append('svg');
            }
            if (!this.margin) {
                this.margin = DotPlot.DefaultMargin;
            }
            if (!this.radius) {
                this.radius = DotPlot.DefaultRadius;
            }
            if (!this.strokeWidth) {
                this.strokeWidth = DotPlot.DefaultStrokeWidth;
            }

            this.svg.classed(DotPlot.VisualClassName, true);
            this.colors = options.style.colorPalette.dataColors;
            this.legend = createLegend(element, false, null);

            this.dotPlot = this.svg
                .append('g')
                .classed(DotPlot.DotPlot.class, true);

            this.axis = this.svg
                .append('g')
                .classed(DotPlot.Axis.class, true);
        }

        public update(options: VisualUpdateOptions): void {
            if (!options.dataViews || !options.dataViews[0]) {
                return;
            }
            let dataView = this.dataView = options.dataViews[0];
            if (!dataView ||
                !dataView.categorical ||
                !dataView.categorical.values ||
                !dataView.categorical.values[0] ||
                !dataView.categorical.values[0].values) {
                return;
            }
            let viewport = this.viewport = options.viewport;
            this.svg
                .attr({
                    'height': viewport.height,
                    'width': viewport.width
                });

            let radius = this.radius + this.strokeWidth;
            let height = viewport.height - this.margin.bottom - this.margin.top;
            let maxDots = Math.round(height / (2 * radius));

            let data = DotPlot.converter(dataView, maxDots, this.colors);
            let dataPoints = data.dataPoints;

            let values = dataView.categorical.values[0].values;
            let xValues = d3.set(values).values();

            let xScale: D3.Scale.OrdinalScale = d3.scale.ordinal()
                .domain(xValues)
                .rangeBands([this.margin.left, viewport.width - this.margin.right]);

            let yScale: D3.Scale.LinearScale = d3.scale.linear()
                .domain([0, maxDots])
                .range([height - radius, this.margin.top]);

            this.legend.drawLegend(data.legendData, viewport);
            this.drawAxis(xValues, xScale, height);
            this.drawDotPlot(dataPoints, xScale, yScale);
        }

        private drawDotPlot(data: DotPlotDatapoint[], xScale: D3.Scale.OrdinalScale, yScale: D3.Scale.LinearScale): void {
            let selection = this.dotPlot.selectAll(DotPlot.Dot.selector).data(data);
            selection
                .enter()
                .append('circle')
                .classed(DotPlot.Dot.class, true);
            selection
                .attr("cx", function(point: DotPlotDatapoint) {
                    return xScale(point.x) + xScale.rangeBand()/2;
                })
                .attr("cy", function(point: DotPlotDatapoint) {
                    return yScale(point.y);
                })
                .attr("fill", d => d.color)
                .attr("stroke", "black")
                .attr("stroke-width", this.strokeWidth)
                .attr("r", this.radius);

            this.renderTooltip(selection);

            selection.exit().remove();
        }

        private renderTooltip(selection: D3.UpdateSelection): void {
            TooltipManager.addTooltip(selection, (tooltipEvent: TooltipEvent) => {
                return (<DotPlotDatapoint>tooltipEvent.data).tooltipInfo;
            });
        }

        private drawAxis(values: any[], xScale: D3.Scale.OrdinalScale, translateY: number) {
            let xAxis = d3.svg.axis()
                .scale(xScale)
                .orient("bottom")
                .tickValues(values);

            this.axis.attr("class", "x axis")
                .attr('transform', SVGUtil.translate(0, translateY));
            this.axis.call(xAxis);
        }
    }
}
