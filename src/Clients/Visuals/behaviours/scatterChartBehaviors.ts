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

    export interface ScatterBehaviorOptions {
        host: ICartesianVisualHost;
        root: D3.Selection;
        mainContext: D3.Selection;
        background: D3.Selection;
        dataPointsSelection: D3.Selection;
        data: ScatterChartData;
        visualInitOptions: VisualInitOptions;
        xAxisProperties: IAxisProperties;
        yAxisProperties: IAxisProperties;
    }

    export class ScatterChartWebBehavior implements IInteractiveBehavior {
        private bubbles: D3.Selection;
        private shouldEnableFill: boolean;
        private colorBorder: boolean;

        public bindEvents(options: ScatterBehaviorOptions, selectionHandler: ISelectionHandler): void {
            let bubbles = this.bubbles = options.dataPointsSelection;
            let data = options.data;
            this.shouldEnableFill = (!data.sizeRange || !data.sizeRange.min) && data.fillPoint;
            this.colorBorder = data.colorBorder;

            bubbles.on('click', (d: SelectableDataPoint) => {
                selectionHandler.handleSelection(d, d3.event.ctrlKey);
            });
        }

        public renderSelection(hasSelection: boolean) {
            let shouldEnableFill = this.shouldEnableFill;
            let colorBorder = this.colorBorder;
            this.bubbles.style("fill-opacity", (d: ScatterChartDataPoint) => (d.size != null || shouldEnableFill) ? ScatterChart.getBubbleOpacity(d, hasSelection) : 0);
            this.bubbles.style("stroke-opacity", (d: ScatterChartDataPoint) => (d.size != null && colorBorder) ? 1 : ScatterChart.getBubbleOpacity(d, hasSelection));
        }
    }

    export const enum DragType {
        Drag,
        DragEnd
    }

    interface MouseCoordinates {
        x: number;
        y: number;
    }

    export class ScatterChartMobileBehavior implements IInteractiveBehavior {
        private static CrosshairClassName = 'crosshair';
        private static ScatterChartCircleTagName = 'circle';
        private static DotClassName = 'dot';
        private static DotClassSelector = '.' + ScatterChartMobileBehavior.DotClassName;

        private static Horizontal: ClassAndSelector = {
            class: 'horizontal',
            selector: '.horizontal'
        };
        private static Vertical: ClassAndSelector = {
            class: 'vertical',
            selector: '.vertical'
        };

        private host: ICartesianVisualHost;
        private mainGraphicsContext: D3.Selection;
        private data: ScatterChartData;
        private crosshair: D3.Selection;
        private crosshairHorizontal: D3.Selection;
        private crosshairVertical: D3.Selection;
        private lastDotIndex: number;
        private xAxisProperties: IAxisProperties;
        private yAxisProperties: IAxisProperties;

        public bindEvents(options: ScatterBehaviorOptions, selectionHandler: ISelectionHandler): void {
            this.setOptions(options);

            this.makeDataPointsSelectable(options.dataPointsSelection);
            this.makeRootSelectable(options.root);
            this.makeDragable(options.root);

            this.selectRoot();
        }

        public renderSelection(HasSelection: boolean): void { }

        public setSelectionHandler(selectionHandler: ISelectionHandler): void { }

        private makeDataPointsSelectable(...selection: D3.Selection[]): void {
            for (let i = 0, len = selection.length; i < len; i++) {
                var sel = selection[i];

                sel.on('click', (d: SelectableDataPoint, i: number) => {
                    this.select(true, sel, d, i);
                });
            }
        }

        private makeRootSelectable(selection: D3.Selection): void {
            selection.on('click', (d: SelectableDataPoint, i: number) => {
                this.selectRoot();
            });
        }

        private makeDragable(...selection: D3.Selection[]): void {
            for (let i = 0, len = selection.length; i < len; i++) {
                let sel = selection[i];

                let drag = d3.behavior.drag()
                    .on('drag', (d) => { this.drag(DragType.Drag); })
                    .on('dragend', (d) => { this.drag(DragType.DragEnd); });

                sel.call(drag);
            }
        }

        public setOptions(options: ScatterBehaviorOptions) {
            this.data = options.data;
            this.mainGraphicsContext = options.mainContext;
            this.xAxisProperties = options.xAxisProperties;
            this.yAxisProperties = options.yAxisProperties;
            this.host = options.host;
        }

        public select(hasSelection: boolean, datapoints: D3.Selection, dataPoint: SelectableDataPoint, index: number) {
            this.selectDotByIndex(index);
        }

        public selectRoot() {
            let marker = jsCommon.PerformanceUtil.create('selectRoot');
            this.onClick();
            marker.end();
        }

        public drag(t: DragType) {
            switch (t) {
                case DragType.Drag:
                    this.onDrag();
                    break;
                case DragType.DragEnd:
                    this.onClick();
                    break;
                default:
                    debug.assertFail('Unknown Drag Type');
            }
        }

        private onDrag(): void {
            //find the current x and y position
            let xy = this.getMouseCoordinates();
            //move the crosshair to the current position
            this.moveCrosshairToXY(xy.x, xy.y);
            //update the style and the legend of the dots
            let selectedIndex = this.findClosestDotIndex(xy.x, xy.y);
            this.selectDot(selectedIndex);
            this.updateLegend(selectedIndex);
        }

        private onClick(): void {
            //find the current x and y position
            let xy = this.getMouseCoordinates();
            let selectedIndex = this.findClosestDotIndex(xy.x, xy.y);
            if (selectedIndex !== -1)
            this.selectDotByIndex(selectedIndex);
        }

        private getMouseCoordinates(): MouseCoordinates {
            let mainGfxContext = this.mainGraphicsContext;
            // select (0,0) in cartesian coordinates
            let x = 0;
            let y = parseInt(mainGfxContext.attr('height'), 10);
            y = y || 0;

            try {
                let mouse = d3.mouse(mainGfxContext.node());
                x = mouse[0];
                y = mouse[1];
            } catch(e){ 
            }

            return { x: x, y: y, };
        }

        private selectDotByIndex(index: number): void {
            this.selectDot(index);
            this.moveCrosshairToIndexDot(index);
            this.updateLegend(index);
        }

        private selectDot(dotIndex: number): void {
            let root = this.mainGraphicsContext;

            root.selectAll(ScatterChartMobileBehavior.ScatterChartCircleTagName + ScatterChartMobileBehavior.DotClassSelector).classed({ selected: false, notSelected: true });
            root.selectAll(ScatterChartMobileBehavior.ScatterChartCircleTagName + ScatterChartMobileBehavior.DotClassSelector).filter((d, i) => {
                let dataPoints = this.data.dataPoints;
                debug.assert(dataPoints.length > dotIndex, "dataPoints length:" + dataPoints.length + "is smaller than index:" + dotIndex);
                let currentPoint: ScatterChartDataPoint = dataPoints[dotIndex];
                return (d.x === currentPoint.x) && (d.y === currentPoint.y);
            }).classed({ selected: true, notSelected: false });
        }

        private moveCrosshairToIndexDot(index: number): void {
            let dataPoints = this.data.dataPoints;
            let root = this.mainGraphicsContext;

            debug.assert(dataPoints.length > index, "dataPoints length:" + dataPoints.length + "is smaller than index:" + index);
            let x = this.xAxisProperties.scale(dataPoints[index].x);
            let y = this.yAxisProperties.scale(dataPoints[index].y);
            if (this.crosshair == null) {
                let width = +root.attr('width');
                let height = +root.attr('height');
                this.crosshair = this.drawCrosshair(root, x, y, width, height);
                this.crosshairHorizontal = this.crosshair.select(ScatterChartMobileBehavior.Horizontal.selector);
                this.crosshairVertical = this.crosshair.select(ScatterChartMobileBehavior.Vertical.selector);
            } else {
                this.moveCrosshairToXY(x, y);
            }
        }

        private moveCrosshairToXY(x: number, y: number): void {
            this.crosshairHorizontal.attr({ y1: y, y2: y });
            this.crosshairVertical.attr({ x1: x, x2: x });
        }

        private drawCrosshair(addTo: D3.Selection, x: number, y: number, width: number, height: number): D3.Selection {
            let crosshair = addTo.append("g");
            crosshair.classed(ScatterChartMobileBehavior.CrosshairClassName, true);
            crosshair.append('line').classed(ScatterChartMobileBehavior.Horizontal.class, true).attr({ x1: 0, x2: width, y1: y, y2: y });
            crosshair.append('line').classed(ScatterChartMobileBehavior.Vertical.class, true).attr({ x1: x, x2: x, y1: height, y2: 0 });
            return crosshair;
        }

        private findClosestDotIndex(x: number, y: number): number {
            let selectedIndex = -1;
            let minDistance = Number.MAX_VALUE;
            let dataPoints = this.data.dataPoints;
            let xAxisPropertiesScale = this.xAxisProperties.scale;
            let yAxisPropertiesScale = this.yAxisProperties.scale;
            for (let i in dataPoints) {
                let currentPoint: ScatterChartDataPoint = dataPoints[i];
                let circleX = xAxisPropertiesScale(currentPoint.x);
                let circleY = yAxisPropertiesScale(currentPoint.y);
                let horizontalDistance = circleX - x;
                let verticalDistance = circleY - y;
                let distanceSqrd = (horizontalDistance * horizontalDistance) + (verticalDistance * verticalDistance);
                if (minDistance === Number.MAX_VALUE) {
                    selectedIndex = i;
                    minDistance = distanceSqrd;
                }
                else if (minDistance && minDistance > distanceSqrd) {
                    selectedIndex = i;
                    minDistance = distanceSqrd;
                }
            }
            return selectedIndex;
        }

        private updateLegend(dotIndex: number): void {
            if (this.lastDotIndex == null || this.lastDotIndex !== dotIndex) {//update the legend only if the data change.
                let legendItems = this.createLegendDataPoints(dotIndex);
                this.host.updateLegend(legendItems);
                this.lastDotIndex = dotIndex;
            }
        }

        private createLegendDataPoints(dotIndex: number): LegendData {
            let formatStringProp = scatterChartProps.general.formatString;
            let legendItems = [];
            let data = this.data;
            debug.assert(data.dataPoints.length > dotIndex, "dataPoints length:" + data.dataPoints.length + "is smaller than index:" + dotIndex);
            let point = data.dataPoints[dotIndex];
            //set the title of the legend to be the category or radius or group or blank
            let blank = valueFormatter.format(null);
            let title = blank;
            let legendData = data.legendData;
            debug.assertValue(legendData, "legendData");
            debug.assertValue(legendData.dataPoints, "legendData");
            let legendDataPoints = legendData.dataPoints;
            if (point.category !== blank) {
                title = point.category;
            } else if (point.radius.sizeMeasure != null) {
                title = valueFormatter.format(point.radius.sizeMeasure.source.groupName);
            } else if (legendDataPoints.length >= dotIndex && legendDataPoints[dotIndex].label !== blank) {
                title = legendDataPoints[dotIndex].label;
            }
            if (data.xCol != null) {
                legendItems.push({
                    category: title,
                    color: point.fill,
                    icon: LegendIcon.Box,
                    label: valueFormatter.format(this.data.axesLabels.x),
                    measure: valueFormatter.format(point.x, valueFormatter.getFormatString(data.xCol, formatStringProp)),
                    iconOnlyOnLabel: true
                });
            }
            if (data.yCol !== undefined && data.yCol !== null) {
                legendItems.push({
                    category: title,
                    color: point.fill,
                    icon: LegendIcon.Box,
                    label: valueFormatter.format(data.axesLabels.y),
                    measure: valueFormatter.format(point.y, valueFormatter.getFormatString(data.yCol, formatStringProp)),
                    iconOnlyOnLabel: true
                });
            }
            if (data.size !== undefined && data.size !== null) {
                legendItems.push({
                    category: title,
                    color: point.fill,
                    icon: LegendIcon.Box,
                    label: valueFormatter.format(data.size.displayName),
                    measure: valueFormatter.format(point.radius.sizeMeasure.values[point.radius.index], valueFormatter.getFormatString(data.size, formatStringProp)),
                    iconOnlyOnLabel: true
                });
            }

            return {dataPoints: legendItems };
        }
    }
} 