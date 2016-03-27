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
    
    export interface PlayBehaviorOptions extends ScatterBehaviorOptions {
        data: PlayChartData;
        dataViewCat?: powerbi.DataViewCategorical;
        svg?: D3.Selection;
        dataView?: powerbi.DataView;
        renderTraceLine?: (options: PlayBehaviorOptions, selectedPoints: SelectableDataPoint[], shouldAnimate: boolean) => void;
        labelsSelection: D3.Selection;
    }

    export class PlayChartWebBehavior implements IInteractiveBehavior {
        private bubbles: D3.Selection;
        private shouldEnableFill: boolean;
        private options: PlayBehaviorOptions;

        public bindEvents(options: PlayBehaviorOptions, selectionHandler: ISelectionHandler): void {
            this.options = options;
            if (options && options.dataPointsSelection) {
                var bubbles = this.bubbles = options.dataPointsSelection;
                var data = options.data;
                this.shouldEnableFill = (!data.sizeRange || !data.sizeRange.min) && data.fillPoint;

                bubbles.on('click', (d: SelectableDataPoint) => {
                    selectionHandler.handleSelection(d, d3.event.ctrlKey);
                });
            }
        }

        public renderSelection(hasSelection: boolean): void {
            var shouldEnableFill = this.shouldEnableFill;
            if (this.bubbles) {
                this.bubbles.style("fill-opacity", (d: PlayChartDataPoint) => (d.size != null || shouldEnableFill) ? PlayChart.getBubbleOpacity(d, hasSelection) : 0);
                this.bubbles.style("stroke-opacity", (d: PlayChartDataPoint) => PlayChart.getBubbleOpacity(d, hasSelection));
                if (this.options.labelsSelection)
                    this.options.labelsSelection.style("opacity", (d: PlayChartDataPoint) => PlayChart.getBubbleOpacity(d, hasSelection));

                let selectedPoints = this.bubbles.filter((d: PlayChartDataPoint) => d.selected);
                if (selectedPoints && selectedPoints.data().length > 0 && this.options.renderTraceLine != null) {
                    this.options.renderTraceLine(this.options, selectedPoints.data(), true);
                }
                else {
                    this.options.svg.selectAll('.traceLine').remove();
                    this.options.svg.selectAll('.traceBubble').remove();
                }
            }
        }
    }
} 