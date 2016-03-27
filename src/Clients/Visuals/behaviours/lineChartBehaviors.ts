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
    export interface LineChartBehaviorOptions {
        lines: D3.Selection;
        interactivityLines: D3.Selection;
        dots: D3.Selection;
        areas: D3.Selection;
        isPartOfCombo?: boolean;
    }

    export class LineChartWebBehavior implements IInteractiveBehavior {
        private lines: D3.Selection;
        private dots: D3.Selection;
        private areas: D3.Selection;

        public bindEvents(options: LineChartBehaviorOptions, selectionHandler: ISelectionHandler): void {
            this.lines = options.lines;
            let interactivityLines = options.interactivityLines;
            let dots = this.dots = options.dots;
            let areas = this.areas = options.areas;

            let clickHandler = (d: SelectableDataPoint) => {
                selectionHandler.handleSelection(d, d3.event.ctrlKey);
            };

            interactivityLines.on('click', clickHandler);
            dots.on('click', clickHandler);
            if (areas)
                areas.on('click', clickHandler);
        }

        public renderSelection(hasSelection: boolean) {
            this.lines.style("stroke-opacity", (d: SelectableDataPoint) => ColumnUtil.getFillOpacity(d.selected, false, hasSelection, false));
            this.dots.style("fill-opacity", (d: SelectableDataPoint) => ColumnUtil.getFillOpacity(d.selected, false, hasSelection, false));
            if (this.areas)
                this.areas.style("fill-opacity", (d: SelectableDataPoint) => (hasSelection && !d.selected) ? LineChart.DimmedAreaFillOpacity : LineChart.AreaFillOpacity);
        }
    }
} 