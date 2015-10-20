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
    export interface DonutBehaviorOptions {
        slices: D3.Selection;
        highlightSlices: D3.Selection;
        clearCatcher: D3.Selection;
        hasHighlights: boolean;
    }

    export class DonutChartWebBehavior implements IInteractiveBehavior {
        private slices: D3.Selection;
        private highlightSlices: D3.Selection;
        private hasHighlights: boolean;

        public bindEvents(options: DonutBehaviorOptions, selectionHandler: ISelectionHandler): void {
            let slices = this.slices = options.slices;
            let highlightSlices = this.highlightSlices = options.highlightSlices;
            let clearCatcher = options.clearCatcher;
            this.hasHighlights = options.hasHighlights;

            let clickHandler = (d: DonutArcDescriptor) => {
                selectionHandler.handleSelection(d.data, d3.event.ctrlKey);
            };

            slices.on('click', clickHandler);
            highlightSlices.on('click', clickHandler);

            clearCatcher.on('click', () => {
                selectionHandler.handleClearSelection();
            });
        }

        public renderSelection(hasSelection: boolean): void {
            let hasHighlights = this.hasHighlights;
            this.slices.style("fill-opacity", (d: DonutArcDescriptor) => ColumnUtil.getFillOpacity(d.data.selected, false, hasSelection, hasHighlights && !d.data.selected));
            this.highlightSlices.style("fill-opacity", (d: DonutArcDescriptor) => ColumnUtil.getFillOpacity(d.data.selected, true, false, hasHighlights));
        }
    }
}  