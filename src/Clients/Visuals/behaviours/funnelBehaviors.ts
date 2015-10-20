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
    export interface FunnelBehaviorOptions {
        bars: D3.Selection;
        interactors: D3.Selection;
        clearCatcher: D3.Selection;
        hasHighlights: boolean;
    }

    export class FunnelWebBehavior implements IInteractiveBehavior {
        private bars: D3.Selection;
        private interactors: D3.Selection;
        private hasHighlights: boolean;

        public bindEvents(options: FunnelBehaviorOptions, selectionHandler: ISelectionHandler): void {

            let bars = this.bars = options.bars;
            let interactors = this.interactors = options.interactors;
            let clearCatcher = options.clearCatcher;

            this.hasHighlights = options.hasHighlights;

            bars.on('click', (d: SelectableDataPoint, i: number) => {
                selectionHandler.handleSelection(d, d3.event.ctrlKey);
            });

            interactors.on('click', (d: SelectableDataPoint, i: number) => {
                selectionHandler.handleSelection(d, d3.event.ctrlKey);
            });

            clearCatcher.on('click', () => {
                selectionHandler.handleClearSelection();
            });
        }

        public renderSelection(hasSelection: boolean): void {
            let hasHighlights = this.hasHighlights;
            this.bars.style("fill-opacity", (d: FunnelSlice) => ColumnUtil.getFillOpacity(d.selected, d.highlight, !d.highlight && hasSelection, !d.selected && hasHighlights));
        }
    }
}
