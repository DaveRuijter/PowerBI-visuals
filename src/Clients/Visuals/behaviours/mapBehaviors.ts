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
    export interface MapBehaviorOptions {
        dataPoints: SelectableDataPoint[];
        bubbles?: D3.Selection;
        slices?: D3.Selection;
        shapes?: D3.Selection;
        clearCatcher: D3.Selection;
    }

    export class MapBehavior implements IInteractiveBehavior {
        private bubbles: D3.Selection;
        private slices: D3.Selection;
        private shapes: D3.Selection;
        private mapPointerEventsDisabled = false;
        private mapPointerTimeoutSet = false;
        private viewChangedSinceLastClearMouseDown = false;

        public bindEvents(options: MapBehaviorOptions, selectionHandler: ISelectionHandler): void {
            let bubbles = this.bubbles = options.bubbles;
            let slices = this.slices = options.slices;
            let shapes = this.shapes = options.shapes;
            let clearCatcher = options.clearCatcher;

            let clickHandler = (d: SelectableDataPoint) => {
                if (bubbles)
                    bubbles.style("pointer-events", "all");
                if (shapes)
                    shapes.style("pointer-events", "all");
                selectionHandler.handleSelection(d, d3.event.ctrlKey);
            };

            if (!this.mapPointerEventsDisabled) {
                if (bubbles)
                    bubbles.style("pointer-events", "all");
                if (slices)
                    slices.style("pointer-events", "all");
                if (shapes)
                    shapes.style("pointer-events", "all");
            }

            if (bubbles) {
                bubbles.on('click', clickHandler);
                bubbles.on('mousewheel', () => {
                    if (!this.mapPointerEventsDisabled)
                        bubbles.style("pointer-events", "none");
                    this.mapPointerEventsDisabled = true;
                    if (!this.mapPointerTimeoutSet) {
                        this.mapPointerTimeoutSet = true;
                        setTimeout(() => {
                            if (bubbles)
                                bubbles.style("pointer-events", "all");
                            this.mapPointerEventsDisabled = false;
                            this.mapPointerTimeoutSet = false;
                        }, 200);
                    }
                });
            }

            if (slices) {
                slices.on('click', (d) => {
                    slices.style("pointer-events", "all");
                    this.mapPointerEventsDisabled = false;
                    selectionHandler.handleSelection(d.data, d3.event.ctrlKey);
                });
                slices.on('mousewheel', () => {
                    if (!this.mapPointerEventsDisabled)
                        slices.style("pointer-events", "none");
                    this.mapPointerEventsDisabled = true;
                    if (!this.mapPointerTimeoutSet) {
                        this.mapPointerTimeoutSet = true;
                        setTimeout(() => {
                            if (slices)
                                slices.style("pointer-events", "all");
                            this.mapPointerEventsDisabled = false;
                            this.mapPointerTimeoutSet = false;
                        }, 200);
                    }
                });
            }

            if (shapes) {
                shapes.on('click', clickHandler);
                shapes.on('mousewheel', () => {
                    if (!this.mapPointerEventsDisabled) {
                        shapes.style("pointer-events", "none");
                    }
                    this.mapPointerEventsDisabled = true;
                    if (!this.mapPointerTimeoutSet) {
                        this.mapPointerTimeoutSet = true;
                        setTimeout(() => {
                            if (shapes)
                                shapes.style("pointer-events", "all");
                            this.mapPointerEventsDisabled = false;
                            this.mapPointerTimeoutSet = false;
                        }, 200);
                    }
                });
            }

            clearCatcher.on('mouseup', () => {
                if (!this.viewChangedSinceLastClearMouseDown)
                    selectionHandler.handleClearSelection();
            });

            clearCatcher.on('mousedown', () => {
                this.viewChangedSinceLastClearMouseDown = false;
            });
        }

        public renderSelection(hasSelection: boolean) {
            if (this.bubbles) {
                this.bubbles
                    .style({
                        'fill-opacity': (d: MapBubble) => ColumnUtil.getFillOpacity(d.selected, false, hasSelection, false),
                        'stroke-opacity': (d: MapBubble) => ColumnUtil.getFillOpacity(d.selected, false, hasSelection, false),
                    });
            }
            if (this.slices) {
                this.slices
                    .style({
                        "fill-opacity": (d) => ColumnUtil.getFillOpacity(d.data.selected, false, hasSelection, false),
                        "stroke-opacity": (d) => ColumnUtil.getFillOpacity(d.data.selected, false, hasSelection, false),
                    });
            }
            if (this.shapes) {
                this.shapes
                    .style({
                        "fill-opacity": (d) => ColumnUtil.getFillOpacity(d.selected, false, hasSelection, false),
                        "stroke-opacity": (d) => ColumnUtil.getFillOpacity(d.selected, false, hasSelection, false),
                    });
            }
        }

        public viewChanged() {
            this.viewChangedSinceLastClearMouseDown = true;
        }
    }
}