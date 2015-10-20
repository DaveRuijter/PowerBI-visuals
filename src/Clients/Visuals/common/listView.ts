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

    export interface IListView {
        data(data: any[], dataIdFunction: (d) => {}, dataAppended: boolean): IListView;
        rowHeight(rowHeight: number): IListView;
        viewport(viewport: IViewport): IListView;
        render(): void;
        empty(): void;
    }

    export module ListViewFactory {
        export function createListView(options): IListView {
            return new ListView(options);
        }
    }

    export interface ListViewOptions {
        enter: (selection: D3.Selection) => void;
        exit: (selection: D3.Selection) => void;
        update: (selection: D3.Selection) => void;
        loadMoreData: () => void;
        baseContainer: D3.Selection;
        rowHeight: number;
        viewport: IViewport;
        scrollEnabled: boolean;
    }

    /**
     * A UI Virtualized List, that uses the D3 Enter, Update & Exit pattern to update rows.
     * It can create lists containing either HTML or SVG elements.
     */
    class ListView implements IListView {
        private getDatumIndex: (d: any) => {};
        private _data: any[];
        private _totalRows: number;

        private options: ListViewOptions;
        private visibleGroupContainer: D3.Selection;
        private scrollContainer: D3.Selection;
        private cancelMeasurePass: () => void;
        private renderTimeoutId: number;
        
        /**
         * The value indicates the percentage of data already shown
         * in the list view that triggers a loadMoreData call.
         */
        private static loadMoreDataThreshold = 0.8;
        private static defaultRowHeight = 1;

        public constructor(options: ListViewOptions) {
            // make a copy of options so that it is not modified later by caller
            this.options = $.extend(true, {}, options);

            this.options.baseContainer
                .style('overflow-y', 'auto')
                .on('scroll', () => this.renderImpl(this.options.rowHeight));
            this.scrollContainer = options.baseContainer
                .append('div')
                .attr('class', 'scrollRegion');
            this.visibleGroupContainer = this.scrollContainer
                .append('div')
                .attr('class', 'visibleGroup');

            ListView.SetDefaultOptions(options);
        }

        private static SetDefaultOptions(options: ListViewOptions) {
            options.rowHeight = options.rowHeight || ListView.defaultRowHeight;
        }

        public rowHeight(rowHeight: number): ListView {
            this.options.rowHeight = Math.ceil(rowHeight);
            return this;
        }

        public data(data: any[], getDatumIndex: (d) => {}, dataReset: boolean = false): IListView {
            this._data = data;
            this.getDatumIndex = getDatumIndex;
            this.setTotalRows();
            if (dataReset) {
                $(this.options.baseContainer.node()).scrollTop(0);
            }
            this.render();
            return this;
        }

        public viewport(viewport: IViewport): IListView {
            this.options.viewport = viewport;
            this.render();
            return this;
        }

        public empty(): void {
            this._data = [];
            this.render();
        }

        public render(): void {
            if (this.renderTimeoutId)
                window.clearTimeout(this.renderTimeoutId);

            this.renderTimeoutId = window.setTimeout(() => {
                this.getRowHeight().then((rowHeight: number) => {
                    this.renderImpl(rowHeight);
                });
                this.renderTimeoutId = undefined;
            },0);
            }

        private renderImpl(rowHeight: number): void {
            var totalHeight = this.options.scrollEnabled ? Math.max(0, (this._totalRows * rowHeight)) : this.options.viewport.height;
            this.scrollContainer
                .style('height', totalHeight + "px")
                .attr('height', totalHeight);

            this.scrollToFrame(true /*loadMoreData*/);
        }

        private scrollToFrame(loadMoreData: boolean): void {
            var options = this.options;
            var visibleGroupContainer = this.visibleGroupContainer;
            var totalRows = this._totalRows;
            var rowHeight = options.rowHeight || ListView.defaultRowHeight;
            var visibleRows = this.getVisibleRows() || 1;
            var scrollTop: number = options.baseContainer.node().scrollTop;
            var scrollPosition = (scrollTop === 0) ? 0 : Math.floor(scrollTop / rowHeight);
            var translateY = scrollPosition * rowHeight;

            visibleGroupContainer
                .attr('transform', d => SVGUtil.translate(0, translateY))
                .style({
                    //order matters for proper overriding
                    'transform': d => SVGUtil.translateWithPixels(0, translateY),
                    '-webkit-transform': d => SVGUtil.translateWithPixels(0, translateY)
                });

            var position0 = Math.max(0, Math.min(scrollPosition, totalRows - visibleRows + 1)),
                position1 = position0 + visibleRows;
            var rowSelection = visibleGroupContainer.selectAll(".row")
                .data(this._data.slice(position0, Math.min(position1, totalRows)), this.getDatumIndex);

            rowSelection
                .enter()
                .append('div')
                .classed('row', true)
                .call(d => options.enter(d));
            rowSelection.order();

            var rowUpdateSelection = visibleGroupContainer.selectAll('.row:not(.transitioning)');

            rowUpdateSelection.call(d => options.update(d));

            rowSelection
                .exit()
                .call(d => options.exit(d))
                .remove();

            if (loadMoreData && visibleRows !== totalRows && position1 >= totalRows * ListView.loadMoreDataThreshold)
                options.loadMoreData();
        }

        private setTotalRows(): void {
            var data = this._data;
            this._totalRows = data ? data.length : 0;
        }

        private getVisibleRows(): number {
            var minimumVisibleRows = 1;
            var rowHeight = this.options.rowHeight;
            var viewportHeight = this.options.viewport.height;

            if (!rowHeight || rowHeight < 1)
                return minimumVisibleRows;
            
            if (this.options.scrollEnabled)
                return Math.min(Math.ceil(viewportHeight / rowHeight) + 1, this._totalRows) || minimumVisibleRows;

            return Math.min(Math.floor(viewportHeight / rowHeight), this._totalRows) || minimumVisibleRows;
        }

        private getRowHeight(): JQueryPromise<number> {
            var deferred = $.Deferred<number>();
            var listView = this;
            var options = listView.options;
            if (this.cancelMeasurePass)
                this.cancelMeasurePass();

            // if there is no data, resolve and return
            if (!(this._data && this._data.length && options)) {
                listView.rowHeight(ListView.defaultRowHeight);
                return deferred.resolve(options.rowHeight).promise();
            }

            //render the first item to calculate the row height
            this.scrollToFrame(false /*loadMoreData*/);
            var requestAnimationFrameId = window.requestAnimationFrame(() => {
                //measure row height
                var firstRow = listView.visibleGroupContainer.select(".row").node().firstChild;
                var rowHeight: number = $(firstRow).outerHeight(true);
                listView.rowHeight(rowHeight);
                deferred.resolve(rowHeight);
                listView.cancelMeasurePass = undefined;
                window.cancelAnimationFrame(requestAnimationFrameId);
            });

            this.cancelMeasurePass = () => {
                window.cancelAnimationFrame(requestAnimationFrameId);
                deferred.reject();
            };

            return deferred.promise();
        }
    }
}