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

/// <reference path="../../../_references.ts"/>

module powerbi.visuals.controls.internal {
    /**
     * This class is used for layouts that don't or cannot
     * rely on DOM measurements.  Instead they compute all required
     * widths and heights and store it in this structure.
     */
    export class SizeComputationManager {

        // Unfortunately since we are doing manual layout, we need to hardcode some layout properties here.
        // These must be kept in sync with what is specified in the .bi-dashboard-tablix class.
        private static DashboardCellPaddingLeft = 10;
        private static DashboardCellPaddingRight = 5;
        private static DashboardRowHeight = 19;

        // Minimum size for a column, used to calculate layout
        private static TablixMinimumColumnWidth = 75;

        private _viewport: IViewport;
        private _columnCount: number;
        private _cellWidth: number;
        private _cellHeight: number;

        public hasImageContent: boolean;

        public get visibleWidth(): number {
            return this._viewport ? this._viewport.width : 0;
        }

        public get visibleHeight(): number {
            return this._viewport ? this._viewport.height : 0;
        }

        public get gridWidth(): number {
            return this.visibleWidth;
        }

        public get gridHeight(): number {
            return this.visibleHeight;
        }

        public get rowHeight(): number {
            return this._cellHeight;
        }

        public get cellWidth(): number {
            return this._cellWidth;
        }

        public get cellHeight(): number {
            return this._cellHeight;
        }

        public get contentWidth(): number {
            return this._cellWidth - SizeComputationManager.DashboardCellPaddingLeft - SizeComputationManager.DashboardCellPaddingRight;
        }

        public get contentHeight(): number {
            return this._cellHeight;
        }

        public updateColumnCount(columnCount: number): void {
            this._columnCount = columnCount;
        }

        public updateViewport(viewport: IViewport): void {

            this._viewport = viewport;

            this._cellWidth = this.computeColumnWidth(this._columnCount);
            this._cellHeight = this.computeColumnHeight();
        }

        private computeColumnWidth(totalColumnCount: number): number {
            let maxAllowedColumns = Math.floor(this._viewport.width / SizeComputationManager.TablixMinimumColumnWidth);
            return this.fitToColumnCount(maxAllowedColumns, totalColumnCount);
            }

        private computeColumnHeight(): number {
            if (!this.hasImageContent)
                return SizeComputationManager.DashboardRowHeight;

            var width = this._viewport.width;
            if (width <= 250) {
                // Small
                return 20;
            }
            else if (width <= 510) {
                // Medium
                return 51;
            }
            else if (width <= 770) {
                // Large
                return 52;
            }

            debug.assertFail("Fixed size is only for viewport up to 770px width.");
        }

        private fitToColumnCount(maxAllowedColumnCount: number, totalColumnCount: number): number {
            let columnsToFit = Math.min(maxAllowedColumnCount, totalColumnCount);
            return Math.floor(this._viewport.width / columnsToFit);
        }
    }

    export class DimensionLayoutManager implements IDimensionLayoutManager {
        public static _pixelPrecision = 1.0001;
        public static _scrollOffsetPrecision = 0.01;

        public _grid: TablixGrid; // internal
        public _gridOffset: number; // internal

        protected _contextualWidthToFill: number;

        private _owner: TablixLayoutManager;
        private _realizationManager: TablixDimensionRealizationManager;
        private _alignToEnd: boolean;
        private _lastScrollOffset: number;
        private _isScrolling: boolean;
        private _fixedSizeEnabled: boolean;
        private _done: boolean;
        private _measureEnabled: boolean;

        constructor(owner: TablixLayoutManager, grid: TablixGrid, realizationManager: TablixDimensionRealizationManager) {
            //debug.assertValue(realizationManager, "Realization Manager must be defined");

            this._owner = owner;
            this._grid = grid;
            this._lastScrollOffset = null;
            this._isScrolling = false;
            this._fixedSizeEnabled = true;
            this._done = false;

            this._realizationManager = realizationManager;
        }

        public get owner(): TablixLayoutManager {
            return this._owner;
        }

        public get realizationManager(): TablixDimensionRealizationManager {
            return this._realizationManager;
        }

        public get fixedSizeEnabled(): boolean {
            return this._fixedSizeEnabled;
        }

        public set fixedSizeEnabled(enable: boolean) {
            this._fixedSizeEnabled = enable;
        }

        public onCornerCellRealized(item: any, cell: ITablixCell, leaf: boolean): void {
            this._realizationManager.onCornerCellRealized(item, cell);
        }

        public onHeaderRealized(item: any, cell: ITablixCell, leaf): void {
            this._realizationManager.onHeaderRealized(item, cell, leaf);
        }

        public get needsToRealize(): boolean {
            return this._realizationManager.needsToRealize;
        }

        public getVisibleSizeRatio(): number {
            return 1 - this.dimension.getFractionScrollOffset();
        }

        public get alignToEnd(): boolean {
            return this._alignToEnd;
        }

        public get done(): boolean {
            return this._done;
        }

        public _requiresMeasure(): boolean {
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager._requiresMeasure");
            return true;
        }

        public startScrollingSession(): void {
            this._isScrolling = true;
        }

        public endScrollingSession(): void {
            this._isScrolling = false;
        }

        public isScrolling(): boolean {
            return this._isScrolling;
        }

        public isResizing(): boolean {
            return false;
        }

        public getOtherHierarchyContextualHeight(): number {
            let otherDimension = this.dimension.otherDimension;
            let count = otherDimension.getDepth();

            let contextualHeight = 0;
            let items = this._getRealizedItems();

            if (items.length > 0) {
                for (let i = 0; i < count; i++) {
                    contextualHeight += items[i].getContextualWidth();
                }
            }

            return contextualHeight;
        }

        public _isAutoSized(): boolean {
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager._isAutoSized");
            return false;
        }

        public onStartRenderingSession() {
            this._measureEnabled = this._requiresMeasure();
            this._gridOffset = this.dimension.otherDimension.getDepth();
        }

        public onEndRenderingSession() {
            this._realizationManager.onEndRenderingSession();
            this._alignToEnd = false;
            this._done = false;
            this._measureEnabled = true;
            this._sendDimensionsToControl();
        }

        /**
         * Implementing classes must override this to send dimentions to TablixControl.
         */
        public _sendDimensionsToControl(): void { //extending class overrides this
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager._sendDimensionsToControl");
        }

        public get measureEnabled() {
            return this._measureEnabled;
        }

        public getFooterContextualWidth(): number {
            return 0;
        }

        public onStartRenderingIteration(clear: boolean, contextualWidth: number) {
            if (this._measureEnabled && !this._done) {
                this._contextualWidthToFill = (contextualWidth - this.otherScrollbarContextualWidth) * this.getGridScale() - this.getFooterContextualWidth();
            }

            this._realizationManager.onStartRenderingIteration();

            if (clear) {
                this._lastScrollOffset = null;
            }
            else if (this._lastScrollOffset !== null) {
                this.swapElements();
            }
        }

        public get allItemsRealized(): boolean {
            return this.getRealizedItemsCount() - this._gridOffset === this.dimension.getItemsCount() || this.dimension.getItemsCount() === 0;
        }

        public onEndRenderingIteration(): void {
            if (this._done) {
                return;
            }

            if (!this._measureEnabled) {
                this._lastScrollOffset = this.dimension.scrollOffset;
                this._done = true;
                return;
            }

            let gridContextualWidth: number = this.getGridContextualWidth();

            let filled: boolean = Double.greaterOrEqualWithPrecision(gridContextualWidth, this._contextualWidthToFill, DimensionLayoutManager._pixelPrecision);
            let allRealized = this.allItemsRealized;

            let newScrollOffset;

            if (filled) {
                newScrollOffset = this.scrollForwardToAlignEnd(gridContextualWidth);
            }
            else {
                newScrollOffset = this.scrollBackwardToFill(gridContextualWidth);
            }

            this._realizationManager.onEndRenderingIteration(gridContextualWidth, filled);

            let originalScrollbarVisible: boolean = this.dimension.scrollbar.visible;

            this.updateScrollbar(gridContextualWidth);

            this._done = (filled || allRealized) &&
            this.dimension.scrollbar.visible === originalScrollbarVisible &&
            Double.equalWithPrecision(newScrollOffset, this.dimension.scrollOffset, DimensionLayoutManager._scrollOffsetPrecision);

            this.dimension.scrollOffset = newScrollOffset;
            this._lastScrollOffset = this.dimension.scrollOffset;
        }

        private getScrollDeltaWithinPage(): number {
            if (this._lastScrollOffset !== null) {
                let delta = this.dimension.getIntegerScrollOffset() - Math.floor(this._lastScrollOffset);
                if (Math.abs(delta) < this.getRealizedItemsCount() - this.dimension.otherDimension.getDepth()) {
                    return delta;
                }
            }
            return null;
        }

        private swapElements() {
            let delta = this.getScrollDeltaWithinPage();
            if (delta !== null) {
                let otherHierarchyDepth = this.dimension.otherDimension.getDepth();

                if (Math.abs(delta) < this.getRealizedItemsCount() - otherHierarchyDepth) {
                    if (delta > 0) {
                        this._moveElementsToBottom(otherHierarchyDepth, delta);
                    }
                    else if (delta < 0) {
                        this._moveElementsToTop(otherHierarchyDepth, -delta);
                    }
                }
            }
        }

        public _getRealizedItems(): ITablixGridItem[] {
            // abstract
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager._getRealizedItems");
            return null;
        }

        public getRealizedItemsCount(): number {
            let realizedItems = this._getRealizedItems();
            return realizedItems.length;
        }

        public _moveElementsToBottom(moveFromIndex: number, count): void {
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager._moveElementsToBottom");
        }

        public _moveElementsToTop(moveToIndex: number, count): void {
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager._moveElementsToTop");
        }

        public isScrollingWithinPage(): boolean {
            return this.getScrollDeltaWithinPage() !== null;
        }

        public getGridContextualWidth(): number {
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager.getGridContextualWidth");
            return 0;
        }

        private updateScrollbar(gridContextualWidth: number): void {
            let scrollbar = this.dimension.scrollbar;
            scrollbar.viewMin = this.dimension.scrollOffset;
            scrollbar.min = 0;
            scrollbar.max = this.dimension.getItemsCount();
            scrollbar.viewSize = this.getViewSize(gridContextualWidth);
            this.dimension.scrollbar.show(this.canScroll(gridContextualWidth));
        }

        public getViewSize(gridContextualWidth: number): number {
            let count: number = this.getRealizedItemsCount();
            if (count === 0)
                return 0;

            let startIndex = this._gridOffset;
            let sizeInItems = 0;
            let sizeInPixels = 0;

            let widthToFill: number = this._contextualWidthToFill;
            let scrollableArea = widthToFill - this.getOtherHierarchyContextualHeight();

            let error = this.getMeaurementError(gridContextualWidth);

            for (let i = startIndex; i < count; i++) {
                let visibleRatio;
                if (i === startIndex) {
                    visibleRatio = this.getVisibleSizeRatio();
                }
                else
                    visibleRatio = 1;

                let itemContextualWidth = this.getItemContextualWidthWithScrolling(i) * error;

                sizeInPixels += itemContextualWidth;
                sizeInItems += visibleRatio;

                if (Double.greaterWithPrecision(sizeInPixels, scrollableArea, DimensionLayoutManager._pixelPrecision)) {
                    sizeInItems -= ((sizeInPixels - scrollableArea) / itemContextualWidth) * visibleRatio;
                    break;
                }
            }

            return sizeInItems;
        }

        public isScrollableHeader(item: any, items: any, index: number): boolean {
            if (index !== 0 || this.dimension.getFractionScrollOffset() === 0) {
                return false;
            }

            let hierarchyNavigator: ITablixHierarchyNavigator = this.dimension._hierarchyNavigator;

            if (hierarchyNavigator.isLeaf(item)) {
                return true;
            }

            let currentItem: any = item;
            let currentItems: any[] = items;

            do {
                currentItems = hierarchyNavigator.getChildren(currentItem);
                currentItem = this.dimension.getFirstVisibleItem(hierarchyNavigator.getLevel(currentItem) + 1);

                if (currentItem === undefined) {
                    break;
                }

                if (!hierarchyNavigator.isLastItem(currentItem, currentItems)) {
                    return false;
                }

            } while (!hierarchyNavigator.isLeaf(currentItem));

            return true;
        }

        public reachedEnd(): boolean {
            return this.dimension.getIntegerScrollOffset() + (this.getRealizedItemsCount() - this._gridOffset) >= this.dimension.getItemsCount();
        }

        public scrollBackwardToFill(gridContextualWidth: number): number {
            let newScrollOffset = this.dimension.scrollOffset;
            if (this.reachedEnd()) {
                let widthToFill = this._contextualWidthToFill - gridContextualWidth;
                if (this.dimension.getItemsCount() > 0) {
                    let averageColumnwidth = gridContextualWidth / (this.getRealizedItemsCount() - this.dimension.getFractionScrollOffset());
                    newScrollOffset = this.dimension.getValidScrollOffset(Math.floor(this.dimension.scrollOffset - (widthToFill / averageColumnwidth)));
                }
                this._alignToEnd = !Double.equalWithPrecision(newScrollOffset, this.dimension.scrollOffset, DimensionLayoutManager._scrollOffsetPrecision); // this is an aproximate scrolling back, we have to ensure it is aligned to the end of the control
            }
            return newScrollOffset;
        }

        private getItemContextualWidth(index: number): number {
            let realizedItems = this._getRealizedItems();
            if (index >= realizedItems.length)
                return null;

            return realizedItems[index].getContextualWidth();
        }

        private getItemContextualWidthWithScrolling(index: number): number {
            return this.getSizeWithScrolling(this.getItemContextualWidth(index), index);
        }

        public getSizeWithScrolling(size: number, index: number): number {
            let ratio;

            if (this._gridOffset === index) {
                ratio = this.getVisibleSizeRatio();
            }
            else {
                ratio = 1;
            }

            return size * ratio;
        }

        public getGridContextualWidthFromItems(): number {
            let count = this.getRealizedItemsCount();
            let contextualWidth = 0;
            for (let i = 0; i < count; i++) {
                contextualWidth += this.getItemContextualWidthWithScrolling(i);
            }
            return contextualWidth;
        }

        private getMeaurementError(gridContextualWidth: number): number {
            return gridContextualWidth / this.getGridContextualWidthFromItems();
        }

        private scrollForwardToAlignEnd(gridContextualWidth: number): number {
            let newScrollOffset = this.dimension.scrollOffset;
            if (this._alignToEnd) {
                let withinThreshold = Double.equalWithPrecision(gridContextualWidth, this._contextualWidthToFill, DimensionLayoutManager._pixelPrecision);
                if (!withinThreshold) { // if it is within the threshold we consider it aligned, skip aliging algorithm
                    let count: number = this.getRealizedItemsCount();
                    let startIndex = this._gridOffset;
                    let widthToScroll = gridContextualWidth - this._contextualWidthToFill;

                    let error = this.getMeaurementError(gridContextualWidth);

                    for (let i = startIndex; i < count; i++) {
                        let itemContextualWidth = this.getItemContextualWidth(i) * error;
                        if (Double.lessWithPrecision(itemContextualWidth, widthToScroll, DimensionLayoutManager._pixelPrecision)) {
                            widthToScroll -= itemContextualWidth;
                        }
                        else {
                            let visibleRatio = startIndex === i ? 1 - this.dimension.getFractionScrollOffset() : 1;
                            newScrollOffset = this.dimension.getValidScrollOffset(this.dimension.scrollOffset + (i - startIndex) + (widthToScroll * visibleRatio / itemContextualWidth));
                            break;
                        }
                    }
                }
                this._alignToEnd = !withinThreshold;
            }
            return newScrollOffset;
        }

        public get dimension(): TablixDimension {
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager.dimension");
            return null;
        }

        public get otherLayoutManager(): DimensionLayoutManager {
            return <DimensionLayoutManager>this.dimension.otherDimension.layoutManager;
        }

        public get contextualWidthToFill(): number {
            return this._contextualWidthToFill;
        }

        public getGridScale(): number {
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager.getGridScale");
            return 0;
        }

        public get otherScrollbarContextualWidth(): number {
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager.otherScrollbarContextualWidth");
            return 0;
        }

        public getActualContextualWidth(gridContextualWidth: number): number {
            if (this._isAutoSized() && !this.canScroll(gridContextualWidth))
                return gridContextualWidth;

            return this._contextualWidthToFill;
        }

        protected canScroll(gridContextualWidth: number): boolean {
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager.canScroll");
            return false;
        }

        public calculateSizes(): void {
            if (this.fixedSizeEnabled) {
                this.calculateContextualWidths();
                this.calculateSpans();
            }
        }

        protected _calculateSize(item: ITablixGridItem): number {
            debug.assertFail("PureVirtualMethod: DimensionLayoutManager._calculateSize");
            return null;
        }

        public calculateContextualWidths(): void {
            let items: ITablixGridItem[] = this._getRealizedItems();
            let count = items.length;

            for (let i = 0; i < count; i++) {
                let item: ITablixGridItem = items[i];

                if (this.measureEnabled)
                    item.setAligningContextualWidth(-1);

                this._calculateSize(item);
            }
        }

        public calculateSpans(): void {
            if (this.measureEnabled) {
                this.updateNonScrollableItemsSpans();
                this.updateScrollableItemsSpans();
            }

            // TODO override in row layout manager to add footer to calculation, this is required for Matrix
        }

        public updateNonScrollableItemsSpans(): void {
            let otherDimensionItems = this.otherLayoutManager._getRealizedItems();
            let otherDimensionItemsCount = otherDimensionItems.length;
            let startIndex = this.dimension.getDepth();
            for (let i = startIndex; i < otherDimensionItemsCount; i++) {
                let otherDimensionItem = otherDimensionItems[i];
                this.updateSpans(otherDimensionItem, otherDimensionItem.getHeaders());
            }
        }

        public updateScrollableItemsSpans(): void {
            let otherRealizedItems = this.otherLayoutManager._getRealizedItems();
            let otherRealizedItemsCount = Math.min(this.dimension.getDepth(), otherRealizedItems.length);
            for (let i = 0; i < otherRealizedItemsCount; i++) {
                let otherRealizedItem = otherRealizedItems[i];
                this.updateSpans(otherRealizedItem, otherRealizedItem.getOtherDimensionHeaders());
            }
        }

        public fixSizes(): void {
            if (this.fixedSizeEnabled) {
                let items: ITablixGridItem[] = this._getRealizedItems();
                let count = items.length;
                for (let i = 0; i < count; i++) {
                    items[i].fixSize();
                }
            }
        }

        private updateSpans(otherRealizedItem: ITablixGridItem, cells: TablixCell[]): void {
            let realizedItems = this._getRealizedItems();
            let cellCount = cells.length;
            for (let j = 0; j < cellCount; j++) {
                let cell = cells[j];
                let owner = otherRealizedItem.getOtherDimensionOwner(cell);
                let span = owner.getCellContextualSpan(cell);
                if (span > 1) {
                    let totalSizeInSpan = 0;
                    let startIndex = owner.getIndex(this._grid);
                    for (let k = 0; k < span; k++) {
                        let item = realizedItems[k + startIndex];
                        totalSizeInSpan += this.getSizeWithScrolling(item.getContentContextualWidth(), k + startIndex);
                        if (k === span - 1)
                            this.updateLastChildSize(cell, item, totalSizeInSpan);
                    }
                }
            }
        }

        private updateLastChildSize(spanningCell: TablixCell, item: ITablixGridItem, totalSpanSize: number): void {
            let delta = item.getCellIContentContextualWidth(spanningCell) - totalSpanSize;
            if (delta > 0) // the parent width is wider than the sum of the children, stretch the last child to compensate the difference
                item.setAligningContextualWidth(Math.max(item.getAligningContextualWidth(), delta + item.getContentContextualWidth()));
        }
    }

    export class ResizeState {
        public item: any;
        public itemType: TablixCellType;
        public column: TablixColumn;
        public startColumnWidth: number;
        public resizingDelta: number;
        public animationFrame: number;
        public scale: number;

        constructor(column: TablixColumn, width: number, scale: number) {
            this.column = column;
            this.item = column.getLeafItem();
            this.itemType = column.itemType;
            this.startColumnWidth = width;
            this.resizingDelta = 0;
            this.animationFrame = null;
            this.scale = scale;
        }
    }

    export class ColumnLayoutManager extends DimensionLayoutManager implements ITablixResizeHandler {
        static minColumnWidth = 10;
        private _resizeState: ResizeState;

        constructor(owner: TablixLayoutManager, grid: TablixGrid, realizationManager: ColumnRealizationManager) {
            super(owner, grid, realizationManager);

            realizationManager.owner = this;

            this.fillProportionally = false;
            this._resizeState = null;
        }

        public get dimension(): TablixDimension {
            return this.owner.owner.columnDimension;
        }

        public isResizing(): boolean {
            return this._resizeState !== null;
        }

        public set fillProportionally(value: boolean) {
            this._grid.fillColumnsProportionally = value;
        }

        public get fillProportionally(): boolean {
            return this._grid.fillColumnsProportionally;
        }

        public getGridScale(): number {
            return this._grid._presenter.getScreenToCssRatioX();
        }

        public get otherScrollbarContextualWidth(): number {
            if (this.dimension.otherDimension.scrollbar.visible) {
                return HTMLElementUtils.getElementWidth(this.dimension.otherDimension.scrollbar.element);
            }
            return 0;
        }

        public _getRealizedItems(): ITablixGridItem[] {
            if (!this._grid.realizedColumns) {
                this._grid.realizedColumns = [];
            }

            return this._grid.realizedColumns;
        }

        public _moveElementsToBottom(moveFromIndex: number, count): void {
            this._grid.moveColumnsToEnd(moveFromIndex, count);
        }

        public _moveElementsToTop(moveToIndex: number, count): void {
            this._grid.moveColumnsToStart(moveToIndex, count);
        }

        public _requiresMeasure(): boolean {
            // if the control is not scrolling in either dimension or is scrolling or is resizing
            return (!this.isScrolling() && !this.otherLayoutManager.isScrolling()) || this.isScrolling() || this.isResizing();
        }

        public getGridContextualWidth(): number {
            return this._grid.getWidth();
        }

        private getFirstVisibleColumn(): TablixColumn {
            return this._grid.realizedColumns[this._gridOffset];
        }

        public _isAutoSized(): boolean {
            return this.owner.owner.autoSizeWidth;
        }

        public applyScrolling() {
            let columnOffset: number = this.dimension.getFractionScrollOffset();
            let firstVisibleColumnWidth: number = 0;

            if (columnOffset !== 0) {
                let firstVisibleColumn: TablixColumn = this.getFirstVisibleColumn();
                if (firstVisibleColumn !== undefined) {
                    firstVisibleColumnWidth = firstVisibleColumn.getContextualWidth();
                    this.scroll(firstVisibleColumn, firstVisibleColumnWidth, columnOffset);
                }
            }
        }

        private scroll(firstVisibleColumn: TablixColumn, width: number, offset: number) {
            this.scrollCells(firstVisibleColumn._realizedColumnHeaders, width, offset);
            this.scrollBodyCells(this._grid.realizedRows, width, offset);

            if (firstVisibleColumn.footer !== null) {
                firstVisibleColumn.footer.scrollHorizontally(width, offset);
            }
        }

        private scrollCells(cells: TablixCell[], width: number, offset: number): void {
            let length: number = cells.length;
            for (let i = 0; i < length; i++) {
                cells[i].scrollHorizontally(width, offset);
            }
        }

        private scrollBodyCells(rows: TablixRow[], width: number, offset: number): void {
            let length: number = rows.length;
            let cells: TablixCell[];
            let cell: TablixCell;
            for (let i = 0; i < length; i++) {
                cells = rows[i]._realizedBodyCells;
                if (cells !== undefined) {
                    cell = cells[0];
                    if (cell !== undefined) {
                        cell.scrollHorizontally(width, offset);
                    }
                }
            }
        }

        public onStartResize(cell: TablixCell, currentX: number, currentY: number): void {
            this._resizeState = new ResizeState(cell._column, cell._column.getContentContextualWidth(), HTMLElementUtils.getAccumulatedScale(this.owner.owner.container));
        }

        public onResize(cell: TablixCell, deltaX: number, deltaY: number): void {
            this._resizeState.resizingDelta = Math.max(deltaX / this._resizeState.scale, ColumnLayoutManager.minColumnWidth - this._resizeState.startColumnWidth);
            if (this._resizeState.animationFrame === null)
                this._resizeState.animationFrame = requestAnimationFrame(() => this.performResizing());
        }

        public onEndResize(cell: TablixCell): void {
            if (this._resizeState.animationFrame !== null) {
                this.performResizing(); // if we reached the end and we are still waiting for the last animation frame, perform the pending resizing and clear the state 
            }
            this._resizeState = null;
        }

        public onReset(cell: TablixCell) {
            this._resizeState = new ResizeState(cell._column, -1, 1);
            cell._column.clearSize();
            this.owner.owner.refresh(false);
            this._resizeState = null;
        }

        public updateItemToResizeState(realizedColumns: TablixColumn[]): void {
            if (this._resizeState === null)
                return;

            let columnCount: number = realizedColumns.length;

            let hierarchyNavigator: ITablixHierarchyNavigator = this.owner.owner.hierarchyNavigator;

            // Only iterate over the columns that belong to column hierachy (i.e. skip the row hierarchy rows)
            // as this post-rendering adjustment only applies to them.
            let startIndex = this.otherLayoutManager.dimension.getDepth();
            for (let i = startIndex; i < columnCount; i++) {
                let column: TablixColumn = realizedColumns[i];

                if (!column.columnHeaderOrCornerEquals(this._resizeState.itemType, this._resizeState.item, column.itemType, column.getLeafItem(), hierarchyNavigator))
                    continue;

                if (column !== this._resizeState.column) {  // we moved the item of the column that is being resized to another one 
                    this._resizeState.column = column;
                    column.resize(this._resizeState.startColumnWidth + this._resizeState.resizingDelta);
                    break;
                }
            }
        }

        private performResizing(): void {
            if (this._resizeState === null) // in case of FireFox we cannot cancel the animation frame request
                return;

            this._resizeState.animationFrame = null;
            let newSize = this._resizeState.startColumnWidth + this._resizeState.resizingDelta;
            this._resizeState.column.resize(newSize);
            this.owner.owner.refresh(false);
        }

        /**
         * Sends column related data (pixel size, column count, etc) to TablixControl.
         */
        public _sendDimensionsToControl(): void {
            let gridContextualWidth: number = this.getGridContextualWidth();
            let widthToFill: number = this.getActualContextualWidth(gridContextualWidth);
            let otherContextualHeight = this.getOtherHierarchyContextualHeight();
            let scale = this.getGridScale(); // in case of canvas we have to convert the size from device pixel to css pixel
            this.owner.owner.updateColumnDimensions(otherContextualHeight / scale,
                (widthToFill - otherContextualHeight) / scale,
                this.getViewSize(gridContextualWidth));
        }

        public getEstimatedHeaderWidth(label: string, headerIndex: number): number {
            debug.assertFail("PureVirtualMethod: ColumnLayoutManager.getEstimatedHeaderWidth");
            return -1;
        }

        public getEstimatedBodyCellWidth(content: string): number {
            debug.assertFail("PureVirtualMethod: ColumnLayoutManager.getEstimatedBodyCellWidth");
            return -1;
        }
    }

    export class DashboardColumnLayoutManager extends ColumnLayoutManager {
        public getEstimatedHeaderWidth(label: string, headerIndex: number): number {
            if (this.ignoreColumn(headerIndex))
                return 0;

            // for dashboard layout it does not matter whether we pass an actual cell or not
            return this.owner.getCellWidth(undefined);
        }

        public getEstimatedBodyCellWidth(content: string): number {
            // for dashboard layout it does not matter whether we pass an actual cell or not
            return this.owner.getCellWidth(undefined);
        }

        protected canScroll(gridContextualWidth: number): boolean {
            return false;
        }

        protected _calculateSize(item: ITablixGridItem): number {
            let headerIndex = item.getIndex(this._grid);
            let computedSize = 0;

            if (!this.ignoreColumn(headerIndex)) {
                // for dashboard layout it does not matter whether we pass an actual cell or not
                computedSize = this.owner.getContentWidth(undefined);
            }

            item.resize(computedSize);
            return computedSize;
        }

        private ignoreColumn(headerIndex: number): boolean {
            // On the dashboard, we need to return 0 if the row header is static
            // (a table or a matrix without row groups)
            return headerIndex === 0 && !this.owner.binder.hasRowGroups();
        }
    }

    export class CanvasColumnLayoutManager extends ColumnLayoutManager {
        public getEstimatedHeaderWidth(label: string, headerIndex: number): number {
            // On the canvas the header width depends on the size of the content
            return this.owner.getEstimatedTextWidth(label);
        }

        public getEstimatedBodyCellWidth(content: string): number {
            return this.owner.getEstimatedTextWidth(content);
        }

        public calculateContextualWidths(): void {
            let items: ITablixGridItem[] = this._getRealizedItems();
            let columnWidths: number[] = [];

            for (let item of items) {
                if (this.measureEnabled)
                    item.setAligningContextualWidth(-1);

                columnWidths.push(this._calculateSize(item));
            }
            
            // Save all column widths. Needed when user turns off auto-sizing for column widths.
            this.owner.columnWidthsToPersist = columnWidths;
        }

        protected canScroll(gridContextualWidth: number): boolean {
            return !Double.equalWithPrecision(this.dimension.scrollOffset, 0, DimensionLayoutManager._scrollOffsetPrecision) ||
                (((this.getRealizedItemsCount() - this._gridOffset) < this.dimension.getItemsCount()) && this._contextualWidthToFill > 0) ||
                Double.greaterWithPrecision(gridContextualWidth, this._contextualWidthToFill, DimensionLayoutManager._pixelPrecision);
        }

        protected _calculateSize(item: ITablixGridItem): number {
            return item.calculateSize();
        }
    }

    export class RowLayoutManager extends DimensionLayoutManager {
        constructor(owner: TablixLayoutManager, grid: TablixGrid, realizationManager: RowRealizationManager) {
            super(owner, grid, realizationManager);

            realizationManager.owner = this;
        }

        public get dimension(): TablixDimension {
            return this.owner.owner.rowDimension;
        }

        public getGridScale(): number {
            return this._grid._presenter.getScreenToCssRatioY();
        }

        public get otherScrollbarContextualWidth(): number {
            if (this.dimension.otherDimension.scrollbar.visible) {
                return HTMLElementUtils.getElementHeight(this.dimension.otherDimension.scrollbar.element);
            }
            return 0;
        }

        public startScrollingSession(): void {
            super.startScrollingSession();
        }

        public _getRealizedItems(): ITablixGridItem[] {
            if (!this._grid.realizedRows) {
                this._grid.realizedRows = [];
            }

            return this._grid.realizedRows;
        }

        public _moveElementsToBottom(moveFromIndex: number, count): void {
            this._grid.moveRowsToEnd(moveFromIndex, count);
        }

        public _moveElementsToTop(moveToIndex: number, count): void {
            this._grid.moveRowsToStart(moveToIndex, count);
        }

        public _requiresMeasure(): boolean {
            // if the control is not scrolling in either dimension and the column dimension is not resizing or row dimension is scrolling and reaching the end while scrolling 
            return (!this.isScrolling() && !this.otherLayoutManager.isScrolling() && !this.otherLayoutManager.isResizing())
                || (this.isScrolling() && (this.dimension.getIntegerScrollOffset() + (this.getRealizedItemsCount() - this._gridOffset) >= this.dimension.getItemsCount()));
        }

        public getGridContextualWidth(): number {
            return this._grid.getHeight();
        }

        private getFirstVisibleRow(): TablixRow {
            return this._grid.realizedRows[this._gridOffset];
        }

        public _isAutoSized(): boolean {
            return this.owner.owner.autoSizeHeight;
        }

        public applyScrolling() {
            let rowOffset: number = this.dimension.getFractionScrollOffset();
            let firstVisibleRowHeight: number = 0;

            if (rowOffset !== 0) {
                let firstVisibleRow: TablixRow = this.getFirstVisibleRow();
                if (firstVisibleRow) {
                    firstVisibleRowHeight = firstVisibleRow.getContextualWidth();
                    this.scroll(firstVisibleRow, firstVisibleRowHeight, rowOffset);
                }
            }
        }

        private scroll(firstVisibleRow: TablixRow, height: number, offset: number) {
            this.scrollCells(firstVisibleRow._realizedRowHeaders, height, offset);
            this.scrollCells(firstVisibleRow._realizedBodyCells, height, offset);
        }

        private scrollCells(cells: TablixCell[], height: number, offset: number): void {
            let length: number = cells.length;
            for (let i = 0; i < length; i++) {
                cells[i].scrollVertically(height, offset);
            }
        }

        public getFooterContextualWidth(): number {
            if (this.owner.owner.rowDimension.hasFooter()) {
                if (this.owner.grid.footerRow) {
                    return this.owner.grid.footerRow.getContextualWidth();
                }
            }

            return 0;
        }

        public calculateContextualWidths(): void {
            super.calculateContextualWidths();
            if (this.fixedSizeEnabled) {
                let footerRow: TablixRow = this._grid.footerRow;
                if (footerRow) {
                    this._calculateSize(footerRow);
                }
            }
        }

        public fixSizes(): void {
            super.fixSizes();
            if (this.fixedSizeEnabled) {
                if (this._grid.footerRow) {
                    this._grid.footerRow.fixSize();
                }
            }
        }

        /**
         * Sends row related data (pixel size, column count, etc) to TablixControl.
         */
        public _sendDimensionsToControl(): void {
            let gridContextualWidth: number = this.getGridContextualWidth();
            let widthToFill: number = this.getActualContextualWidth(gridContextualWidth);
            let otherContextualHeight = this.getOtherHierarchyContextualHeight();
            let scale = this.getGridScale();
            this.owner.owner.updateRowDimensions(otherContextualHeight / scale,
                (widthToFill - otherContextualHeight) / scale,
                gridContextualWidth / scale,
                this.getViewSize(gridContextualWidth),
                (this._grid.footerRow ? this._grid.footerRow.getContextualWidth() / scale : 0));
        }

        public getEstimatedHeaderWidth(label: string, headerIndex: number): number {
            debug.assertFail("PureVirtualMethod: RowLayoutManager.getEstimatedHeaderWidth");
            return -1;
        }
    }

    export class DashboardRowLayoutManager extends RowLayoutManager {
        public getEstimatedHeaderWidth(label: string, headerIndex: number): number {
            return this.getHeaderWidth(headerIndex);
        }

        protected canScroll(gridContextualWidth: number): boolean {
            return false;
        }

        protected _calculateSize(item: ITablixGridItem): number {
            let computedSize = this.owner.getEstimatedRowHeight();
            item.resize(computedSize);
            return computedSize;
        }

        private getHeaderWidth(headerIndex: number): number {
            // On the dashboard, we need to return 0 if the row header is static
            // (a table or a matrix without row groups)
            if (headerIndex === 0 && !this.owner.binder.hasRowGroups())
                return 0;

            // for dashboard layout it does not matter whether we pass an actual text or not
            return this.owner.getEstimatedTextWidth(undefined);
        }
    }

    export class CanvasRowLayoutManager extends RowLayoutManager {
        public getEstimatedHeaderWidth(label: string, headerIndex: number): number {
            // On the canvas the header width depends on the size of the content
            return this.owner.getEstimatedTextWidth(label);
        }

        protected canScroll(gridContextualWidth: number): boolean {
            return !Double.equalWithPrecision(this.dimension.scrollOffset, 0, DimensionLayoutManager._scrollOffsetPrecision) ||
                (((this.getRealizedItemsCount() - this._gridOffset) < this.dimension.getItemsCount()) && this._contextualWidthToFill > 0) ||
                Double.greaterWithPrecision(gridContextualWidth, this._contextualWidthToFill, DimensionLayoutManager._pixelPrecision);
        }

        protected _calculateSize(item: ITablixGridItem): number {
            return item.calculateSize();
        }
    }

    export class TablixLayoutManager {
        protected _owner: TablixControl;
        protected _container: HTMLElement;
        protected _columnLayoutManager: ColumnLayoutManager;
        protected _rowLayoutManager: RowLayoutManager;
        private _binder: ITablixBinder;

        private _scrollingDimension: TablixDimension;
        private _gridHost: HTMLElement;
        private _footersHost: HTMLElement;
        private _grid: internal.TablixGrid;
        private _allowHeaderResize: boolean = true;
        private _columnWidthsToPersist: number[];

        constructor(
            binder: ITablixBinder,
            grid: TablixGrid,
            columnLayoutManager: ColumnLayoutManager,
            rowLayoutManager: RowLayoutManager) {
            this._binder = binder;
            this._grid = grid;
            this._columnLayoutManager = columnLayoutManager;
            this._rowLayoutManager = rowLayoutManager;
            this._columnWidthsToPersist = [];
        }

        public initialize(owner: TablixControl): void {
            this._owner = owner;
            this._container = owner.container;
            this._gridHost = owner.contentHost;
            this._footersHost = owner.footerHost;
            this._grid.initialize(owner, this._gridHost, this._footersHost);
        }

        public get owner(): TablixControl {
            return this._owner;
        }

        public get binder(): ITablixBinder {
            return this._binder;
        }

        public get columnWidthsToPersist(): number[] {
            return this._columnWidthsToPersist;
        }

        public set columnWidthsToPersist(columnWidths: number[]) {
            this._columnWidthsToPersist = columnWidths;
        }        

        public getTablixClassName(): string {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.getTablixClassName");
            return null;
        }

        public getLayoutKind(): TablixLayoutKind {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.getLayoutKind");

            // TODO ckerer: this method should not be necessary when we are done refactoring!
            return null;
        }

        public getOrCreateColumnHeader(item: any, items: any, rowIndex: number, columnIndex: number): ITablixCell {
            let row: TablixRow = this._grid.getOrCreateRow(rowIndex);
            let column: TablixColumn = this._grid.getOrCreateColumn(columnIndex + this._columnLayoutManager._gridOffset);
            let isLeaf = this.owner.hierarchyNavigator.isLeaf(item);
            let cell: TablixCell = row.getOrCreateColumnHeader(column, this._columnLayoutManager.isScrollableHeader(item, items, columnIndex), isLeaf);
            this.enableCellHorizontalResize(isLeaf, cell);
            return cell;
        }

        public getOrCreateRowHeader(item: any, items: any, rowIndex: number, columnIndex: number): ITablixCell {
            let row: TablixRow = this._grid.getOrCreateRow(rowIndex + this._rowLayoutManager._gridOffset);
            let column: TablixColumn = this._grid.getOrCreateColumn(columnIndex);
            let scrollable: boolean = this._rowLayoutManager.isScrollableHeader(item, items, rowIndex);

            if (row.getRealizedCellCount() === 0) {
                this.alignRowHeaderCells(item, row);
            }

            let cell: TablixCell = row.getOrCreateRowHeader(column, scrollable, this.owner.hierarchyNavigator.isLeaf(item));
            cell.enableHorizontalResize(false, this._columnLayoutManager);
            return cell;
        }

        public getOrCreateCornerCell(item: any, rowLevel: number, columnLevel: number): ITablixCell {
            let row: TablixRow = this._grid.getOrCreateRow(columnLevel);
            let column: TablixColumn = this._grid.getOrCreateColumn(rowLevel);
            let cell: TablixCell = row.getOrCreateCornerCell(column);
            let columnDepth = this._columnLayoutManager.dimension.getDepth();
            let isLeaf = columnLevel === (columnDepth - 1);
            this.enableCellHorizontalResize(isLeaf, cell);
            return cell;
        }

        public getOrCreateBodyCell(cellItem: any, rowItem: any, rowItems: any, rowIndex: number, columnIndex: number): ITablixCell {
            let scrollable: boolean;
            let row: TablixRow = this._grid.getOrCreateRow(rowIndex + this._rowLayoutManager._gridOffset);
            let column: TablixColumn = this._grid.getOrCreateColumn(columnIndex + this._columnLayoutManager._gridOffset);

            if (row._realizedBodyCells.length === 0 && this._owner.columnDimension.getFractionScrollOffset() !== 0) {
                scrollable = true;
            }
            else {
                scrollable = this._rowLayoutManager.isScrollableHeader(rowItem, rowItems, rowIndex);
            }

            let cell: TablixCell = row.getOrCreateBodyCell(column, scrollable);
            cell.enableHorizontalResize(false, this._columnLayoutManager);
            return cell;
        }

        public getOrCreateFooterBodyCell(cellItem: any, columnIndex: number): ITablixCell {
            let scrollable: boolean;
            let row: TablixRow = this._grid.getOrCreateFootersRow();
            let column: TablixColumn = this._grid.getOrCreateColumn(columnIndex + this._columnLayoutManager._gridOffset);

            scrollable = (row._realizedBodyCells.length === 0 && this._owner.columnDimension.getFractionScrollOffset() !== 0);

            let cell: TablixCell = row.getOrCreateFooterBodyCell(column, scrollable);
            cell.enableHorizontalResize(false, this._columnLayoutManager);
            return cell;
        }

        public getOrCreateFooterRowHeader(item: any, items: any): ITablixCell {
            let row: TablixRow = this._grid.getOrCreateFootersRow();
            let column: TablixColumn = this._grid.getOrCreateColumn(0);

            //debug.assert(this.owner.hierarchyNavigator.isLeaf(item), "Leaf item expected");

            let cell: TablixCell = row.getOrCreateFooterRowHeader(column);
            cell.enableHorizontalResize(false, this._columnLayoutManager);
            return cell;
        }

        public getVisibleWidth(): number {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.getVisibleWidth");
            return -1;
        }

        public getVisibleHeight(): number {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.getVisibleHeight");
            return -1;
        }

        public updateColumnCount(rowDimension: TablixRowDimension, columnDimension: TablixColumnDimension) {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.updateColumnCount");
        }

        public updateViewport(viewport: IViewport) {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.updateViewport");
        }

        public getEstimatedRowHeight(): number {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.getEstimatedRowHeight");
            return -1;
        }

        public getCellWidth(cell: ITablixCell): number {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.getCellWidth");
            return -1;
        }

        public getContentWidth(cell: ITablixCell): number {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.getContentWidth");
            return -1;
        }

        public adjustContentSize(hasImage: boolean): void {
            // default implementation has no adjustment
        }

        /**
         * This call makes room for parent header cells where neccessary.
         * Since HTML cells that span vertically displace other rows,
         * room has to be made for spanning headers that leave an exiting 
         * row to enter the new row that it starts from and removed when
         * returning to an entering row.
         */
        private alignRowHeaderCells(item: any, currentRow: TablixRow): void {
            let index = currentRow.getRowHeaderLeafIndex();

            if (index === -1) {
                return;
            }

            let rowDimension: TablixRowDimension = this._owner.rowDimension;
            let leaf = rowDimension.getFirstVisibleChildLeaf(item);

            if (!this.owner.hierarchyNavigator.headerItemEquals(leaf, currentRow.getAllocatedCellAt(index).item)) {
                return;
            }

            currentRow.moveCellsBy(this.owner.hierarchyNavigator.getLevel(leaf) - this.owner.hierarchyNavigator.getLevel(item) - index);
        }

        public get grid(): TablixGrid {
            return this._grid;
        }

        public get rowLayoutManager(): DimensionLayoutManager {
            return this._rowLayoutManager;
        }

        public get columnLayoutManager(): DimensionLayoutManager {
            return this._columnLayoutManager;
        }

        protected showEmptySpaceHeader(): boolean {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.showEmptySpaceHeader");
            return false;
        }

        public onStartRenderingSession(scrollingDimension: TablixDimension, parentElement: HTMLElement, clear: boolean): void {
            if (this.showEmptySpaceHeader()) {
                let cell: ITablixCell = this._grid.emptySpaceHeaderCell;
                if (cell) {
                    this._binder.unbindEmptySpaceHeaderCell(cell);
                }

                cell = this._grid.emptySpaceFooterCell;
                if (cell) {
                    this._binder.unbindEmptySpaceFooterCell(cell);
                }

                this._grid.HideEmptySpaceCells();
            }

            this._scrollingDimension = scrollingDimension;

            if (this._scrollingDimension) {
                (<DimensionLayoutManager>this._scrollingDimension.layoutManager).startScrollingSession();
            }

            this._rowLayoutManager.onStartRenderingSession();
            this._columnLayoutManager.onStartRenderingSession();
            this._grid.onStartRenderingSession(clear);

            let measureEnabled = this._columnLayoutManager.measureEnabled || this._rowLayoutManager.measureEnabled;
            if (measureEnabled)
                this.measureSampleText(parentElement);
        }

        public onEndRenderingSession(): void {
            this._rowLayoutManager.onEndRenderingSession();
            this._columnLayoutManager.onEndRenderingSession();

            if (this._scrollingDimension) {
                (<DimensionLayoutManager>this._scrollingDimension.layoutManager).endScrollingSession();
            }

            this._scrollingDimension = null;

            if (this.showEmptySpaceHeader()) {
                let emptySpace = this._columnLayoutManager.contextualWidthToFill - this._columnLayoutManager.getGridContextualWidth();
                if (emptySpace > 0) {
                    this._grid.ShowEmptySpaceCells(this._owner.columnDimension.getDepth(), emptySpace);

                    let cell: ITablixCell = this._grid.emptySpaceHeaderCell;
                    if (cell) {
                        this._binder.bindEmptySpaceHeaderCell(cell);
                    }
                    cell = this._grid.emptySpaceFooterCell;
                    if (cell) {
                        this._binder.bindEmptySpaceFooterCell(cell);
                    }
                }
            }
        }

        public onStartRenderingIteration(clear: boolean): void {
            this._rowLayoutManager.onStartRenderingIteration(clear, this.getVisibleHeight());
            this._columnLayoutManager.onStartRenderingIteration(clear, this.getVisibleWidth());
            this._grid.onStartRenderingIteration();
        }

        public onEndRenderingIteration(): boolean {
            this._grid.onEndRenderingIteration();

            // ANDREMI: Comment out for static tablix
            this._columnLayoutManager.calculateSizes(); // calculate the entire grid first without altering the tree to avoid multiple measure pass invoking
            this._rowLayoutManager.calculateSizes();
            this._columnLayoutManager.fixSizes();  // now assign the sizes
            this._rowLayoutManager.fixSizes();

            this._columnLayoutManager.updateItemToResizeState(this._grid.realizedColumns); // if we are in a middle of a resize, the column to resize might have been swaped during the render, restore its resize state

            this._columnLayoutManager.applyScrolling();
            this._rowLayoutManager.applyScrolling();
            this._columnLayoutManager.onEndRenderingIteration();
            this._rowLayoutManager.onEndRenderingIteration();

            return this._columnLayoutManager.done && this._rowLayoutManager.done;
        }

        public onCornerCellRealized(item: any, cell: ITablixCell): void {
            let columnLeaf: boolean = this.owner.hierarchyNavigator.isColumnHierarchyLeaf(item);
            let rowLeaf: boolean = this.owner.hierarchyNavigator.isRowHierarchyLeaf(item);
            if (columnLeaf)
                (<TablixCell>cell)._column.OnLeafRealized(this._owner.hierarchyNavigator);
            this._columnLayoutManager.onCornerCellRealized(item, cell, columnLeaf);
            this._rowLayoutManager.onCornerCellRealized(item, cell, rowLeaf);
        }

        public onRowHeaderRealized(item: any, cell: ITablixCell): void {
            let hierarchyNavigator: ITablixHierarchyNavigator = this._owner.hierarchyNavigator;
            let leaf = hierarchyNavigator.isLeaf(item);
            this._rowLayoutManager.onHeaderRealized(item, cell, leaf);
        }

        public onRowHeaderFooterRealized(item: any, cell: ITablixCell): void {
        }

        public onColumnHeaderRealized(item: any, cell: ITablixCell): void {
            let hierarchyNavigator: ITablixHierarchyNavigator = this._owner.hierarchyNavigator;
            let leaf = hierarchyNavigator.isLeaf(item);
            if (leaf)
                (<TablixCell>cell)._column.OnLeafRealized(this._owner.hierarchyNavigator);
            this._columnLayoutManager.onHeaderRealized(item, cell, leaf);
        }

        public onBodyCellRealized(item: any, cell: ITablixCell): void {
        }

        public onBodyCellFooterRealized(item: any, cell: ITablixCell): void {
        }

        public setAllowHeaderResize(value: boolean) {
            this._allowHeaderResize = value;
        }

        public enableCellHorizontalResize(isLeaf: boolean, cell: TablixCell) {
            let enableCellHorizontalResize = isLeaf && this._allowHeaderResize;
            cell.enableHorizontalResize(enableCellHorizontalResize, this._columnLayoutManager);
        }

        public getEstimatedTextWidth(label: string): number {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.getEstimatedTextWidth");
            return -1;
        }

        public measureSampleText(parentElement: HTMLElement): void {
            debug.assertFail("PureVirtualMethod: TablixLayoutManager.measureSampleText");
        }
    }

    export class DashboardTablixLayoutManager extends TablixLayoutManager {
        private _sizeComputationManager: SizeComputationManager;

        constructor(
            binder: ITablixBinder,
            sizeComputationManager: SizeComputationManager,
            grid: TablixGrid,
            rowRealizationManager: RowRealizationManager,
            columnRealizationManager: ColumnRealizationManager) {
            super(
                binder,
                grid,
                new DashboardColumnLayoutManager(this, grid, columnRealizationManager),
                new DashboardRowLayoutManager(this, grid, rowRealizationManager));
            this._sizeComputationManager = sizeComputationManager;
        }

        public static createLayoutManager(binder: ITablixBinder): DashboardTablixLayoutManager {
            // computed sizes are shared between layout manager and grid presenter
            let sizeComputationManager = new SizeComputationManager();
            return new DashboardTablixLayoutManager(
                binder,
                sizeComputationManager,
                new TablixGrid(new DashboardTablixGridPresenter(sizeComputationManager), undefined /* Dashboad columns have fixed sizes. */),
                new RowRealizationManager(binder),
                new ColumnRealizationManager(binder));
        }

        public getTablixClassName(): string {
            return "bi-dashboard-tablix";
        }

        public getLayoutKind(): TablixLayoutKind {
            return TablixLayoutKind.DashboardTile;
        }

        protected showEmptySpaceHeader(): boolean {
            return false;
        }

        public measureSampleText(parentElement: HTMLElement): void {
            // Dashboard layout does not use DOM measurements; nothing to do
        }

        public getVisibleWidth(): number {
            return this._sizeComputationManager.visibleWidth;
        }

        public getVisibleHeight(): number {
            return this._sizeComputationManager.visibleHeight;
        }

        public getCellWidth(cell: ITablixCell): number {
            return this._sizeComputationManager.cellWidth;
        }

        public getContentWidth(cell: ITablixCell): number {
            return this._sizeComputationManager.contentWidth;
        }

        public getEstimatedTextWidth(label: string): number {
            // On the dashboard it does not matter what text we render, 
            // we always use the same content width
            return this._sizeComputationManager.contentWidth;
        }

        public adjustContentSize(hasImage: boolean): void {
            this._sizeComputationManager.hasImageContent = hasImage;
        }

        public updateColumnCount(rowDimension: TablixRowDimension, columnDimension: TablixColumnDimension): void {
            // The total number of columns is the number (depth) of row groups + the number of (leaf) column group instances
            let rowDimensionDepth = rowDimension ? rowDimension.getDepth() : 0;
            let columnInstances = columnDimension ? columnDimension.getItemsCount() : 0;
            let totalColumnCount = rowDimensionDepth + columnInstances;

            // Adjust the column count by the static row header (if any)
            if (!this.binder.hasRowGroups())
                totalColumnCount--;

            this._sizeComputationManager.updateColumnCount(totalColumnCount);
        }

        public updateViewport(viewport: IViewport) {
            this._sizeComputationManager.updateViewport(viewport);
        }

        public getEstimatedRowHeight(): number {
            return this._sizeComputationManager.cellHeight;
        }
    }

    export class CanvasTablixLayoutManager extends TablixLayoutManager {
        private characterWidth: number;
        private characterHeight: number;

        constructor(
            binder: ITablixBinder,
            grid: TablixGrid,
            rowRealizationManager: RowRealizationManager,
            columnRealizationManager: ColumnRealizationManager,
            columnWidthChangedCallback: ColumnWidthCallbackType) {
            super(
                binder,
                grid,
                new CanvasColumnLayoutManager(this, grid, columnRealizationManager),
                new CanvasRowLayoutManager(this, grid, rowRealizationManager));
        }

        public static createLayoutManager(binder: ITablixBinder, columnWidthsCallback: () => number[], columnWidthChangedCallback: ColumnWidthCallbackType): CanvasTablixLayoutManager {
            return new CanvasTablixLayoutManager(
                binder,
                new TablixGrid(new controls.internal.CanvasTablixGridPresenter(columnWidthsCallback), columnWidthChangedCallback),
                new RowRealizationManager(binder),
                new ColumnRealizationManager(binder),
                columnWidthChangedCallback);
        }

        public getTablixClassName(): string {
            return "bi-tablix";
        }

        public getLayoutKind(): TablixLayoutKind {
            return TablixLayoutKind.Canvas;
        }

        public measureSampleText(parentElement: HTMLElement): void {
            // TODO: Use TextMeasurementService once the DOM methods are fixed (they are not working right now)
            let textDiv = controls.internal.TablixUtils.createDiv();
            textDiv.style.cssFloat = 'left';
            parentElement.appendChild(textDiv);

            let textNode = document.createTextNode("a");
            textDiv.appendChild(textNode);

            this.characterWidth = controls.HTMLElementUtils.getElementWidth(textDiv);
            this.characterHeight = controls.HTMLElementUtils.getElementHeight(textDiv);

            textDiv.removeChild(textNode);
            parentElement.removeChild(textDiv);
        }

        protected showEmptySpaceHeader(): boolean {
            return !this._columnLayoutManager.fillProportionally;
        }

        public getVisibleWidth(): number {
            if (this._columnLayoutManager.measureEnabled) {
                if (this._owner.autoSizeWidth && this._owner.maxWidth) {
                    return this._owner.maxWidth;
                } else {
                    return HTMLElementUtils.getElementWidth(this._container);
                }
            }

            return -1;
        }

        public getVisibleHeight(): number {
            if (this._rowLayoutManager.measureEnabled) {
                if (this._owner.autoSizeHeight && this._owner.maxHeight) {
                    return this._owner.maxHeight;
                } else {
                    return HTMLElementUtils.getElementHeight(this._container);
                }
            }

            return -1;
        }

        public getCellWidth(cell: ITablixCell): number {
            return HTMLElementUtils.getElementWidth((<TablixCell>cell)._presenter.tableCell);
        }

        public getContentWidth(cell: ITablixCell): number {
            return HTMLElementUtils.getElementWidth((<TablixCell>cell)._presenter.contentElement);
        }

        public getEstimatedTextWidth(text: string): number {
            return text ? text.length * this.characterWidth : 0;
        }

        public updateColumnCount(rowDimension: TablixRowDimension, columnDimension: TablixColumnDimension) {
            // We currently only need to update model information when using dashboard layouts
        }

        public updateViewport(viewport: IViewport) {
            // We currently only need to update model information when using dashboard layouts
        }

        public getEstimatedRowHeight(): number {
            return this.characterHeight;
        }
    }
}