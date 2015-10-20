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

/// <reference path="_references.ts"/>

module powerbi.visuals {

    import TouchUtils = powerbi.visuals.controls.TouchUtils;

    export interface TooltipDataItem {
        displayName: string;
        value: string;
    }

    export interface TooltipOptions {
        opacity: number;
        animationDuration: number;
        offsetX: number;
        offsetY: number;
    }

    export interface TooltipEnabledDataPoint {
        tooltipInfo?: TooltipDataItem[];
    }

    export interface TooltipCategoryDataItem {
        value?: any;
        metadata: DataViewMetadataColumn;
    }

    export interface TooltipSeriesDataItem {
        value?: any;
        highlightedValue?: any;
        metadata: DataViewValueColumn;
    }

    export interface TooltipLocalizationOptions {
        highlightedValueDisplayName: string;
    }

    export interface TooltipEvent {
        data: any;
        index: number;
        coordinates: number[];
        elementCoordinates: number[];
        context: any;
        isTouchEvent: boolean;
    }

    const enum ScreenArea {
        TopLeft,
        TopRight,
        BottomRight,
        BottomLeft
    };

    export class ToolTipComponent {

        private static DefaultTooltipOptions: TooltipOptions = {
            opacity: 1,
            animationDuration: 250,
            offsetX: 10,
            offsetY: 10
        };

        private tooltipContainer: D3.Selection;
        private isTooltipVisible: boolean = false;

        private customScreenWidth: number;
        private customScreenHeight: number;

        private static containerClassName: string = "tooltip-container";
        private static contentContainerClassName: string = "tooltip-content-container";
        private static arrowClassName: string = "arrow";
        private static tooltipRowClassName: string = "tooltip-row";
        private static tooltipTitleCellClassName: string = "tooltip-title-cell";
        private static tooltipValueCellClassName: string = "tooltip-value-cell";

        public static parentContainerSelector: string = "body";
        public static highlightedValueDisplayNameResorceKey: string = "Tooltip_HighlightedValueDisplayName";
        public static localizationOptions: TooltipLocalizationOptions;

        constructor(public tooltipOptions?: TooltipOptions) {
            if (!tooltipOptions) {
                this.tooltipOptions = ToolTipComponent.DefaultTooltipOptions;
            }
        }

        /** Note: For tests only */
        public isTooltipComponentVisible(): boolean {
            return this.isTooltipVisible;
        }

        /** Note: For tests only */
        public setTestScreenSize(width: number, height: number) {
            this.customScreenWidth = width;
            this.customScreenHeight = height;
        }

        public show(tooltipData: TooltipDataItem[], clickedArea: TouchUtils.Rectangle) {
            this.isTooltipVisible = true;

            if (!this.tooltipContainer) {
                this.tooltipContainer = this.createTooltipContainer();
            }

            this.setTooltipContent(tooltipData);

            this.tooltipContainer
                .style("visibility", "visible")
                .transition()
                .duration(0) // Cancel previous transitions
                .style("opacity", this.tooltipOptions.opacity);

            this.setPosition(clickedArea);
        }

        public move(tooltipData: TooltipDataItem[], clickedArea: TouchUtils.Rectangle) {
            if (this.isTooltipVisible) {
                if (tooltipData) {
                    this.setTooltipContent(tooltipData);
                }

                this.setPosition(clickedArea);
            }
        }

        public hide() {
            if (this.isTooltipVisible) {
                this.isTooltipVisible = false;
                this.tooltipContainer
                    .transition()
                    .duration(this.tooltipOptions.animationDuration)
                    .style("opacity", 0)
                    .each('end', function () { this.style.visibility = "hidden"; });
            }
        }

        private createTooltipContainer(): D3.Selection {
            let container: D3.Selection = d3.select(ToolTipComponent.parentContainerSelector)
                .append("div")
                .attr("class", ToolTipComponent.containerClassName);

            container.append("div").attr("class", ToolTipComponent.arrowClassName);
            container.append("div").attr("class", ToolTipComponent.contentContainerClassName);

            return container;
        }

        private setTooltipContent(tooltipData: TooltipDataItem[]): void {

            let rowsSelector: string = "." + ToolTipComponent.tooltipRowClassName;
            let contentContainer = this.tooltipContainer.select("." + ToolTipComponent.contentContainerClassName);

            // Clear existing content
            contentContainer.selectAll(".tooltip-row").remove();

            let tooltipRow: D3.UpdateSelection = contentContainer.selectAll(rowsSelector).data(tooltipData);
            let newRow: D3.Selection = tooltipRow.enter().append("div").attr("class", ToolTipComponent.tooltipRowClassName);
            let newTitleCell: D3.Selection = newRow.append("div").attr("class", ToolTipComponent.tooltipTitleCellClassName);
            let newValueCell: D3.Selection = newRow.append("div").attr("class", ToolTipComponent.tooltipValueCellClassName);

            newTitleCell.text(function (d: TooltipDataItem) { return d.displayName; });
            newValueCell.text(function (d: TooltipDataItem) { return d.value; });
        }

        private getTooltipPosition(clickedArea: TouchUtils.Rectangle, clickedScreenArea: ScreenArea): TouchUtils.Point {
            let tooltipContainerBounds: ClientRect = this.tooltipContainer.node().getBoundingClientRect();
            let centerPointOffset: number = Math.floor(clickedArea.width / 2);
            let offsetX: number = 0;
            let offsetY: number = 0;
            let centerPoint: TouchUtils.Point = new TouchUtils.Point(clickedArea.x + centerPointOffset, clickedArea.y + centerPointOffset);
            let arrowOffset: number = 7;

            if (clickedScreenArea === ScreenArea.TopLeft) {
                offsetX += 3 * arrowOffset + centerPointOffset;
                offsetY -= 2 * arrowOffset + centerPointOffset;
            }
            else if (clickedScreenArea === ScreenArea.TopRight) {
                offsetX -= (2 * arrowOffset + tooltipContainerBounds.width + centerPointOffset);
                offsetY -= 2 * arrowOffset + centerPointOffset;
            }
            else if (clickedScreenArea === ScreenArea.BottomLeft) {
                offsetX += 3 * arrowOffset + centerPointOffset;
                offsetY -= (tooltipContainerBounds.height - 2 * arrowOffset + centerPointOffset);
            }
            else if (clickedScreenArea === ScreenArea.BottomRight) {
                offsetX -= (2 * arrowOffset + tooltipContainerBounds.width + centerPointOffset);
                offsetY -= (tooltipContainerBounds.height - 2 * arrowOffset + centerPointOffset);
            }

            centerPoint.offset(offsetX, offsetY);

            return centerPoint;
        }

        private setPosition(clickedArea: TouchUtils.Rectangle): void {
            let clickedScreenArea: ScreenArea = this.getClickedScreenArea(clickedArea);
            let tooltipPosition: TouchUtils.Point = this.getTooltipPosition(clickedArea, clickedScreenArea);

            this.tooltipContainer.style({ "left": tooltipPosition.x + "px", "top": tooltipPosition.y + "px" });
            this.setArrowPosition(clickedArea, clickedScreenArea);
        }

        private setArrowPosition(clickedArea: TouchUtils.Rectangle, clickedScreenArea: ScreenArea): void {
            let arrow: D3.Selection = this.getArrowElement();
            let arrowClassName: string;

            if (clickedScreenArea === ScreenArea.TopLeft) {
                arrowClassName = "top left";
            }
            else if (clickedScreenArea === ScreenArea.TopRight) {
                arrowClassName = "top right";
            }
            else if (clickedScreenArea === ScreenArea.BottomLeft) {
                arrowClassName = "bottom left";
            }
            else if (clickedScreenArea === ScreenArea.BottomRight) {
                arrowClassName = "bottom right";
            }

            arrow
                .attr('class', 'arrow') // Reset all classes
                .classed(arrowClassName, true);
        }

        private getArrowElement(): D3.Selection {
            return this.tooltipContainer.select("." + ToolTipComponent.arrowClassName);
        }

        private getClickedScreenArea(clickedArea: TouchUtils.Rectangle): ScreenArea {
            let screenWidth: number = this.customScreenWidth || window.innerWidth;
            let screenHeight: number = this.customScreenHeight || window.innerHeight;
            let centerPointOffset: number = clickedArea.width / 2;
            let centerPoint: TouchUtils.Point = new TouchUtils.Point(clickedArea.x + centerPointOffset, clickedArea.y + centerPointOffset);
            let halfWidth: number = screenWidth / 2;
            let halfHeight: number = screenHeight / 2;

            if (centerPoint.x < halfWidth && centerPoint.y < halfHeight) {
                return ScreenArea.TopLeft;
            }
            else if (centerPoint.x >= halfWidth && centerPoint.y < halfHeight) {
                return ScreenArea.TopRight;
            }
            else if (centerPoint.x < halfWidth && centerPoint.y >= halfHeight) {
                return ScreenArea.BottomLeft;
            }
            else if (centerPoint.x >= halfWidth && centerPoint.y >= halfHeight) {
                return ScreenArea.BottomRight;
            }
        }
    }

    export module TooltipManager {

        export var ShowTooltips: boolean = true;
        export var ToolTipInstance: ToolTipComponent = new ToolTipComponent();
        var GlobalTooltipEventsAttached: boolean = false;
        var tooltipMouseOverDelay: number = 500;
        var tooltipTouchDelay: number = 500;
        var tooltipTimeoutId: number;
        var handleTouchDelay: number = 1000;
        var handleTouchTimeoutId: number = 0;
        var mouseCoordinates: number[];

        export function addTooltip(d3Selection: D3.Selection, getTooltipInfoDelegate: (tooltipEvent: TooltipEvent) => TooltipDataItem[], reloadTooltipDataOnMouseMove?: boolean): void {

            if (!ShowTooltips) {
                return;
            }

            debug.assertValue(d3Selection, "d3Selection");

            let rootNode = d3.select(ToolTipComponent.parentContainerSelector).node();
            let touchStartEventName: string = getTouchStartEventName();
            let touchEndEventName: string = getTouchEndEventName();
            let isPointerEvent: boolean = touchStartEventName === "pointerdown" || touchStartEventName === "MSPointerDown";

            // Mouse events
            d3Selection.on("mouseover", function (d, i) {
                // Ignore mouseover (and other mouse events) while handling touch events
                if (!handleTouchTimeoutId && canDisplayTooltip(d3.event)) {
                    mouseCoordinates = getCoordinates(rootNode, true);
                    let elementCoordinates: number[] = getCoordinates(this, true);
                    let tooltipEvent: TooltipEvent = {
                        data: d,
                        index: i,
                        coordinates: mouseCoordinates,
                        elementCoordinates: elementCoordinates,
                        context: this,
                        isTouchEvent: false
                    };
                    clearTooltipTimeout();
                    tooltipTimeoutId = showDelayedTooltip(tooltipEvent, getTooltipInfoDelegate, tooltipMouseOverDelay);
                }
            });

            d3Selection.on("mouseout", function (d, i) {
                if (!handleTouchTimeoutId) {
                    clearTooltipTimeout();
                    hideTooltipEventHandler();
                }
            });

            d3Selection.on("mousemove", function (d, i) {
                if (!handleTouchTimeoutId && canDisplayTooltip(d3.event)) {
                    mouseCoordinates = getCoordinates(rootNode, true);
                    let elementCoordinates: number[] = getCoordinates(this, true);
                    let tooltipEvent: TooltipEvent = {
                        data: d,
                        index: i,
                        coordinates: mouseCoordinates,
                        elementCoordinates: elementCoordinates,
                        context: this,
                        isTouchEvent: false
                    };
                    moveTooltipEventHandler(tooltipEvent, getTooltipInfoDelegate, reloadTooltipDataOnMouseMove);
                }
            });
            
            // Touch events
            if (!GlobalTooltipEventsAttached) {
                // Add root container hide tooltip event
                attachGlobalEvents(touchStartEventName);
                GlobalTooltipEventsAttached = true;
            }

            d3Selection.on(touchStartEventName, function (d, i) {
                mouseCoordinates = null;
                hideTooltipEventHandler();
                let coordinates: number[] = getCoordinates(rootNode, isPointerEvent);
                let elementCoordinates: number[] = getCoordinates(this, isPointerEvent);
                let tooltipEvent: TooltipEvent = {
                    data: d,
                    index: i,
                    coordinates: coordinates,
                    elementCoordinates: elementCoordinates,
                    context: this,
                    isTouchEvent: true
                };
                clearTooltipTimeout();
                tooltipTimeoutId = showDelayedTooltip(tooltipEvent, getTooltipInfoDelegate, tooltipTouchDelay);
            });

            d3Selection.on(touchEndEventName, function (d, i) {
                clearTooltipTimeout();
                if (handleTouchTimeoutId)
                    clearTimeout(handleTouchTimeoutId);

                // At the end of touch action, set a timeout that will let us ignore the incoming mouse events for a small amount of time
                handleTouchTimeoutId = setTimeout(() => {
                    handleTouchTimeoutId = 0;
                }, handleTouchDelay);
            });
        }

        export function showDelayedTooltip(tooltipEvent: TooltipEvent, getTooltipInfoDelegate: (tooltipEvent: TooltipEvent) => TooltipDataItem[], delayInMs: number) {
            return setTimeout(() => showTooltipEventHandler(tooltipEvent, getTooltipInfoDelegate), delayInMs);
        }

        export function setLocalizedStrings(localizationOptions: TooltipLocalizationOptions): void {
            ToolTipComponent.localizationOptions = localizationOptions;
        }

        function showTooltipEventHandler(tooltipEvent: TooltipEvent, getTooltipInfoDelegate: (tooltipEvent: TooltipEvent) => TooltipDataItem[]) {
            let tooltipInfo: TooltipDataItem[] = getTooltipInfoDelegate(tooltipEvent);
            if (tooltipInfo) {
                let coordinates: number[] = mouseCoordinates || tooltipEvent.coordinates;
                let clickedArea: TouchUtils.Rectangle = getClickedArea(coordinates[0], coordinates[1], tooltipEvent.isTouchEvent);
                ToolTipInstance.show(tooltipInfo, clickedArea);
            }
        }

        function moveTooltipEventHandler(tooltipEvent: TooltipEvent, getTooltipInfoDelegate: (tooltipEvent: TooltipEvent) => TooltipDataItem[], reloadTooltipDataOnMouseMove: boolean) {
            let tooltipInfo: TooltipDataItem[];
            if (reloadTooltipDataOnMouseMove) {
                tooltipInfo = getTooltipInfoDelegate(tooltipEvent);
            }
            let clickedArea: TouchUtils.Rectangle = getClickedArea(tooltipEvent.coordinates[0], tooltipEvent.coordinates[1], tooltipEvent.isTouchEvent);
            ToolTipInstance.move(tooltipInfo, clickedArea);
        };

        function hideTooltipEventHandler() {
            ToolTipInstance.hide();
        };

        function clearTooltipTimeout() {
            if (tooltipTimeoutId) {
                clearTimeout(tooltipTimeoutId);
            }
        }

        function canDisplayTooltip(d3Event: any): boolean {
            let cadDisplay: boolean = true;
            let mouseEvent: MouseEvent = <MouseEvent>d3Event;
            if (mouseEvent.buttons !== undefined) {
                // Check mouse buttons state
                let hasMouseButtonPressed = mouseEvent.buttons !== 0;
                cadDisplay = !hasMouseButtonPressed;
            }
            return cadDisplay;
        }

        function getTouchStartEventName(): string {
            let eventName: string = "touchstart";

            if (window["PointerEvent"]) {
                // IE11
                eventName = "pointerdown";
            } else if (window["MSPointerEvent"]) {
                // IE10
                eventName = "MSPointerDown";
            }

            return eventName;
        }

        function getTouchEndEventName(): string {
            let eventName: string = "touchend";

            if (window["PointerEvent"]) {
                // IE11
                eventName = "pointerup";
            } else if (window["MSPointerEvent"]) {
                // IE10
                eventName = "MSPointerUp";
            }

            return eventName;
        }

        function getCoordinates(rootNode: Element, isPointerEvent: boolean): number[] {
            let coordinates: number[];

            if (isPointerEvent) {
                coordinates = d3.mouse(rootNode);
            }
            else {
                let touchCoordinates = d3.touches(rootNode);
                if (touchCoordinates && touchCoordinates.length > 0) {
                    coordinates = touchCoordinates[0];
                }
            }

            return coordinates;
        }

        function attachGlobalEvents(touchStartEventName: string): void {
            d3.select(ToolTipComponent.parentContainerSelector).on(touchStartEventName, function (d, i) {
                ToolTipInstance.hide();
            });
        }

        function getClickedArea(x: number, y: number, isTouchEvent: boolean): TouchUtils.Rectangle {
            let width: number = 0;
            let pointX: number = x;
            let pointY: number = y;

            if (isTouchEvent) {
                width = 12;
                let offset: number = width / 2;
                pointX = Math.max(x - offset, 0);
                pointY = Math.max(y - offset, 0);
            }

            return new TouchUtils.Rectangle(pointX, pointY, width, width);
        }
    }

    export module TooltipBuilder {

        // TODO: implement options bag as input parameter
        export function createTooltipInfo(
            formatStringProp: DataViewObjectPropertyIdentifier,
            dataViewCat: DataViewCategorical,
            categoryValue: any,
            value?: any,
            categories?: DataViewCategoryColumn[],
            seriesData?: TooltipSeriesDataItem[],
            seriesIndex?: number,
            categoryIndex?: number,
            highlightedValue?: any): TooltipDataItem[] {
            let categorySource: TooltipCategoryDataItem;
            let seriesSource: TooltipSeriesDataItem[] = [];
            let valuesSource: DataViewMetadataColumn = undefined;
            seriesIndex = seriesIndex | 0;

            let categoriesData = dataViewCat ? dataViewCat.categories : categories;
            if (categoriesData && categoriesData.length > 0) {
                categorySource = { value: categoryValue, metadata: categoriesData[0].source };
            }
            if (dataViewCat && dataViewCat.values) {
                if (categorySource && categorySource.metadata === dataViewCat.values.source) {
                    // Category/series on the same column -- don't repeat its value in the tooltip.
                }
                else {
                    valuesSource = dataViewCat.values.source;
                }
                if (dataViewCat.values.length > 0) {
                    let valueColumn: DataViewValueColumn = dataViewCat.values[seriesIndex];
                    let isAutoGeneratedColumn: boolean = !!(valueColumn && valueColumn.source && (<DataViewMetadataAutoGeneratedColumn>valueColumn.source).isAutoGeneratedColumn);

                    if (!isAutoGeneratedColumn) {
                        seriesSource.push({ value: value, highlightedValue: highlightedValue, metadata: valueColumn });
                    }
                }

                // check for gradient measure index 
                let gradientMeasureIndex: number = GradientUtils.getGradientMeasureIndex(dataViewCat);
                let gradientValueColumn: DataViewValueColumn = gradientMeasureIndex === - 1 ? null : dataViewCat.values[gradientMeasureIndex];
                //If the same column has both Y and Gradient roles then make sure we don't add it more than once
                if (gradientValueColumn && seriesIndex !== gradientMeasureIndex) {
                    // Saturation color
                    seriesSource.push({ value: gradientValueColumn.values[categoryIndex], metadata: { source: gradientValueColumn.source, values: [] } });
                }
            }
            if (seriesData) {
                for (let i: number = 0, len: number = seriesData.length; i < len; i++) {
                    let singleSeriesData: TooltipSeriesDataItem = seriesData[i];
                    if (categorySource && categorySource.metadata === singleSeriesData.metadata.source)
                        continue;

                    seriesSource.push({ value: singleSeriesData.value, metadata: singleSeriesData.metadata });
                }
            }

            let tooltipInfo: TooltipDataItem[] = createTooltipData(formatStringProp, categorySource, valuesSource, seriesSource);

            return tooltipInfo;
        }

        function createTooltipData(
            formatStringProp: DataViewObjectPropertyIdentifier,
            categoryValue: TooltipCategoryDataItem,
            valuesSource: DataViewMetadataColumn,
            seriesValues: TooltipSeriesDataItem[]): TooltipDataItem[] {

            debug.assertValue(seriesValues, "seriesSource");
            debug.assertValue(ToolTipComponent.localizationOptions, "ToolTipComponent.localizationOptions");
            debug.assertAnyValue(formatStringProp, 'formatStringProp');

            let items: TooltipDataItem[] = [];

            if (categoryValue) {
                let categoryFormattedValue: string = getFormattedValue(categoryValue.metadata, formatStringProp, categoryValue.value);
                items.push({ displayName: categoryValue.metadata.displayName, value: categoryFormattedValue });
            }

            if (valuesSource) {
                // Dynamic series value
                let dynamicValue: string;
                if (seriesValues.length > 0) {
                    let dynamicValueMetadata: DataViewMetadataColumn = seriesValues[0].metadata.source;
                    dynamicValue = getFormattedValue(valuesSource, formatStringProp, dynamicValueMetadata.groupName);
                }
                items.push({ displayName: valuesSource.displayName, value: dynamicValue });
            }

            for (let i = 0; i < seriesValues.length; i++) {
                let seriesData = seriesValues[i];

                if (seriesData && seriesData.metadata) {
                    let seriesMetadataColumn = seriesData.metadata.source;
                    let value = seriesData.value;
                    let highlightedValue = seriesData.highlightedValue;

                    if (value || value === 0) {
                        let formattedValue: string = getFormattedValue(seriesMetadataColumn, formatStringProp, value);
                        items.push({ displayName: seriesMetadataColumn.displayName, value: formattedValue });
                    }

                    if (highlightedValue || highlightedValue === 0) {
                        let formattedHighlightedValue: string = getFormattedValue(seriesMetadataColumn, formatStringProp, highlightedValue);
                        let displayName = ToolTipComponent.localizationOptions.highlightedValueDisplayName;
                        items.push({ displayName: displayName, value: formattedHighlightedValue });
                    }
                }
            }

            return items;
        }

        function getFormattedValue(column: DataViewMetadataColumn, formatStringProp: DataViewObjectPropertyIdentifier, value: any) {
            let formatString: string = getFormatStringFromColumn(column, formatStringProp);
            return valueFormatter.format(value, formatString);
        }

        function getFormatStringFromColumn(column: DataViewMetadataColumn, formatStringProp: DataViewObjectPropertyIdentifier): string {
            if (column) {
                let formatString: string = valueFormatter.getFormatString(column, formatStringProp, true);
                return formatString || column.format;
            }
            return null;
        }
    }
}