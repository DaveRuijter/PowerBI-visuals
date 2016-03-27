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
    import PixelConverter = jsCommon.PixelConverter;

    export interface SlicerConstructorOptions {
        behavior?: SlicerWebBehavior;
    }

    export interface SlicerData {
        categorySourceName: string;
        formatString: string;
        slicerDataPoints: SlicerDataPoint[];
        slicerSettings: SlicerSettings;
        hasSelectionOverride?: boolean;
    }

    export interface SlicerDataPoint extends SelectableDataPoint {
        value: string;
        isSelectAllDataPoint?: boolean;
    }

    export interface SlicerSettings {
        general: {
            outlineColor: string;
            outlineWeight: number;
        };
        header: {
            borderBottomWidth: number;
            show: boolean;
            outline: string;
            fontColor: string;
            background: string;
            textSize: number;
        };
        slicerText: {
            color: string;
            outline: string;
            background: string;
            textSize: number;
        };
    }

    interface CheckboxStyle {
        transform: string;
        'transform-origin': string;
        'font-size': string;
    }

    export class Slicer implements IVisual {
        private element: JQuery;
        private currentViewport: IViewport;
        private dataView: DataView;
        private slicerHeader: D3.Selection;
        private slicerBody: D3.Selection;
        private listView: IListView;
        private slicerData: SlicerData;
        private settings: SlicerSettings;
        private interactivityService: IInteractivityService;
        private behavior: SlicerWebBehavior;
        private hostServices: IVisualHostServices;
        private static clearTextKey = 'Slicer_Clear';
        private static selectAllTextKey = 'Slicer_SelectAll';
        private waitingForData: boolean;
        private textProperties: TextProperties = {
            'fontFamily': 'wf_segoe-ui_normal, helvetica, arial, sans-serif',
            'fontSize': '14px',
        };

        // Based on sprite size defined in CSS sprite sheet for slicer-checkbox
        private static CheckboxSpritePixelSizeMinimum: number = 8;
        private static CheckboxSpritePixelSize: number = 13;
        private static CheckboxSpritePixelSizeRange: number = Slicer.CheckboxSpritePixelSize - Slicer.CheckboxSpritePixelSizeMinimum;

        private static Container: ClassAndSelector = {
            class: 'slicerContainer',
            selector: '.slicerContainer'
        };
        private static Header: ClassAndSelector = {
            class: 'slicerHeader',
            selector: '.slicerHeader'
        };
        private static HeaderText: ClassAndSelector = {
            class: 'headerText',
            selector: '.headerText'
        };
        private static Body: ClassAndSelector = {
            class: 'slicerBody',
            selector: '.slicerBody'
        };
        private static ItemContainer: ClassAndSelector = {
            class: 'slicerItemContainer',
            selector: '.slicerItemContainer'
        };
        private static LabelText: ClassAndSelector = {
            class: 'slicerText',
            selector: '.slicerText'
        };
        private static Input: ClassAndSelector = {
            class: 'slicerCheckbox',
            selector: '.slicerCheckbox'
        };
        private static Checkbox: ClassAndSelector = {
            class: 'checkbox',
            selector: '.checkbox'
        };
        private static Clear: ClassAndSelector = {
            class: 'clear',
            selector: '.clear'
        };

        public static DefaultStyleProperties(): SlicerSettings {
            return {
                general: {
                    outlineColor: '#808080',
                    outlineWeight: 1
                },
                header: {
                    borderBottomWidth: 1,
                    show: true,
                    outline: 'BottomOnly',
                    fontColor: '#000000',
                    background: '#ffffff',
                    textSize: 10,
                },
                slicerText: {
                    color: '#666666',
                    outline: 'None',
                    background: '#ffffff',
                    textSize: 10,
                },
            };
        }

        constructor(options?: SlicerConstructorOptions) {
            if (options) {
                this.behavior = options.behavior;
            }
        }

        public static converter(dataView: DataView, localizedSelectAllText: string, interactivityService: IInteractivityService): SlicerData {
            let slicerData: SlicerData;
            if (!dataView) {
                return;
            }

            let dataViewCategorical = dataView.categorical;
            if (dataViewCategorical == null || dataViewCategorical.categories == null || dataViewCategorical.categories.length === 0)
                return;

            let isInvertedSelectionMode = undefined;
            let objects = dataView.metadata ? <any> dataView.metadata.objects : undefined;
            let categories = dataViewCategorical.categories[0];

            let numberOfScopeIds: number;
            if (objects && objects.general && objects.general.filter) {
                let identityFields = categories.identityFields;
                if (!identityFields)
                    return;
                let filter = <powerbi.data.SemanticFilter>objects.general.filter;
                let scopeIds = powerbi.data.SQExprConverter.asScopeIdsContainer(filter, identityFields);
                if (scopeIds) {
                    isInvertedSelectionMode = scopeIds.isNot;
                    numberOfScopeIds = scopeIds.scopeIds ? scopeIds.scopeIds.length : 0;
                }
                else {
                    isInvertedSelectionMode = false;
                }
            }

            if (interactivityService) {
                if (isInvertedSelectionMode === undefined) {
                    // The selection state is read from the Interactivity service in case of SelectAll or Clear when query doesn't update the visual
                    isInvertedSelectionMode = interactivityService.isSelectionModeInverted();
                }
                else {
                    interactivityService.setSelectionModeInverted(isInvertedSelectionMode);
                }
            }

            let categoryValuesLen = categories && categories.values ? categories.values.length : 0;
            let slicerDataPoints: SlicerDataPoint[] = [];

            slicerDataPoints.push({
                value: localizedSelectAllText,
                identity: SelectionId.createWithMeasure(localizedSelectAllText),
                selected: !!isInvertedSelectionMode,
                isSelectAllDataPoint: true
            });                    
                                     
            // Pass over the values to see if there's a positive or negative selection
            let hasSelection: boolean = undefined;

            for (let idx = 0; idx < categoryValuesLen; idx++) {
                let selected = isCategoryColumnSelected(slicerProps.selectedPropertyIdentifier, categories, idx);
                if (selected != null) {
                    hasSelection = selected;
                    break;
                }
            }

            let numberOfCategoriesSelectedInData = 0;
            for (let idx = 0; idx < categoryValuesLen; idx++) {
                let categoryIdentity = categories.identity ? categories.identity[idx] : null;
                let categoryIsSelected = isCategoryColumnSelected(slicerProps.selectedPropertyIdentifier, categories, idx);

                if (hasSelection != null) {
                    // If the visual is in InvertedSelectionMode, all the categories should be selected by default unless they are not selected
                    // If the visual is not in InvertedSelectionMode, we set all the categories to be false except the selected category                         
                    if (isInvertedSelectionMode) {
                        if (categories.objects == null)
                            categoryIsSelected = undefined;

                        if (categoryIsSelected != null) {
                            categoryIsSelected = hasSelection;
                        }
                        else if (categoryIsSelected == null)
                            categoryIsSelected = !hasSelection;
                    }
                    else {
                        if (categoryIsSelected == null) {
                            categoryIsSelected = !hasSelection;
                        }
                    }
                }

                if (categoryIsSelected)
                    numberOfCategoriesSelectedInData++;

                slicerDataPoints.push({
                    value: categories.values[idx],
                    identity: SelectionId.createWithId(categoryIdentity),
                    selected: categoryIsSelected
                });
            }

            let defaultSettings = this.DefaultStyleProperties();
            objects = dataView.metadata.objects;
            if (objects) {
                defaultSettings.general.outlineColor = DataViewObjects.getFillColor(objects, slicerProps.general.outlineColor, defaultSettings.general.outlineColor);
                defaultSettings.general.outlineWeight = DataViewObjects.getValue<number>(objects, slicerProps.general.outlineWeight, defaultSettings.general.outlineWeight);

                defaultSettings.header.show = DataViewObjects.getValue<boolean>(objects, slicerProps.header.show, defaultSettings.header.show);
                defaultSettings.header.fontColor = DataViewObjects.getFillColor(objects, slicerProps.header.fontColor, defaultSettings.header.fontColor);
                defaultSettings.header.background = DataViewObjects.getFillColor(objects, slicerProps.header.background, defaultSettings.header.background);
                defaultSettings.header.outline = DataViewObjects.getValue<string>(objects, slicerProps.header.outline, defaultSettings.header.outline);
                defaultSettings.header.textSize = DataViewObjects.getValue<number>(objects, slicerProps.header.textSize, defaultSettings.header.textSize);

                defaultSettings.slicerText.color = DataViewObjects.getFillColor(objects, slicerProps.Rows.fontColor, defaultSettings.slicerText.color);
                defaultSettings.slicerText.background = DataViewObjects.getFillColor(objects, slicerProps.Rows.background, defaultSettings.slicerText.background);
                defaultSettings.slicerText.outline = DataViewObjects.getValue<string>(objects, slicerProps.Rows.outline, defaultSettings.slicerText.outline);
                defaultSettings.slicerText.textSize = DataViewObjects.getValue<number>(objects, slicerProps.Rows.textSize, defaultSettings.slicerText.textSize);
            }

            slicerData = {
                categorySourceName: categories.source.displayName,
                formatString: valueFormatter.getFormatString(categories.source, slicerProps.formatString),
                slicerSettings: defaultSettings,
                slicerDataPoints: slicerDataPoints,
            };

            // Override hasSelection if a objects contained more scopeIds than selections we found in the data
            if (numberOfScopeIds != null && numberOfScopeIds > numberOfCategoriesSelectedInData) {
                slicerData.hasSelectionOverride = true;
            }

            return slicerData;
        }

        public init(options: VisualInitOptions): void {
            this.element = options.element;
            this.currentViewport = options.viewport;
            if (this.behavior) {
                this.interactivityService = createInteractivityService(options.host);
            }
            this.hostServices = options.host;
            this.settings = Slicer.DefaultStyleProperties();

            this.initContainer();
        }

        public onDataChanged(options: VisualDataChangedOptions): void {
            let dataViews = options.dataViews;
            debug.assertValue(dataViews, 'dataViews');

            let existingDataView = this.dataView;
            if (dataViews && dataViews.length > 0) {
                this.dataView = dataViews[0];
            }

            let resetScrollbarPosition = false;
            // Null check is needed here. If we don't check for null, selecting a value on loadMore event will evaluate the below condition to true and resets the scrollbar
            if (options.operationKind !== undefined) {
                resetScrollbarPosition = options.operationKind !== VisualDataChangeOperationKind.Append
                && !DataViewAnalysis.hasSameCategoryIdentity(existingDataView, this.dataView);
            }

            this.updateInternal(resetScrollbarPosition);
            this.waitingForData = false;
        }

        public onResizing(finalViewport: IViewport): void {
            this.currentViewport = finalViewport;
            this.updateInternal(false /* resetScrollbarPosition */);
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] {
            let data = this.slicerData;
            if (!data)
                return;

            let objectName = options.objectName;
            switch (objectName) {
                case 'Rows':
                    return this.enumerateRows(data);
                case 'header':
                    return this.enumerateHeader(data);
                case 'general':
                    return this.enumerateGeneral(data);
            }
        }

        private enumerateHeader(data: SlicerData): VisualObjectInstance[] {
            let slicerSettings = this.settings;
            let fontColor = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.header && data.slicerSettings.header.fontColor ?
                data.slicerSettings.header.fontColor : slicerSettings.header.fontColor;
            let background = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.header && data.slicerSettings.header.background ?
                data.slicerSettings.header.background : slicerSettings.header.background;
            return [{
                selector: null,
                objectName: 'header',
                properties: {
                    show: slicerSettings.header.show,
                    fontColor: fontColor,
                    background: background,
                    outline: slicerSettings.header.outline,
                    textSize: slicerSettings.header.textSize,
                }
            }];
        }

        private enumerateRows(data: SlicerData): VisualObjectInstance[] {
            let slicerSettings = this.settings;
            let fontColor = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.slicerText && data.slicerSettings.slicerText.color ?
                data.slicerSettings.slicerText.color : slicerSettings.slicerText.color;
            let background = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.slicerText && data.slicerSettings.slicerText.background ?
                data.slicerSettings.slicerText.background : slicerSettings.slicerText.background;
            return [{
                selector: null,
                objectName: 'rows',
                properties: {
                    fontColor: fontColor,
                    background: background,
                    outline: slicerSettings.slicerText.outline,
                    textSize: slicerSettings.slicerText.textSize,
                }
            }];
        }

        private enumerateGeneral(data: SlicerData): VisualObjectInstance[] {
            let slicerSettings = this.settings;
            let outlineColor = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.general && data.slicerSettings.general.outlineColor ?
                data.slicerSettings.general.outlineColor : slicerSettings.general.outlineColor;
            let outlineWeight = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.general && data.slicerSettings.general.outlineWeight ?
                data.slicerSettings.general.outlineWeight : slicerSettings.general.outlineWeight;

            return [{
                selector: null,
                objectName: 'general',
                properties: {
                    outlineColor: outlineColor,
                    outlineWeight: outlineWeight
                }
            }];
        }

        private updateInternal(resetScrollbarPosition: boolean) {
            this.updateSlicerBodyDimensions();

            let localizedSelectAllText = this.hostServices.getLocalizedString(Slicer.selectAllTextKey);
            let data = Slicer.converter(this.dataView, localizedSelectAllText, this.interactivityService);
            if (!data) {
                this.listView.empty();
                return;
            }

            data.slicerSettings.general.outlineWeight = data.slicerSettings.general.outlineWeight < 0 ? 0 : data.slicerSettings.general.outlineWeight;
            this.slicerData = data;
            this.settings = this.slicerData.slicerSettings;

            this.listView
                .viewport(this.getSlicerBodyViewport(this.currentViewport))
                .rowHeight(this.getRowHeight())
                .data(
                    data.slicerDataPoints,
                    (d: SlicerDataPoint) => $.inArray(d, data.slicerDataPoints),
                    resetScrollbarPosition
                    );
        }

        private initContainer() {
            let settings = this.settings;
            let slicerBodyViewport = this.getSlicerBodyViewport(this.currentViewport);

            let slicerContainerDiv = document.createElement('div');
            slicerContainerDiv.className = Slicer.Container.class;
            let slicerContainer: D3.Selection = d3.select(slicerContainerDiv);

            this.slicerHeader = slicerContainer.append('div').classed(Slicer.Header.class, true);

            this.slicerHeader.append('span')
                .classed(Slicer.Clear.class, true)
                .attr('title', this.hostServices.getLocalizedString(Slicer.clearTextKey));

            this.slicerHeader.append('div').classed(Slicer.HeaderText.class, true)
                .style({
                    'border-style': this.getBorderStyle(settings.header.outline),
                    'border-color': settings.general.outlineColor,
                    'border-width': this.getBorderWidth(settings.header.outline, settings.general.outlineWeight),
                    'font-size': PixelConverter.fromPoint(settings.header.textSize),
                });

            this.slicerBody = slicerContainer.append('div').classed(Slicer.Body.class, true)
                .style({
                    'height': PixelConverter.toString(slicerBodyViewport.height),
                    'width': PixelConverter.toString(slicerBodyViewport.width),
                });

            let rowEnter = (rowSelection: D3.Selection) => {
                let settings = this.settings;
                let listItemElement = rowSelection.append('li')
                    .classed(Slicer.ItemContainer.class, true);

                let labelElement = listItemElement.append('div')
                    .classed(Slicer.Input.class, true);

                labelElement.append('input')
                    .attr('type', 'checkbox');

                labelElement.append('span')
                    .classed(Slicer.Checkbox.class, true)
                    .style(this.buildCheckboxStyle());

                listItemElement.append('span')
                    .classed(Slicer.LabelText.class, true)
                    .style({
                        'font-size': PixelConverter.fromPoint(settings.slicerText.textSize),
                    });
            };

            let rowUpdate = (rowSelection: D3.Selection) => {
                let settings = this.settings;
                let data = this.slicerData;
                if (data && settings) {

                    if (settings.header.show) {
                        this.slicerHeader.style('display', 'block');
                        this.slicerHeader.select(Slicer.HeaderText.selector)
                            .text(this.slicerData.categorySourceName)
                            .style({
                                'border-style': this.getBorderStyle(settings.header.outline),
                                'border-color': settings.general.outlineColor,
                                'border-width': this.getBorderWidth(settings.header.outline, settings.general.outlineWeight),
                                'color': settings.header.fontColor,
                                'background-color': settings.header.background,
                                'font-size': PixelConverter.fromPoint(settings.header.textSize),
                            });
                    }
                    else {
                        this.slicerHeader.style('display', 'none');
                    }

                    let slicerText = rowSelection.selectAll(Slicer.LabelText.selector);
                    let formatString = data.formatString;
                    slicerText.text((d: SlicerDataPoint) => valueFormatter.format(d.value, formatString));
                    slicerText.style({
                        'color': settings.slicerText.color,
                        'background-color': settings.slicerText.background,
                        'border-style': this.getBorderStyle(settings.slicerText.outline),
                        'border-color': settings.general.outlineColor,
                        'border-width': this.getBorderWidth(settings.slicerText.outline, settings.general.outlineWeight),
                        'font-size': PixelConverter.fromPoint(settings.slicerText.textSize),
                    });

                    let slicerCheckbox = rowSelection.selectAll(Slicer.Input.selector).selectAll('span');
                    slicerCheckbox.style(this.buildCheckboxStyle());

                    if (this.interactivityService && this.slicerBody) {
                        let slicerBody = this.slicerBody.attr('width', this.currentViewport.width);
                        let slicerItemContainers = slicerBody.selectAll(Slicer.ItemContainer.selector);
                        let slicerItemLabels = slicerBody.selectAll(Slicer.LabelText.selector);
                        let slicerItemInputs = slicerBody.selectAll(Slicer.Input.selector);
                        let slicerClear = this.slicerHeader.select(Slicer.Clear.selector);

                        let behaviorOptions: SlicerBehaviorOptions = {
                            dataPoints: data.slicerDataPoints,
                            slicerItemContainers: slicerItemContainers,
                            slicerItemLabels: slicerItemLabels,
                            slicerItemInputs: slicerItemInputs,
                            slicerClear: slicerClear,
                            interactivityService: this.interactivityService,
                            slicerSettings: data.slicerSettings,
                        };

                        this.interactivityService.bind(data.slicerDataPoints, this.behavior, behaviorOptions, { overrideSelectionFromData: true, hasSelectionOverride: data.hasSelectionOverride });
                        SlicerWebBehavior.styleSlicerInputs(rowSelection.select(Slicer.Input.selector), this.interactivityService.hasSelection());
                    }
                    else {
                        SlicerWebBehavior.styleSlicerInputs(rowSelection.select(Slicer.Input.selector), false);
                    }
                }
            };

            let rowExit = (rowSelection: D3.Selection) => {
                rowSelection.remove();
            };

            let listViewOptions: ListViewOptions = {
                rowHeight: this.getRowHeight(),
                enter: rowEnter,
                exit: rowExit,
                update: rowUpdate,
                loadMoreData: () => this.onLoadMoreData(),
                scrollEnabled: true,
                viewport: this.getSlicerBodyViewport(this.currentViewport),
                baseContainer: this.slicerBody,
            };

            this.listView = ListViewFactory.createListView(listViewOptions);

            // Append container to DOM
            this.element.get(0).appendChild(slicerContainerDiv);
        }

        private onLoadMoreData(): void {
            if (!this.waitingForData && this.dataView.metadata && this.dataView.metadata.segment) {
                this.hostServices.loadMoreData();
                this.waitingForData = true;
            }
        }

        private getSlicerBodyViewport(currentViewport: IViewport): IViewport {
            let settings = this.settings;
            let headerHeight = (settings.header.show) ? this.getHeaderHeight() : 0;
            let slicerBodyHeight = currentViewport.height - (headerHeight + settings.header.borderBottomWidth);
            return {
                height: slicerBodyHeight,
                width: currentViewport.width
            };
        }

        private updateSlicerBodyDimensions(): void {
            let slicerViewport = this.getSlicerBodyViewport(this.currentViewport);
            this.slicerBody.style({
                'height': PixelConverter.toString(slicerViewport.height),
                'width': PixelConverter.toString(slicerViewport.width),
            });
        }

        private getTextProperties(textSize: number): TextProperties {
            this.textProperties.fontSize = PixelConverter.fromPoint(textSize);
            return this.textProperties;
        }

        private getHeaderHeight(): number {
            return TextMeasurementService.estimateSvgTextHeight(
                this.getTextProperties(this.settings.header.textSize)
                );
        }

        private getRowHeight(): number {
            return TextMeasurementService.estimateSvgTextHeight(
                this.getTextProperties(this.settings.slicerText.textSize)
                );
        }

        private getCheckboxScale(): string {
            let scale = jsCommon.TextSizeDefaults.getScale(this.settings.slicerText.textSize);
            let size = (Slicer.CheckboxSpritePixelSizeMinimum + (Slicer.CheckboxSpritePixelSizeRange * scale));
            let relativeZoom = size / Slicer.CheckboxSpritePixelSize;

            let rounded = powerbi.Double.toIncrement(relativeZoom, 0.05);

            return SVGUtil.scale(rounded);
        }

        private buildCheckboxStyle(): CheckboxStyle {
            let checkboxScale = this.getCheckboxScale();
            let checkboxTransformOrigin = SVGUtil.transformOrigin('left', 'center');
            return {
                'transform': checkboxScale,
                'transform-origin': checkboxTransformOrigin,
                'font-size': PixelConverter.fromPoint(this.settings.slicerText.textSize),
            };
        }

        private getBorderStyle(outlineElement: string): string {
            return outlineElement === '0px' ? 'none' : 'solid';
        }

        private getBorderWidth(outlineElement: string, outlineWeight: number): string {
            switch (outlineElement) {
                case 'None':
                    return '0px';
                case 'BottomOnly':
                    return '0px 0px ' + outlineWeight + 'px 0px';
                case 'TopOnly':
                    return outlineWeight + 'px 0px 0px 0px';
                case 'TopBottom':
                    return outlineWeight + 'px 0px ' + outlineWeight + 'px 0px';
                case 'LeftRight':
                    return '0px ' + outlineWeight + 'px 0px ' + outlineWeight + 'px';
                case 'Frame':
                    return outlineWeight + 'px';
                default:
                    return outlineElement.replace("1", outlineWeight.toString());

            }
        }

    }
}