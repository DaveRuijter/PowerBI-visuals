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
    export interface BasicShapeDataViewObjects extends DataViewObjects {
        general: BasicShapeDataViewObject;
        line: LineObject;
        fill: FillObject;
        lockAspect: LockAspectObject;
        rotation: RotationObject;
    }

    export interface LineObject extends DataViewObject {
        lineColor: Fill;
        roundEdge: number;
        weight: number;
        transparency: number;
    }

    export interface FillObject extends DataViewObject {
        transparency: number;
        fillColor: Fill;
        show: boolean;
    }

    export interface LockAspectObject extends DataViewObject {
        show: boolean;
    }

    export interface RotationObject extends DataViewObject {
        angle: number;
    }

    export interface BasicShapeDataViewObject extends DataViewObject {
        shapeType: string;
        shapeSvg: string;
    }

    export interface BasicShapeData {
        shapeType: string;
        lineColor: string;
        lineTransparency: number;
        lineWeight: number;
        showFill: boolean;
        fillColor: string;
        shapeTransparency: number;
        lockAspectRatio: boolean;
        roundEdge: number;
        angle: number;
    }

    export class BasicShapeVisual implements IVisual {
        private currentViewport: IViewport;
        private element: JQuery;
        private data: BasicShapeData;
        private selection: D3.Selection;

        public static DefaultShape: string = powerbi.basicShapeType.rectangle;
        public static DefaultStrokeColor: string = '#00B8AA';
        public static DefaultFillColor: string = '#E6E6E6';
        public static DefaultFillTransValue: number = 100;
        public static DefaultWeightValue: number = 3;
        public static DefaultLineTransValue: number = 100;
        public static DefaultRoundEdgeValue: number = 0;
        public static DefaultAngle: number = 0;

        /**property for the shape line color */
        get shapeType(): string {
            return this.data ? this.data.shapeType : BasicShapeVisual.DefaultShape;
        }
        set shapeType(shapeType: string) {
            this.data.shapeType = shapeType;
        }

        /**property for the shape line color */
        get lineColor(): string {
            return this.data ? this.data.lineColor : BasicShapeVisual.DefaultStrokeColor;
        }
        set lineColor(color: string) {
            this.data.lineColor = color;
        }
        /**property for the shape line transparency */
        get lineTransparency(): number {
            return this.data ? this.data.lineTransparency : BasicShapeVisual.DefaultLineTransValue;
        }
        set lineTransparency(trans: number) {
            this.data.lineTransparency = trans;
        }
        /**property for the shape line weight */
        get lineWeight(): number {
            return this.data ? this.data.lineWeight : BasicShapeVisual.DefaultWeightValue;
        }
        set lineWeight(weight: number) {
            this.data.lineWeight = weight;
        }
        /**property for the shape round edge */
        get roundEdge(): number {
            return this.data ? this.data.roundEdge : BasicShapeVisual.DefaultRoundEdgeValue;
        }
        set roundEdge(roundEdge: number) {
            this.data.roundEdge = roundEdge;
        }
        /**property for showing the fill properties */
        get showFill(): boolean {
            return this.data ? this.data.showFill : true;
        }
        set showFill(show: boolean) {
            this.data.showFill = show;
        }
        /**property for the shape line color */
        get fillColor(): string {
            return this.data ? this.data.fillColor : BasicShapeVisual.DefaultFillColor;
        }
        set fillColor(color: string) {
            this.data.fillColor = color;
        }
        /**property for the shape fill transparency */
        get shapeTransparency(): number {
            return this.data ? this.data.shapeTransparency : BasicShapeVisual.DefaultFillTransValue;
        }
        set shapeTransparency(trans: number) {
            this.data.shapeTransparency = trans;
        }
        /**property for showing the lock aspect ratio */
        get lockAspectRatio(): boolean {
            return this.data ? this.data.lockAspectRatio : false;
        }
        set lockAspectRatio(show: boolean) {
            this.data.lockAspectRatio = show;
        }
        /**property for the shape angle */
        get angle(): number {
            return this.data ? this.data.angle : BasicShapeVisual.DefaultAngle;
        }
        set angle(angle: number) {
            this.data.angle = angle;
        }

        public init(options: VisualInitOptions) {
            this.element = options.element;
            this.selection = d3.select(this.element.context);
            this.currentViewport = options.viewport;
        }

        public constructor(options?: VisualInitOptions) {
        }

        public update(options: VisualUpdateOptions): void {
            debug.assertValue(options, 'options');

            this.currentViewport = options.viewport;

            let dataViews = options.dataViews;
            if (!_.isEmpty(dataViews)) {
                let dataView = options.dataViews[0];
                if (dataView.metadata && dataView.metadata.objects) {
                    let dataViewObject = <BasicShapeDataViewObjects>options.dataViews[0].metadata.objects;
                    this.data = this.getDataFromDataView(dataViewObject);
                }
            }

            this.render();
        }

        private getDataFromDataView(dataViewObject: BasicShapeDataViewObjects): BasicShapeData {
            if (dataViewObject) {
                return {
                    shapeType: dataViewObject.general !== undefined && dataViewObject.general.shapeType !== undefined ? dataViewObject.general.shapeType : this.shapeType,
                    lineColor: dataViewObject.line !== undefined && dataViewObject.line.lineColor !== undefined ? dataViewObject.line.lineColor.solid.color : this.lineColor,
                    lineTransparency: dataViewObject.line !== undefined && dataViewObject.line.transparency !== undefined ? dataViewObject.line.transparency : this.lineTransparency,
                    lineWeight: dataViewObject.line !== undefined && dataViewObject.line.weight !== undefined ? dataViewObject.line.weight : this.lineWeight,
                    roundEdge: dataViewObject.line !== undefined && dataViewObject.line.roundEdge !== undefined ? dataViewObject.line.roundEdge : this.roundEdge,
                    shapeTransparency: dataViewObject.fill !== undefined && dataViewObject.fill.transparency !== undefined ? dataViewObject.fill.transparency : this.shapeTransparency,
                    fillColor: dataViewObject.fill !== undefined && dataViewObject.fill.fillColor !== undefined ? dataViewObject.fill.fillColor.solid.color : this.fillColor,
                    showFill: dataViewObject.fill !== undefined && dataViewObject.fill.show !== undefined ? dataViewObject.fill.show : this.showFill,
                    lockAspectRatio: dataViewObject.lockAspect !== undefined && dataViewObject.lockAspect.show !== undefined ? dataViewObject.lockAspect.show : this.lockAspectRatio,
                    angle: dataViewObject.rotation !== undefined && dataViewObject.rotation.angle !== undefined ? dataViewObject.rotation.angle : this.angle
                };
            }

            return null;
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] {
            let objectInstances: VisualObjectInstance[] = [];

            switch (options.objectName) {
                case 'line':
                    let instance: VisualObjectInstance = {
                        selector: null,
                        properties: {
                            lineColor: this.lineColor,
                            transparency: this.lineTransparency,
                            weight: this.lineWeight
                        },
                        objectName: options.objectName
                    };
                    if (this.data.shapeType === powerbi.basicShapeType.rectangle) {
                        instance.properties['roundEdge'] = this.roundEdge;
                    }

                    objectInstances.push(instance);
                    return objectInstances;
                case 'fill':
                    objectInstances.push({
                        selector: null,
                        properties: {
                            show: this.showFill,
                            fillColor: this.fillColor,
                            transparency: this.shapeTransparency
                        },
                        objectName: options.objectName
                    });

                    return objectInstances;
                case 'lockAspect':
                    objectInstances.push({
                        selector: null,
                        properties: {
                            show: this.lockAspectRatio
                        },
                        objectName: options.objectName
                    });

                    return objectInstances;
                case 'rotation':
                    objectInstances.push({
                        selector: null,
                        properties: {
                            angle: this.angle
                        },
                        objectName: options.objectName
                    });

                    return objectInstances;
            }

            return null;

        }

        public render(): void {
            this.selection.html('');
            
            switch (this.shapeType) {
                case powerbi.basicShapeType.rectangle:
                    shapeFactory.createRectangle(
                        this.data, this.currentViewport.height, this.currentViewport.width, this.selection, this.angle);
                    break;
                case powerbi.basicShapeType.oval:
                    shapeFactory.createOval(
                        this.data, this.currentViewport.height, this.currentViewport.width, this.selection, this.angle);
                    break;
                case powerbi.basicShapeType.line:
                    shapeFactory.createLine(
                        this.data, this.currentViewport.height, this.currentViewport.width, this.selection, this.angle);
                    break;
                case powerbi.basicShapeType.arrow:
                    shapeFactory.createUpArrow(
                        this.data, this.currentViewport.height, this.currentViewport.width, this.selection, this.angle);
                    break;
                case powerbi.basicShapeType.triangle:
                    shapeFactory.createTriangle(
                        this.data, this.currentViewport.height, this.currentViewport.width, this.selection, this.angle);
                    break;
                default:
                    break;
            }
        }
    }
}
