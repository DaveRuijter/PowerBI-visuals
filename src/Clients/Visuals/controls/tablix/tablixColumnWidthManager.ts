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

/// <reference path="../../_references.ts"/>

module powerbi.visuals.controls {
    export const AutoSizeColumnWidthDefault: boolean = true;

    export interface TablixColumnWidthObject {
        queryName: string;
        width: number;
    }

    export class TablixColumnWidthManager {
        public static columnWidthProp: DataViewObjectPropertyIdentifier = { objectName: 'general', propertyName: 'columnWidth' };
        public static autoSizeWidthProp: DataViewObjectPropertyIdentifier = { objectName: 'general', propertyName: 'autoSizeColumnWidth' };

        private columnWidths: number[];
        private tablixColumnWidthsObject: controls.TablixColumnWidthObject[];
        private previousAutoColumnSizePropertyValue: boolean;
        private tablixQueryNames: string[];
        private dataView: DataView;
        private visualObjectInstancesToPersist: VisualObjectInstance[];
        private matrixLeafNodes: MatrixVisualNode[];
        // TODO: Can we remove isMatrix and suppresNotification flag?
        private isMatrix: boolean;
        private suppressNotification: boolean;
        private currentPersistedWidths: number[];
        private callHostPersistProperties: boolean;
        private dataViewUpdated: boolean;

        constructor(dataView: DataView, isMatrix: boolean, matrixLeafNodes?: MatrixVisualNode[]) {
            this.columnWidths = [];
            this.tablixColumnWidthsObject = [];
            this.tablixQueryNames = [];
            this.dataView = dataView;
            this.matrixLeafNodes = matrixLeafNodes;
            this.isMatrix = isMatrix;
            this.suppressNotification = false;
            this.currentPersistedWidths = [];
            this.visualObjectInstancesToPersist = [];
            this.callHostPersistProperties = false;
            this.dataViewUpdated = false;
        }

        public getColumnWidths(): number[] {
            return this.columnWidths;
        }

        public getTablixColumnWidthsObject(): controls.TablixColumnWidthObject[] {
            return this.tablixColumnWidthsObject;
        }

        public updateDataView(dataView: DataView, matrixLeafNodes?: MatrixVisualNode[]): void {
            this.dataView = dataView;
            this.matrixLeafNodes = matrixLeafNodes;
            this.dataViewUpdated = true;
        }

        public getVisualObjectInstancesToPersist(): VisualObjectInstance[] {
            return this.visualObjectInstancesToPersist;
        }

        public persistColumnWidthsOnHost(): boolean {
            return this.callHostPersistProperties;
        }

        public getTablixQueryNames(): string[]{
            return this.tablixQueryNames;
        }

        public get suppressOnDataChangedNotification(): boolean {
            return this.suppressNotification;
        }

        public set suppressOnDataChangedNotification(notify: boolean) {
            this.suppressNotification = notify;
        }

        public deserializeTablixColumnWidths(): void {
            let columnMetaData = this.dataView && this.dataView.metadata && this.dataView.metadata.columns;
            if (columnMetaData) {
                let autoSize = this.shouldAutoSizeColumnWidth();
                if (this.previousAutoColumnSizePropertyValue !== undefined && !this.previousAutoColumnSizePropertyValue && this.previousAutoColumnSizePropertyValue !== autoSize) {
                    // Auto-Size false to true. In this case blow away any saved widths and revert back to default of calculating column sizes
                    this.columnWidths = [];
                    this.removePersistedVisualObjectInstances();
                    this.tablixColumnWidthsObject = [];
                    this.callHostPersistProperties = true;
                }
                else {
                    this.deserializeColumnWidths(columnMetaData);
                    this.callHostPersistProperties = false;
                }

                this.previousAutoColumnSizePropertyValue = autoSize;
            }
        }

        public columnWidthChanged(index: number, width: number): void {
            this.populateTablixQueryNames();
            if (this.tablixQueryNames.length === 0)
                return;
            // Only persist manually resized columns
            let colIndex = this.isMatrix ? index : Math.max(index - 1, 0); // Table indices are offset with an empty header.
            let newColumnWidthObj = {
                queryName: this.tablixQueryNames[colIndex],
                width: width
            };

            let objExists = false;
            for (let i = 0, len = this.tablixColumnWidthsObject.length; i < len; i++) {
                if (this.tablixColumnWidthsObject[i].queryName === newColumnWidthObj.queryName) {
                    this.tablixColumnWidthsObject[i] = newColumnWidthObj;
                    objExists = true;
                }
            }
            if (!objExists)
                this.tablixColumnWidthsObject.push(newColumnWidthObj);

            // Let onDataChanged call pass through so all instances of resized column gets updated for the matrix
            this.suppressNotification = !this.isMatrix;
            this.generateVisualObjectInstancesToPersist();
            this.callHostPersistProperties = true;
            // Keep all persisted column widths in sync
            if (!this.shouldAutoSizeColumnWidth() && this.currentPersistedWidths.length > 0 && index < this.currentPersistedWidths.length)
                this.currentPersistedWidths[index] = width;
        }

        public persistAllColumnWidths(widthsToPersist: number[]): void {
            if (!this.shouldAutoSizeColumnWidth() && this.dataViewUpdated) {
                let shouldPersist = true;
                if (!_.isEqual(widthsToPersist, this.currentPersistedWidths)) {
                    this.currentPersistedWidths = widthsToPersist;
                    this.suppressNotification = true;
                }
                else {
                    shouldPersist = false;
                }
                // Persist all column widths
                let widths = this.isMatrix ? widthsToPersist
                    : widthsToPersist.slice(1, widthsToPersist.length); // Table indices are offset with an empty header. 
                let savedColumnWidths: controls.TablixColumnWidthObject[] = [];

                if (this.tablixQueryNames.length !== widths.length) {
                    shouldPersist = false;
                }

                if (!shouldPersist) {
                    this.suppressNotification = false;
                    this.callHostPersistProperties = false;
                    return;
                }

                for (let colIndex = 0, len = widths.length; colIndex < len; colIndex++) {
                    let widthObj = {
                        queryName: this.tablixQueryNames[colIndex],
                        width: widths[colIndex]
                    };

                    for (let columnWidthObject of this.tablixColumnWidthsObject) {
                        if (columnWidthObject.queryName === widthObj.queryName)
                            widthObj.width = columnWidthObject.width;
                    }
                    savedColumnWidths.push(widthObj);
                }

                this.tablixColumnWidthsObject = savedColumnWidths;
                this.generateVisualObjectInstancesToPersist();
                this.callHostPersistProperties = true;
            }
        }

        public shouldAutoSizeColumnWidth(): boolean {
            let objects = this.dataView.metadata.objects;
            if (objects && objects[TablixColumnWidthManager.autoSizeWidthProp.objectName]) {
                return objects[TablixColumnWidthManager.autoSizeWidthProp.objectName][TablixColumnWidthManager.autoSizeWidthProp.propertyName] !== false;
            }

            // Auto adjust is turned on by default
            return AutoSizeColumnWidthDefault;
        }

        private generateVisualObjectInstancesToPersist(): void {
            this.visualObjectInstancesToPersist = [];
            // Auto-Size property
            this.visualObjectInstancesToPersist.push({
                selector: null,
                objectName: 'general',
                properties: {
                    autoSizeColumnWidth: data.SQExprBuilder.boolean(this.shouldAutoSizeColumnWidth())
                }
            });

            // Column Widths
            for (let columnWidthObject of this.tablixColumnWidthsObject) {
                let queryNameSelector = columnWidthObject.queryName;
                // Only persist width if we have a valid queryName to use as selector
                if (!queryNameSelector)
                    continue;
                this.visualObjectInstancesToPersist.push({
                    selector: { metadata: queryNameSelector },
                    objectName: 'general',
                    properties: {
                        columnWidth: data.SQExprBuilder.double(columnWidthObject.width)
                    }
                });
            }
        }

        private removePersistedVisualObjectInstances(): void {
            this.visualObjectInstancesToPersist = [];
            // Auto-Size property
            this.visualObjectInstancesToPersist.push({
                selector: null,
                objectName: 'general',
                properties: {
                    autoSizeColumnWidth: data.SQExprBuilder.boolean(this.shouldAutoSizeColumnWidth())
                }
            });

            for (let columnWidthObject of this.tablixColumnWidthsObject) {
                this.visualObjectInstancesToPersist.push({
                    selector: { metadata: columnWidthObject.queryName },
                    objectName: 'general',
                    properties: {
                        columnWidth: undefined
                    }
                });
            }
        }

        private deserializeColumnWidths(columnMetaData: DataViewMetadataColumn[]): void {
            this.columnWidths.length = 0;
            this.tablixColumnWidthsObject.length = 0;
            this.populateTablixQueryNames();
            if (this.tablixQueryNames.length === 0)
                return;
            for (let column of columnMetaData) {
                let columnWidthPropValue = DataViewObjects.getValue<number>(column.objects, TablixColumnWidthManager.columnWidthProp);
                if (columnWidthPropValue === null || columnWidthPropValue === undefined)
                    continue;
                for (let colIndex = 0, len = this.tablixQueryNames.length; colIndex < len; colIndex++) {
                    let propertySelector = this.tablixQueryNames[colIndex];
                    if (column.queryName === propertySelector) {
                        let columnWidth = columnWidthPropValue;
                        let index = this.isMatrix ? colIndex : colIndex + 1; // Table indices are offset with an empty header.
                        this.columnWidths[index] = columnWidth;
                        this.tablixColumnWidthsObject.push({
                            queryName: propertySelector,
                            width: columnWidth
                        });
                    }
                }
            }
        }

        private populateTablixQueryNames(): void {
            if (this.isMatrix)
                this.getMatrixQueryNames();
            else
                this.getTableQueryNames();
        }

        private getTableQueryNames(): void {
            this.tablixQueryNames.length = 0;
            let columnMetaData = this.dataView.table.columns;
            for (let column of columnMetaData) {
                this.tablixQueryNames.push(column.queryName);
            }
        }

        private getMatrixQueryNames(): void {
            this.tablixQueryNames.length = 0;
            let matrixRows = this.dataView.matrix.rows;
            for (let level of matrixRows.levels) {
                for (let source of level.sources) {
                    this.tablixQueryNames.push(source.queryName);
                }
            }

            if (this.matrixLeafNodes) {
                for (let leafNode of this.matrixLeafNodes) {
                    this.tablixQueryNames.push(leafNode.queryName);
                }
            }
        }
    }
}