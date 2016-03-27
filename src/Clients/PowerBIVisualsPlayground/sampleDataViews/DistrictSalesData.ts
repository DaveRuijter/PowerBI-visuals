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

module powerbi.visuals.sampleDataViews {
    import DataViewTransform = powerbi.data.DataViewTransform;

    export class DistrictSalesData extends SampleDataViews implements ISampleDataViewsMethods {

        public name: string = "DistrictSalesData";
        public displayName: string = "District Sales";

        public visuals: string[] = ['scatterChart'];

        private sampleData = [
            [300, 100, 200, 120, 80, 50, 240],
            [200, 150, 70, 30, 68, 145, 190],
            [30, 5, 10, 20, 40, 60, 70]
        ];

        private sampleMin: number = 10;
        private sampleMax: number = 300;
        
        public getDataViews(): DataView[] {

            let fieldExpr = powerbi.data.SQExprBuilder.fieldExpr({ column: { schema: 's', entity: "table1", name: "country" } });

            let categoryValues = ["FD - 01", "FD - 02", "FD - 03", "FD - 04", "LI - 01", "LI - 02", "LI - 03"];
            let categoryIdentities = categoryValues.map(function (value) {
                var expr = powerbi.data.SQExprBuilder.equal(fieldExpr, powerbi.data.SQExprBuilder.text(value));
                return powerbi.data.createDataViewScopeIdentity(expr);
            });
        
            // Metadata, describes the data columns, and provides the visual with hints
            // so it can decide how to best represent the data
            let dataViewMetadata: powerbi.DataViewMetadata = {
                columns: [
                    {
                        displayName: 'series',
                        isMeasure: false,
                        queryName: 'series',
                        roles: {
                            "Category": true,
                            "Series": true
                        },
                        type: ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Text)
                    },
                    {
                        displayName: 'Total Sales Variance %',
                        groupName: 'Total Sales Variance %',
                        isMeasure: true,
                        queryName: "x",
                        roles: { "X": true },
                        type: ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Double)
                    },
                    {
                        displayName: 'Sales Per Sq Ft',
                        groupName: 'Sales Per Sq Ft',
                        isMeasure: true,
                        queryName: "y",
                        roles: { "Y": true },
                        type: ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Double)
                    },
                    {
                        displayName: 'valueSize',
                        groupName: 'valueSize',
                        isMeasure: true,
                        queryName: "size",
                        roles: { "Size": true },
                        type: ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Double)
                    }
                ],
                objects: { categoryLabels: { show: true } },
            };
            
            let columns = [
                {
                    source: dataViewMetadata.columns[1],
                    values: this.sampleData[0]
                }, {
                    source: dataViewMetadata.columns[2],
                    values: this.sampleData[1]
                }, {
                    source: dataViewMetadata.columns[3],
                    values: this.sampleData[2]
                }];
                        
            let dataValues: DataViewValueColumns = DataViewTransform.createValueColumns(columns);
 
            let seriesIdentityField = powerbi.data.SQExprBuilder.fieldExpr({ column: { schema: 's', entity: 'e', name: 'series' } });
            
            return [{
                metadata: dataViewMetadata,
                    categorical: {
                        categories: [{
                            source: dataViewMetadata.columns[0],
                            values: categoryValues,
                            identity: categoryIdentities,
                            identityFields: [seriesIdentityField],
                            objects: [
                                { dataPoint: { fill: { solid: { color: 'red' } } } },
                                { dataPoint: { fill: { solid: { color: 'blue' } } } },
                                { dataPoint: { fill: { solid: { color: 'green' } } } },
                                { dataPoint: { fill: { solid: { color: 'black' } } } },
                                { dataPoint: { fill: { solid: { color: 'yellow' } } } },
                                { dataPoint: { fill: { solid: { color: 'brown' } } } },
                                { dataPoint: { fill: { solid: { color: 'pink' } } } }]
                        }],
                        values: dataValues
                    }
                }];
        }

        
        public randomize(): void {

            this.sampleData = this.sampleData.map((item) => {
                return item.map(() => this.getRandomValue(this.sampleMin, this.sampleMax));
            });
        }
        
    }
}