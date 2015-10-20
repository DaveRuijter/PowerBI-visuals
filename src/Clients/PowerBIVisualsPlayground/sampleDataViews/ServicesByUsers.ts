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

    export class ServicesByUsers extends SampleDataViews implements ISampleDataViewsMethods {

        public name: string = "ServicesByUsers";
        public displayName: string = "Services by users";

        public visuals: string[] = ['radarChart'];

        private sampleData = [
            [59, 56, 42, 34, 48, 14, 7, 78, 85, 90, 18, 7, 8, 9, 10],
            [48, 46, 29, 11, 14, 5, 14, 6, 24, 17, 15, 12, 67, 56, 16]
        ];

        private sampleMin: number = 1;
        private sampleMax: number = 100;

        private sampleSingleData: number = 50;

        public getDataViews(): DataView[] {

            let fieldExpr = powerbi.data.SQExprBuilder.fieldExpr({ column: { schema: 's', entity: "table1", name: "country" } });

            const categoryValues = ["Email", "Social Networks", "Internet Banking", "News Sportsites", "Search Engine",
                "View Shopping sites", "Paying Online", "Buy Online", "Online Gaming", "Offline Gaming", "Photo Video",
                "Reading", "Listen Music", "Watch TV", "Listen Radio"];
            let categoryIdentities = categoryValues.map(function(value) {
                let expr = powerbi.data.SQExprBuilder.equal(fieldExpr, powerbi.data.SQExprBuilder.text(value));
                return powerbi.data.createDataViewScopeIdentity(expr);
            });
        
            // Metadata, describes the data columns, and provides the visual with hints
            // so it can decide how to best represent the data
            let dataViewMetadata: powerbi.DataViewMetadata = {
                columns: [
                    {
                        displayName: 'Devices',
                        queryName: 'Devices',
                        type: powerbi.ValueType.fromDescriptor({ text: true })
                    },
                    {
                        displayName: 'Smartphone',
                        isMeasure: true,
                        format: "0.00",
                        queryName: 'smartphone',
                        type: powerbi.ValueType.fromDescriptor({ numeric: true }),
                        objects: { dataPoint: { fill: { solid: { color: '#1F77B4' } } } },
                    },
                    {
                        displayName: 'Tablet',
                        isMeasure: true,
                        format: "0.00",
                        queryName: 'Tablet',
                        type: powerbi.ValueType.fromDescriptor({ numeric: true }),
                        objects: { dataPoint: { fill: { solid: { color: '#FF7F0E' } } } }
                    }
                ]
            };

            let columns = [
                {
                    source: dataViewMetadata.columns[1],
                    values: this.sampleData[0],
                },
                {
                    source: dataViewMetadata.columns[2],
                    values: this.sampleData[1],
                }
            ];

            let dataValues: DataViewValueColumns = DataViewTransform.createValueColumns(columns);
            let tableDataValues = categoryValues.map(function(countryName, idx) {
                return [countryName, columns[0].values[idx], columns[1].values[idx]];
            });

            return [{
                metadata: dataViewMetadata,
                categorical: {
                    categories: [{
                        source: dataViewMetadata.columns[0],
                        values: categoryValues,
                        identity: categoryIdentities,
                    }],
                    values: dataValues
                },
                table: {
                    rows: tableDataValues,
                    columns: dataViewMetadata.columns,
                },
                single: { value: this.sampleSingleData }
            }];
        }

        public randomize(): void {

            this.sampleData = this.sampleData.map((item) => {
                return item.map(() => this.getRandomValue(this.sampleMin, this.sampleMax));
            });

            this.sampleSingleData = this.getRandomValue(this.sampleMin, this.sampleMax);
        }
    }
}
