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

    export class ProfitLossData extends SampleDataViews implements ISampleDataViewsMethods {

        public name: string = "ProfitLossData";
        public displayName: string = "Profit/Loss data";

        public visuals: string[] = ['waterfallChart'];

        private sampleData = [
            [742731.43, -162066.43, 283085.78, -300263.49, 376074.57, -814724.34]
        ];
        
        private sampleMin: number = -1000000;
        private sampleMax: number = 1000000;
        
        public getDataViews(): DataView[] {

            let fieldExpr = powerbi.data.SQExprBuilder.fieldExpr({ column: { schema: 's', entity: "table1", name: "country" } });

            let categoryValues = ["Australia", "Canada", "France", "Germany", "United Kingdom", "United States"];
            let categoryIdentities = categoryValues.map(function (value) {
                let expr = powerbi.data.SQExprBuilder.equal(fieldExpr, powerbi.data.SQExprBuilder.text(value));
                return powerbi.data.createDataViewScopeIdentity(expr);
            });
        
            // Metadata, describes the data columns, and provides the visual with hints
            // so it can decide how to best represent the data
            let dataViewMetadata: powerbi.DataViewMetadata = {
                columns: [
                    {
                        displayName: 'Country',
                        groupName: 'Country',
                        queryName: 'Country',
                        type: powerbi.ValueType.fromDescriptor({ text: true })
                    },
                    {
                        displayName: 'Sales Amount (2014)',
                        groupName: 'Sales Amount (2014)',
                        queryName: 'Sales Amount (2014)',
                        format: "$0,000.00",
                        type: powerbi.ValueType.fromDescriptor({ numeric: true }),
                        objects: { dataPoint: { fill: { solid: { color: 'purple' } } } },
                    }
                ]
            };

            let columns = [
                {
                    source: dataViewMetadata.columns[1],
                    // Sales Amount for 2014
                    values: this.sampleData[0],
                }
            ];

            let seriesIdentityField = powerbi.data.SQExprBuilder.fieldExpr({ column: { schema: 's', entity: 'e', name: 'series' } });

            let dataValues: DataViewValueColumns = DataViewTransform.createValueColumns(columns);
            return [{
                metadata: dataViewMetadata,
                categorical: {
                    categories: [{
                        source: dataViewMetadata.columns[0],
                        values: categoryValues,
                        identity: categoryIdentities,
                        identityFields: [seriesIdentityField]
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