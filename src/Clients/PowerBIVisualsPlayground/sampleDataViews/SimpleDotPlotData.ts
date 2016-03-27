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
    import DataView = powerbi.DataView;
    import DataViewMetadata = powerbi.DataViewMetadata;
    import ValueType = powerbi.ValueType;
    import DataViewTransform = powerbi.data.DataViewTransform;

    export class SimpleDotPlotData 
        extends SampleDataViews
        implements ISampleDataViewsMethods {

        public name: string = "SimpleDotPlotData";
        public displayName: string = "Simple DotPlot Data";

        public visuals: string[] = ["dotPlot"];

        private minValue: number = 1;
        private maxValue: number = 9;

        private minLength: number = 10;
        private maxLength: number = 1000;

        private sampleData: number[] = [
            1, 2, 3, 4, 5, 6, 7, 8, 9,
            2, 3, 3, 4, 4, 4, 5, 5, 5, 5,
            6, 6, 6, 7, 7, 8
        ];

        public getDataViews(): DataView[] {
            let dataViewMetadata: DataViewMetadata = {
                    columns: [{
                        displayName: "Observations",
                        queryName: "Observations",
                        type: ValueType.fromDescriptor({
                            text: true
                        }),
                        objects: {
                            dataPoint: {
                                fill: {
                                    solid: {
                                        color: "rgb(1, 184, 170)"
                                    }
                                }
                            }
                        }
                    }]
                },
                columns = [{
                    source: dataViewMetadata.columns[0], 
                    values: this.sampleData
                }],
                dataValues: DataViewValueColumns = 
                    DataViewTransform.createValueColumns(columns);

            return [{
                metadata: dataViewMetadata,
                categorical: {
                    values: dataValues
                }
            }];
        }

        public randomize(): void {
            let length: number =
                Math.round(this.getRandomValue(this.minLength, this.maxLength));

            let values = [];
            for (let i = this.minValue; i <= this.maxValue; i++) {
                values.push(this.getRandomValue(this.minValue, this.maxValue));
            }

            this.sampleData = [];

            for (let j = 0; j < length; j++) {
                let value: number = this.randomElement(values);

                this.sampleData.push(Math.round(value));
            }
        }
    }
}
