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
    import ValueType = powerbi.ValueType;
    import PrimitiveType = powerbi.PrimitiveType;

    export class SimpleMatrixData extends SampleDataViews implements ISampleDataViewsMethods {

        public name: string = "SimpleMatrixData";
        public displayName: string = "Simple matrix data";

        public visuals: string[] = ['matrix','sunburst'
        ];

        public getDataViews(): DataView[] {
            var dataTypeNumber = ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Double);
            var dataTypeString = ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Text);

            var measureSource1: DataViewMetadataColumn = { displayName: 'measure1', type: dataTypeNumber, isMeasure: true, index: 3, objects: { general: { formatString: '#.0' } } };
            var measureSource2: DataViewMetadataColumn = { displayName: 'measure2', type: dataTypeNumber, isMeasure: true, index: 4, objects: { general: { formatString: '#.00' } } };
            var measureSource3: DataViewMetadataColumn = { displayName: 'measure3', type: dataTypeNumber, isMeasure: true, index: 5, objects: { general: { formatString: '#' } } };

            var rowGroupSource1: DataViewMetadataColumn = { displayName: 'RowGroup1', queryName: 'RowGroup1', type: dataTypeString, index: 0 };
            var rowGroupSource1Column = powerbi.data.SQExprBuilder.fieldExpr({ column: { schema: 's', entity: "t", name: "rowgroup1" } });
            var rowGroupSource2: DataViewMetadataColumn = { displayName: 'RowGroup2', queryName: 'RowGroup2', type: dataTypeString, index: 1 };
            var rowGroupSource2Column = powerbi.data.SQExprBuilder.fieldExpr({ column: { schema: 's', entity: "t", name: "rowgroup2" } });
            var rowGroupSource3: DataViewMetadataColumn = { displayName: 'RowGroup3', queryName: 'RowGroup3', type: dataTypeString, index: 2 };
            var rowGroupSource3Column = powerbi.data.SQExprBuilder.fieldExpr({ column: { schema: 's', entity: "t", name: "rowgroup3" } });

            var matrixThreeMeasuresThreeRowGroups: DataViewMatrix = {
                rows: {
                    root: {
                        children: [
                            {
                             
                                value: 'North America',
                                children: [
                                    {
                                       
                                        value: 'Canada',
                                        children: [
                                            {
                                              
                                                value: 'Ontario',
                                                values: {
                                                    0: { value: 1000 },
                                                    1: { value: 1001, valueSourceIndex: 1 },
                                                    2: { value: 1002, valueSourceIndex: 2 }
                                                },
                                                identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource3Column, powerbi.data.SQExprBuilder.text('Ontario'))),
                                            },
                                            {
                                               
                                                value: 'Quebec',
                                                values: {
                                                    0: { value: 1010 },
                                                    1: { value: 1011, valueSourceIndex: 1 },
                                                    2: { value: 1012, valueSourceIndex: 2 }
                                                },
                                                identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource3Column, powerbi.data.SQExprBuilder.text('Quebec'))),
                                            }
                                        ],
                                        identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource2Column, powerbi.data.SQExprBuilder.text('Canada'))),
                                        childIdentityFields: [rowGroupSource3Column]
                                    },
                                    {
                                        
                                        value: 'USA',
                                        children: [
                                            {
                                             
                                                value: 'Washington',
                                                values: {
                                                    0: { value: 1100 },
                                                    1: { value: 1101, valueSourceIndex: 1 },
                                                    2: { value: 1102, valueSourceIndex: 2 }
                                                },
                                                identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource3Column, powerbi.data.SQExprBuilder.text('Washington'))),
                                            },
                                            {
                                               
                                                value: 'Oregon',
                                                values: {
                                                    0: { value: 1110 },
                                                    1: { value: 1111, valueSourceIndex: 1 },
                                                    2: { value: 1112, valueSourceIndex: 2 }
                                                },
                                                identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource3Column, powerbi.data.SQExprBuilder.text('Oregon'))),
                                            }
                                        ],
                                        identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource2Column, powerbi.data.SQExprBuilder.text('USA'))),
                                        childIdentityFields: [rowGroupSource3Column]
                                    }
                                ],
                                identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource1Column, powerbi.data.SQExprBuilder.text('North America'))),
                                childIdentityFields: [rowGroupSource2Column]
                            },
                            {
                              
                                value: 'South America',
                                children: [
                                    {
                                      
                                        value: 'Brazil',
                                        children: [
                                            {
                                               
                                                value: 'Amazonas',
                                                values: {
                                                    0: { value: 2000 },
                                                    1: { value: 2001, valueSourceIndex: 1 },
                                                    2: { value: 2002, valueSourceIndex: 2 }
                                                },
                                                identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource3Column, powerbi.data.SQExprBuilder.text('Amazonas'))),
                                            },
                                            {
                                             
                                                value: 'Mato Grosso',
                                                values: {
                                                    0: { value: 2010 },
                                                    1: { value: 2011, valueSourceIndex: 1 },
                                                    2: { value: 2012, valueSourceIndex: 2 }
                                                },
                                                identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource3Column, powerbi.data.SQExprBuilder.text('Mato Grosso'))),
                                            }
                                        ],
                                        identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource2Column, powerbi.data.SQExprBuilder.text('Brazil'))),
                                        childIdentityFields: [rowGroupSource3Column]
                                    },
                                    {
                                       
                                        value: 'Chile',
                                        children: [
                                            {
                                            
                                                value: 'Arica',
                                                values: {
                                                    0: { value: 2100 },
                                                    1: { value: 2101, valueSourceIndex: 1 },
                                                    2: { value: 2102, valueSourceIndex: 2 }
                                                },
                                                identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource3Column, powerbi.data.SQExprBuilder.text('Arica'))),
                                            },
                                            {
                                       
                                                value: 'Parinacota',
                                                values: {
                                                    0: { value: 2110 },
                                                    1: { value: 2111, valueSourceIndex: 1 },
                                                    2: { value: 2112, valueSourceIndex: 2 }
                                                },
                                                identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource3Column, powerbi.data.SQExprBuilder.text('Parinacota'))),
                                            }
                                        ],
                                        identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource2Column, powerbi.data.SQExprBuilder.text('Chile'))),
                                        childIdentityFields: [rowGroupSource3Column]
                                    }
                                ],
                                identity: powerbi.data.createDataViewScopeIdentity(powerbi.data.SQExprBuilder.equal(rowGroupSource1Column, powerbi.data.SQExprBuilder.text('South America'))),
                                childIdentityFields: [rowGroupSource2Column]
                            },

                        ],
                        childIdentityFields: [rowGroupSource1Column],
                    },
                    levels: [
                        { sources: [rowGroupSource1] },
                        { sources: [rowGroupSource2] },
                        { sources: [rowGroupSource3] }
                    ]
                },
                columns: {
                    root: {
                        children: [
                            { level: 0 },
                            { level: 0, levelSourceIndex: 1 },
                            { level: 0, levelSourceIndex: 2 }
                        ]
                    },
                    levels: [{
                        sources: [
                            measureSource1,
                            measureSource2,
                            measureSource3
                        ]
                    }]
                },
                valueSources: [
                    measureSource1,
                    measureSource2,
                    measureSource3
                ]
            };

            return [{
                metadata: { columns: [rowGroupSource1, rowGroupSource2, rowGroupSource3], segment: {} },
                matrix: matrixThreeMeasuresThreeRowGroups
            }];
        }

        public randomize(): void {
        }

    }
}