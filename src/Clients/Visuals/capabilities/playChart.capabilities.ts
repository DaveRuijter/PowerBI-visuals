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

module powerbi.visuals {
    export var playChartCapabilities: VisualCapabilities = {
        dataRoles: [
            {
                name: 'Rows',
                kind: VisualDataRoleKind.Grouping,
                displayName: 'Play / Detail',
            }, {
                name: 'Columns',
                kind: VisualDataRoleKind.Grouping,
                displayName: data.createDisplayNameGetter('Role_DisplayName_Legend'),
            }, {
                name: 'Values',
                kind: VisualDataRoleKind.Measure,
                displayName: 'X / Y / Size',
            }
        ],
        objects: {
            dataPoint: {
                displayName: data.createDisplayNameGetter('Visual_DataPoint'),
                properties: {
                    defaultColor: {
                        displayName: data.createDisplayNameGetter('Visual_DefaultColor'),
                        type: { fill: { solid: { color: true } } }
                    },
                    showAllDataPoints: {
                        displayName: data.createDisplayNameGetter('Visual_DataPoint_Show_All'),
                        type: { bool: true }
                    },
                    fill: {
                        displayName: data.createDisplayNameGetter('Visual_Fill'),
                        type: { fill: { solid: { color: true } } }
                    },
                    fillRule: {
                        displayName: data.createDisplayNameGetter('Visual_Gradient'),
                        type: { fillRule: {} },
                        rule: {
                            inputRole: 'Gradient',
                            output: {
                                property: 'fill',
                                selector: ['Category'],
                            },
                        },
                    },
                }
            },
            general: {
                displayName: data.createDisplayNameGetter('Visual_General'),
                properties: {
                    formatString: {
                        type: { formatting: { formatString: true } },
                    },
                    colorByCategory: {
                        displayName: 'Color by Category',//data.createDisplayNameGetter('Visual_Show'),
                        type: { bool: true }
                    },
                    frozenIndex: {
                        type: { numeric: true },
                    },
                },
            },
            categoryAxis: {
                displayName: data.createDisplayNameGetter('Visual_XAxis'),
                properties: {
                    show: {
                        displayName: data.createDisplayNameGetter('Visual_Show'),
                        type: { bool: true }
                    },
                    start: {
                        displayName: data.createDisplayNameGetter('Visual_Axis_Start'),
                        type: { numeric: true }
                    },
                    end: {
                        displayName: data.createDisplayNameGetter('Visual_Axis_End'),
                        type: { numeric: true }
                    },
                    showAxisTitle: {
                        displayName: data.createDisplayNameGetter('Visual_Axis_Title'),
                        type: { bool: true }
                    },
                    axisStyle: {
                        displayName: data.createDisplayNameGetter('Visual_Axis_Style'),
                        type: { formatting: { axisStyle: true } }
                    }
                }
            },
            valueAxis: {
                displayName: data.createDisplayNameGetter('Visual_YAxis'),
                properties: {
                    show: {
                        displayName: data.createDisplayNameGetter('Visual_Show'),
                        type: { bool: true }
                    },
                    position: {
                        displayName: data.createDisplayNameGetter('Visual_YAxis_Position'),
                        type: { formatting: { yAxisPosition: true } }
                    },
                    start: {
                        displayName: data.createDisplayNameGetter('Visual_Axis_Start'),
                        type: { numeric: true }
                    },
                    end: {
                        displayName: data.createDisplayNameGetter('Visual_Axis_End'),
                        type: { numeric: true }
                    },
                    showAxisTitle: {
                        displayName: data.createDisplayNameGetter('Visual_Axis_Title'),
                        type: { bool: true }
                    },
                    axisStyle: {
                        displayName: data.createDisplayNameGetter('Visual_Axis_Style'),
                        type: { formatting: { axisStyle: true } }
                    }

                }
            },
            legend: {
                displayName: data.createDisplayNameGetter('Visual_Legend'),
                properties: {
                    show: {
                        displayName: data.createDisplayNameGetter('Visual_Show'),
                        type: { bool: true }
                    },
                    position: {
                        displayName: data.createDisplayNameGetter('Visual_LegendPosition'),
                        type: { formatting: { legendPosition: true } }
                    },
                    showTitle: {
                        displayName: data.createDisplayNameGetter('Visual_LegendShowTitle'),
                        type: { bool: true }
                    },
                    titleText: {
                        displayName: data.createDisplayNameGetter('Visual_LegendTitleText'),
                        type: { text: true }
                    }
                }
            },
            categoryLabels: {
                displayName: data.createDisplayNameGetter('Visual_CategoryLabels'),
                properties: {
                    show: {
                        displayName: data.createDisplayNameGetter('Visual_Show'),
                        type: { bool: true }
                    },
                    color: {
                        displayName: data.createDisplayNameGetter('Visual_LabelsFill'),
                        type: { fill: { solid: { color: true } } }
                    },
                },
            },
        },

        dataViewMappings: [{
            conditions: [
                { 'Rows': { min: 1, max: 2 }, 'Columns': { min: 0, max: 1 }, 'Values': { min: 0, max: 3 } },
                { 'Rows': { min: 0, max: 2 }, 'Columns': { min: 1, max: 1 }, 'Values': { min: 0, max: 3 } },
                { 'Rows': { min: 0, max: 2 }, 'Columns': { min: 0, max: 1 }, 'Values': { min: 1, max: 3 } },
            ],
            matrix: {
                rows: {
                    for: { in: 'Rows' },
                    /* Explicitly override the server data reduction to make it appropriate for matrix. */
                    dataReductionAlgorithm: { window: { count: 1000 } }
                },
                columns: {
                    for: { in: 'Columns' },
                    /* Explicitly override the server data reduction to make it appropriate for matrix. */
                    dataReductionAlgorithm: { top: { count: 100 } }
                },
                values: {
                    for: { in: 'Values' }
                }
            }
        }],
        sorting: {
            custom: {},
        },
    };

    export var playChartProps = {
        general: {
            formatString: <DataViewObjectPropertyIdentifier>{ objectName: 'general', propertyName: 'formatString' },
        },
        dataPoint: {
            defaultColor: <DataViewObjectPropertyIdentifier>{ objectName: 'dataPoint', propertyName: 'defaultColor' },
            fill: <DataViewObjectPropertyIdentifier>{ objectName: 'dataPoint', propertyName: 'fill' },
        },
    };
}
