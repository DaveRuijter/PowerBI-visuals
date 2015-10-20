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
    import StringExtensions = jsCommon.StringExtensions;

    export module KpiUtil {
        export const enum KpiImageSize {
            Small,
            Big,
        }

        export interface KpiImageMetadata {
            statusGraphic: string;
            caption: string;
        }

        const statusGraphicFormatStrings = {
            'Traffic Light - Single': {
                formatStr: 'kpiTrafficLightSingle{0}{1}',
                numOfValues: 3,
            },
            'Three Flags Colored': {
                formatStr: 'kpiThreeFlags{0}Colored{1}',
                numOfValues: 3,
            },
            'Road Signs': {
                formatStr: 'kpiRoadSigns{0}{1}',
                numOfValues: 3,
            },
            'Traffic Light': {
                formatStr: 'kpiTrafficLight{0}{1}',
                numOfValues: 3,
            },
            'Three Symbols UnCircled Colored': {
                formatStr: 'kpiThreeSymbolsUnCircled{0}Colored{1}',
                numOfValues: 3,
            },
            'Shapes': {
                formatStr: 'kpiShapes{0}{1}',
                numOfValues: 3,
            },
            'Three Stars Colored': {
                formatStr: 'kpiThreeStars{0}Colored{1}',
                numOfValues: 3,
            },
            'Five Bars Colored': {
                formatStr: 'kpiFiveBars{0}{1}',
                numOfValues: 5,
            },
            'Five Boxes Colored': {
                formatStr: 'kpiFiveBoxes{0}{1}',
                numOfValues: 5,
            },
            'Gauge - Ascending': {
                formatStr: 'kpiGauge{0}{1}',
                numOfValues: 5,
            }
        };

        export function getClassForKpi(statusGraphic: string, value: string, kpiImageSize?: KpiImageSize): string {
            debug.assertValue(statusGraphic, 'statusGraphic');
            debug.assertValue(value, 'value');

            var intValue = parseInt(value, 10);
            if (!isNaN(intValue)) {
                var sizeName: string = kpiImageSize && kpiImageSize === KpiImageSize.Big ? 'Big' : '';

                var numStr: string;
                var statusGraphicFormat = statusGraphicFormatStrings[statusGraphic];

                if (!statusGraphicFormat)
                    return undefined;

                var formatStr = statusGraphicFormat.formatStr;
                var numOfValues = statusGraphicFormat.numOfValues;
                
                // Convert values from the range of (-n/2, ..., 0, ..., n/2) to (0, 1, ..., n-1)
                numStr = (intValue + Math.floor(numOfValues / 2)).toString();
                
                return StringExtensions.format(
                    formatStr,
                    sizeName,
                    numStr);
            }
        }

        export function getKpiImageMetadata(metaDataColumn: DataViewMetadataColumn, target: string, kpiImageSize?: KpiImageSize): KpiImageMetadata {
            let statusGraphic: string = metaDataColumn && metaDataColumn.kpiStatusGraphic;

            if (statusGraphic) {
                let columnCaption = getClassForKpi(statusGraphic, target, kpiImageSize);
                if (columnCaption) {
                    return {
                        caption: columnCaption,
                        statusGraphic: statusGraphic,
                    };
                }
            }
        }
    }
}