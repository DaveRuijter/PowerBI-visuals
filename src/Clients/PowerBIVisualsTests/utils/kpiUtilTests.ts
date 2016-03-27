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

module powerbitests {
    import KpiUtil = powerbi.visuals.KpiUtil;

    describe("kpiUnitTests", () => {
        it("Traffic light single bad", () => {
            verifyGetClassForKpi('Traffic Light - Single', '-1', KpiUtil.KpiImageSize.Small, 'kpiTrafficLightSingle0');
        });

        it("Traffic light single warning", () => {
            verifyGetClassForKpi('Traffic Light - Single', '0', KpiUtil.KpiImageSize.Small, 'kpiTrafficLightSingle1');
        });

        it("Traffic light single good", () => {
            verifyGetClassForKpi('Traffic Light - Single', '1', KpiUtil.KpiImageSize.Small, 'kpiTrafficLightSingle2');
        });

        it("Traffic light single big image bad", () => {
            verifyGetClassForKpi('Traffic Light - Single', '-1', KpiUtil.KpiImageSize.Big, 'kpiTrafficLightSingleBig0');
        });

        it("Traffic light single big image warning", () => {
            verifyGetClassForKpi('Traffic Light - Single', '0', KpiUtil.KpiImageSize.Big, 'kpiTrafficLightSingleBig1');
        });

        it("Traffic light single big image good", () => {
            verifyGetClassForKpi('Traffic Light - Single', '1', KpiUtil.KpiImageSize.Big, 'kpiTrafficLightSingleBig2');
        });

        it("Three Flags Colored bad", () => {
            verifyGetClassForKpi('Three Flags Colored', '-1', KpiUtil.KpiImageSize.Small, 'kpiThreeFlagsColored0');
        });

        it("Three Flags Colored warning", () => {
            verifyGetClassForKpi('Three Flags Colored', '0', KpiUtil.KpiImageSize.Small, 'kpiThreeFlagsColored1');
        });

        it("Three Flags Colored good", () => {
            verifyGetClassForKpi('Three Flags Colored', '1', KpiUtil.KpiImageSize.Small, 'kpiThreeFlagsColored2');
        });

        it("Three Flags Colored big image bad", () => {
            verifyGetClassForKpi('Three Flags Colored', '-1', KpiUtil.KpiImageSize.Big, 'kpiThreeFlagsBigColored0');
        });

        it("Three Flags Colored big image warning", () => {
            verifyGetClassForKpi('Three Flags Colored', '0', KpiUtil.KpiImageSize.Big, 'kpiThreeFlagsBigColored1');
        });

        it("Three Flags Colored big image good", () => {
            verifyGetClassForKpi('Three Flags Colored', '1', KpiUtil.KpiImageSize.Big, 'kpiThreeFlagsBigColored2');
        });

        it("Road Signs bad", () => {
            verifyGetClassForKpi('Road Signs', '-1', KpiUtil.KpiImageSize.Small, 'kpiRoadSigns0');
        });

        it("Road Signs warning", () => {
            verifyGetClassForKpi('Road Signs', '0', KpiUtil.KpiImageSize.Small, 'kpiRoadSigns1');
        });

        it("Road Signs good", () => {
            verifyGetClassForKpi('Road Signs', '1', KpiUtil.KpiImageSize.Small, 'kpiRoadSigns2');
        });

        it("Road Signs big image bad", () => {
            verifyGetClassForKpi('Road Signs', '-1', KpiUtil.KpiImageSize.Big, 'kpiRoadSignsBig0');
        });

        it("Road Signs big image warning", () => {
            verifyGetClassForKpi('Road Signs', '0', KpiUtil.KpiImageSize.Big, 'kpiRoadSignsBig1');
        });

        it("Road Signs big image good", () => {
            verifyGetClassForKpi('Road Signs', '1', KpiUtil.KpiImageSize.Big, 'kpiRoadSignsBig2');
        });

        it("Traffic Light bad", () => {
            verifyGetClassForKpi('Traffic Light', '-1', KpiUtil.KpiImageSize.Small, 'kpiTrafficLight0');
        });

        it("Traffic Light warning", () => {
            verifyGetClassForKpi('Traffic Light', '0', KpiUtil.KpiImageSize.Small, 'kpiTrafficLight1');
        });

        it("Traffic Light good", () => {
            verifyGetClassForKpi('Traffic Light', '1', KpiUtil.KpiImageSize.Small, 'kpiTrafficLight2');
        });

        it("Traffic Light big image bad", () => {
            verifyGetClassForKpi('Traffic Light', '-1', KpiUtil.KpiImageSize.Big, 'kpiTrafficLightBig0');
        });

        it("Traffic Light big image warning", () => {
            verifyGetClassForKpi('Traffic Light', '0', KpiUtil.KpiImageSize.Big, 'kpiTrafficLightBig1');
        });

        it("Traffic Light big image good", () => {
            verifyGetClassForKpi('Traffic Light', '1', KpiUtil.KpiImageSize.Big, 'kpiTrafficLightBig2');
        });

        it("Three Symbols UnCircled Colored bad", () => {
            verifyGetClassForKpi('Three Symbols UnCircled Colored', '-1', KpiUtil.KpiImageSize.Small, 'kpiThreeSymbolsUnCircledColored0');
        });

        it("Three Symbols UnCircled Colored warning", () => {
            verifyGetClassForKpi('Three Symbols UnCircled Colored', '0', KpiUtil.KpiImageSize.Small, 'kpiThreeSymbolsUnCircledColored1');
        });

        it("Three Symbols UnCircled Colored good", () => {
            verifyGetClassForKpi('Three Symbols UnCircled Colored', '1', KpiUtil.KpiImageSize.Small, 'kpiThreeSymbolsUnCircledColored2');
        });

        it("Three Symbols UnCircled Colored big image bad", () => {
            verifyGetClassForKpi('Three Symbols UnCircled Colored', '-1', KpiUtil.KpiImageSize.Big, 'kpiThreeSymbolsUnCircledBigColored0');
        });

        it("Three Symbols UnCircled Colored big image warning", () => {
            verifyGetClassForKpi('Three Symbols UnCircled Colored', '0', KpiUtil.KpiImageSize.Big, 'kpiThreeSymbolsUnCircledBigColored1');
        });

        it("Three Symbols UnCircled Colored big image good", () => {
            verifyGetClassForKpi('Three Symbols UnCircled Colored', '1', KpiUtil.KpiImageSize.Big, 'kpiThreeSymbolsUnCircledBigColored2');
        });

        it("Shapes bad", () => {
            verifyGetClassForKpi('Shapes', '-1', KpiUtil.KpiImageSize.Small, 'kpiShapes0');
        });

        it("Shapes warning", () => {
            verifyGetClassForKpi('Shapes', '0', KpiUtil.KpiImageSize.Small, 'kpiShapes1');
        });

        it("Shapes good", () => {
            verifyGetClassForKpi('Shapes', '1', KpiUtil.KpiImageSize.Small, 'kpiShapes2');
        });

        it("Shapes big image bad", () => {
            verifyGetClassForKpi('Shapes', '-1', KpiUtil.KpiImageSize.Big, 'kpiShapesBig0');
        });

        it("Shapes big image warning", () => {
            verifyGetClassForKpi('Shapes', '0', KpiUtil.KpiImageSize.Big, 'kpiShapesBig1');
        });

        it("Shapes big image good", () => {
            verifyGetClassForKpi('Shapes', '1', KpiUtil.KpiImageSize.Big, 'kpiShapesBig2');
        });

        it("Three Stars Colored bad", () => {
            verifyGetClassForKpi('Three Stars Colored', '-1', KpiUtil.KpiImageSize.Small, 'kpiThreeStarsColored0');
        });

        it("Three Stars Colored warning", () => {
            verifyGetClassForKpi('Three Stars Colored', '0', KpiUtil.KpiImageSize.Small, 'kpiThreeStarsColored1');
        });

        it("Three Stars Colored good", () => {
            verifyGetClassForKpi('Three Stars Colored', '1', KpiUtil.KpiImageSize.Small, 'kpiThreeStarsColored2');
        });

        it("Three Stars Colored big image bad", () => {
            verifyGetClassForKpi('Three Stars Colored', '-1', KpiUtil.KpiImageSize.Big, 'kpiThreeStarsBigColored0');
        });

        it("Three Stars Colored big image warning", () => {
            verifyGetClassForKpi('Three Stars Colored', '0', KpiUtil.KpiImageSize.Big, 'kpiThreeStarsBigColored1');
        });

        it("Three Stars Colored big image good", () => {
            verifyGetClassForKpi('Three Stars Colored', '1', KpiUtil.KpiImageSize.Big, 'kpiThreeStarsBigColored2');
        });

        it("Five Bars Colored 0", () => {
            verifyGetClassForKpi('Five Bars Colored', '-2', KpiUtil.KpiImageSize.Small, 'kpiFiveBars0');
        });

        it("Five Bars Colored 1", () => {
            verifyGetClassForKpi('Five Bars Colored', '-1', KpiUtil.KpiImageSize.Small, 'kpiFiveBars1');
        });

        it("Five Bars Colored 2", () => {
            verifyGetClassForKpi('Five Bars Colored', '0', KpiUtil.KpiImageSize.Small, 'kpiFiveBars2');
        });

        it("Five Bars Colored 3", () => {
            verifyGetClassForKpi('Five Bars Colored', '1', KpiUtil.KpiImageSize.Small, 'kpiFiveBars3');
        });

        it("Five Bars Colored 4", () => {
            verifyGetClassForKpi('Five Bars Colored', '2', KpiUtil.KpiImageSize.Small, 'kpiFiveBars4');
        });

        it("Five Bars Colored big image 0", () => {
            verifyGetClassForKpi('Five Bars Colored', '-2', KpiUtil.KpiImageSize.Big, 'kpiFiveBarsBig0');
        });

        it("Five Bars Colored big image 1", () => {
            verifyGetClassForKpi('Five Bars Colored', '-1', KpiUtil.KpiImageSize.Big, 'kpiFiveBarsBig1');
        });

        it("Five Bars Colored big image 2", () => {
            verifyGetClassForKpi('Five Bars Colored', '0', KpiUtil.KpiImageSize.Big, 'kpiFiveBarsBig2');
        });

        it("Five Bars Colored big image 3", () => {
            verifyGetClassForKpi('Five Bars Colored', '1', KpiUtil.KpiImageSize.Big, 'kpiFiveBarsBig3');
        });

        it("Five Bars Colored big image 4", () => {
            verifyGetClassForKpi('Five Bars Colored', '2', KpiUtil.KpiImageSize.Big, 'kpiFiveBarsBig4');
        });

        it("Five Boxes Colored 0", () => {
            verifyGetClassForKpi('Five Boxes Colored', '-2', KpiUtil.KpiImageSize.Small, 'kpiFiveBoxes0');
        });

        it("Five Boxes Colored 1", () => {
            verifyGetClassForKpi('Five Boxes Colored', '-1', KpiUtil.KpiImageSize.Small, 'kpiFiveBoxes1');
        });

        it("Five Boxes Colored 2", () => {
            verifyGetClassForKpi('Five Boxes Colored', '0', KpiUtil.KpiImageSize.Small, 'kpiFiveBoxes2');
        });

        it("Five Boxes Colored 3", () => {
            verifyGetClassForKpi('Five Boxes Colored', '1', KpiUtil.KpiImageSize.Small, 'kpiFiveBoxes3');
        });

        it("Five Boxes Colored 4", () => {
            verifyGetClassForKpi('Five Boxes Colored', '2', KpiUtil.KpiImageSize.Small, 'kpiFiveBoxes4');
        });

        it("Five Boxes Colored big image 0", () => {
            verifyGetClassForKpi('Five Boxes Colored', '-2', KpiUtil.KpiImageSize.Big, 'kpiFiveBoxesBig0');
        });

        it("Five Boxes Colored big image 1", () => {
            verifyGetClassForKpi('Five Boxes Colored', '-1', KpiUtil.KpiImageSize.Big, 'kpiFiveBoxesBig1');
        });

        it("Five Boxes Colored big image 2", () => {
            verifyGetClassForKpi('Five Boxes Colored', '0', KpiUtil.KpiImageSize.Big, 'kpiFiveBoxesBig2');
        });

        it("Five Boxes Colored big image 3", () => {
            verifyGetClassForKpi('Five Boxes Colored', '1', KpiUtil.KpiImageSize.Big, 'kpiFiveBoxesBig3');
        });

        it("Five Boxes Colored big image 4", () => {
            verifyGetClassForKpi('Five Boxes Colored', '2', KpiUtil.KpiImageSize.Big, 'kpiFiveBoxesBig4');
        });

        it("Gauge - Ascending 0", () => {
            verifyGetClassForKpi('Gauge - Ascending', '-2', KpiUtil.KpiImageSize.Small, 'kpiGauge0');
        });

        it("Gauge - Ascending 1", () => {
            verifyGetClassForKpi('Gauge - Ascending', '-1', KpiUtil.KpiImageSize.Small, 'kpiGauge1');
        });

        it("Gauge - Ascending 2", () => {
            verifyGetClassForKpi('Gauge - Ascending', '0', KpiUtil.KpiImageSize.Small, 'kpiGauge2');
        });

        it("Gauge - Ascending 3", () => {
            verifyGetClassForKpi('Gauge - Ascending', '1', KpiUtil.KpiImageSize.Small, 'kpiGauge3');
        });

        it("Gauge - Ascending 4", () => {
            verifyGetClassForKpi('Gauge - Ascending', '2', KpiUtil.KpiImageSize.Small, 'kpiGauge4');
        });

        it("Gauge - Ascending big image 0", () => {
            verifyGetClassForKpi('Gauge - Ascending', '-2', KpiUtil.KpiImageSize.Big, 'kpiGaugeBig0');
        });

        it("Gauge - Ascending big image 1", () => {
            verifyGetClassForKpi('Gauge - Ascending', '-1', KpiUtil.KpiImageSize.Big, 'kpiGaugeBig1');
        });

        it("Gauge - Ascending big image 2", () => {
            verifyGetClassForKpi('Gauge - Ascending', '0', KpiUtil.KpiImageSize.Big, 'kpiGaugeBig2');
        });

        it("Gauge - Ascending big image 3", () => {
            verifyGetClassForKpi('Gauge - Ascending', '1', KpiUtil.KpiImageSize.Big, 'kpiGaugeBig3');
        });

        it("Gauge - Ascending big image 4", () => {
            verifyGetClassForKpi('Gauge - Ascending', '2', KpiUtil.KpiImageSize.Big, 'kpiGaugeBig4');
        });

        function verifyGetClassForKpi(kpiName: string, value: string, size: KpiUtil.KpiImageSize, expectedCss: string) {
            var css = KpiUtil.getClassForKpi(kpiName, value, size);
            expect(css).toBe(expectedCss);
        }

        it("Bad KPI graphic status name should return undefined style", () => {
            var css = KpiUtil.getClassForKpi('Bad KPI Name', '1', KpiUtil.KpiImageSize.Big);
            expect(css).toBeUndefined();
        });

        it("GetMetadata three values small", () => {
            verifyGetMetadata('Shapes', '0', KpiUtil.KpiImageSize.Small, 'kpiShapes1');
        });

        it("GetMetadata three values small", () => {
            verifyGetMetadata('Shapes', '-1', KpiUtil.KpiImageSize.Big, 'kpiShapesBig0');
        });

        it("GetMetadata three values small", () => {
            verifyGetMetadata('Gauge - Ascending', '-2', KpiUtil.KpiImageSize.Small, 'kpiGauge0');
        });

        it("GetMetadata three values small", () => {
            verifyGetMetadata('Gauge - Ascending', '2', KpiUtil.KpiImageSize.Big, 'kpiGaugeBig4');
        });

        it("GetMetadata unknown kpi", () => {
            verifyGetMetadata('unknown foo', '1', KpiUtil.KpiImageSize.Small, undefined);
        });

        function verifyGetMetadata(kpiName: string, value: string, size: KpiUtil.KpiImageSize, expectedCss?: string) {
            var column = {
                kpiStatusGraphic: kpiName,
                displayName: 'test',
            };

            var metadata = KpiUtil.getKpiImageMetadata(column, value, size);

            if (expectedCss) {
                expect(metadata.caption).toBe(expectedCss);
                expect(metadata.statusGraphic).toBe(kpiName);
            }
            else {
                expect(metadata).toBeUndefined();
            }
        }
    });
}