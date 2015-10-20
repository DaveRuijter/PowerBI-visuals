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

module powerbitests {
    import ColorConvertor = powerbitests.utils.ColorUtility.convertFromRGBorHexToHex;
    import BasicShapeVisual = powerbi.visuals.BasicShapeVisual;
    import basicShapeCapabilities = powerbi.visuals.basicShapeCapabilities;

    module BasicShapeHelpers {
        const RotateTransformStyleRegex = /transform: rotate\((\d+)deg\)/g;

        export function getRotateFromStyle(rotatedElement: JQuery): number {
            let style = rotatedElement.attr('style');
            let rotate = RotateTransformStyleRegex.exec(style);
            return parseInt(rotate[1], 10);
        };

        export function colorMatch(actualColor: string, expectedColor: string): void {
            let actualColorHex = ColorConvertor(actualColor).toUpperCase();
            let expectedColorHex = ColorConvertor(expectedColor).toUpperCase();

            expect(actualColorHex).toBe(expectedColorHex);
        };

        export function buildUpdateOptions(viewport: powerbi.IViewport, objects: powerbi.DataViewObjects): powerbi.VisualUpdateOptions {
            return {
                viewport: viewport,
                dataViews: [{
                    metadata: {
                        columns: [],
                        objects: objects,
                    }
                }],
            };
        };
    }

    describe("basicShape Tests", () => {

        it('registered capabilities', () => {
            var pluginCapabilities = powerbi.visuals.visualPluginFactory.create().getPlugin('basicShape').capabilities;
            expect(pluginCapabilities.toString()).toBe(basicShapeCapabilities.toString());
        });

        it("capabilities should suppressDefaultTitle", () => {
            expect(basicShapeCapabilities.suppressDefaultTitle).toBe(true);
        });

        it('registered capabilities: objects', () => {
            expect(powerbi.visuals.visualPluginFactory.create().getPlugin('basicShape').capabilities.objects).toBeDefined();
        });

        describe('rendering', () => {
            var element: JQuery;
            var viewport: powerbi.IViewport;
            var options: powerbi.VisualInitOptions;
            var basicShape: BasicShapeVisual;

            beforeEach(() => {
                element = powerbitests.helpers.testDom('200', '300');
                viewport = {
                    height: element.height(),
                    width: element.width()
                };
                options = {
                    element: $(element[0]),
                    host: mocks.createVisualHostServices(),
                    style: powerbi.visuals.visualStyles.create(),
                    viewport: viewport,
                    animation: {
                        transitionImmediate: true
                    },
                };

                basicShape = new BasicShapeVisual();
                basicShape.init(options);
            });

            it('no visual configuration', () => {
                expect(element.children().length).toBe(0);
            });

            it('default shape', () => {
                let visualUpdateOptions = BasicShapeHelpers.buildUpdateOptions(viewport, {});
                basicShape.update(visualUpdateOptions);

                //Verifying the DOM
                var rect = element.find('rect');

                BasicShapeHelpers.colorMatch(rect.css('stroke'), BasicShapeVisual.DefaultStrokeColor); // lineColor
                BasicShapeHelpers.colorMatch(rect.css('fill'), BasicShapeVisual.DefaultFillColor); // fillColor
                expect(rect.css('fill-opacity')).toBe((BasicShapeVisual.DefaultFillTransValue / 100).toString()); // shapeTransparency
            });

            it('rect', () => {
                let visualUpdateOptions = BasicShapeHelpers.buildUpdateOptions(viewport, {
                    general: { shapeType: 'rectangle' },
                    line: { lineColor: { solid: { color: '#00b8ad' } }, transparency: 75, weight: 15 },
                    fill: { transparency: 65, fillColor: { solid: { color: '#e6e6e4' } } },
                });
                basicShape.update(visualUpdateOptions);

                //Verifying the DOM
                var rect = element.find('rect');
                BasicShapeHelpers.colorMatch(rect.css('stroke'), "#00b8ad"); // lineColor
                BasicShapeHelpers.colorMatch(rect.css('fill'), "#e6e6e4"); // fillColor
                expect(rect.css('stroke-opacity')).toBe("0.75"); // lineTransparency
                expect(rect.css('stroke-width')).toBe("15px"); // weight
                expect(rect.css('fill-opacity')).toBeCloseTo("0.65", 1); // shapeTransparency
            });

            /** 
              * In Chrome and IE11, this test passes
              * Running by command line via PhantomJS, this test fails
              * In PhantomJS, for some reason, the style on the parent is not being set
              * so the transform never happens and cannot be verified.
              * http://sqlbuvsts01:8080/Main/SQL%20Server/_workitems/edit/5701900
              */
            //it('rect with rotation', () => {
            //    let rotation = 270;

            //    let visualUpdateOptions = BasicShapeHelpers.buildUpdateOptions(viewport, {
            //        rotation: { angle: rotation },
            //    });
            //    basicShape.update(visualUpdateOptions);

            //    //Verifying the DOM
            //    var rect = element.find('rect');
            //    var rotator = rect.parents('div[style*=transform]').eq(0);
            //    expect(BasicShapeHelpers.getRotateFromStyle(rotator)).toBe(270);
            //});
        });
    });
}
