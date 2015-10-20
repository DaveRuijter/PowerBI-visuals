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
module powerbi.visuals {
    export module shapeFactory {

        export module ShapeFactoryConsts {
            export const PaddingConstRatio: number = 0.01;
            export const ShapeConstRatio: number = 1.0 - (ShapeFactoryConsts.PaddingConstRatio * 2);
            export const SmallPaddingConstValue: number = 10;
            export const OvalRadiusConst: number = 2;
            export const OvalRadiusConstPadding: number = 0.2;
        }

        export function createRectangle(data: BasicShapeData, viewportHeight: number, viewportWidth: number, selectedElement: D3.Selection, degrees: number): void {
            let x = viewportWidth * ShapeFactoryConsts.PaddingConstRatio;
            let y = viewportHeight * ShapeFactoryConsts.PaddingConstRatio;
            let width = viewportWidth * ShapeFactoryConsts.ShapeConstRatio;
            let height = viewportHeight * ShapeFactoryConsts.ShapeConstRatio;
            let scale = getScale(width, height, degrees);

            // create the inner path with the wanted shape
            selectedElement
                .append('div')
                .style({
                    'transform': 'rotate(' + degrees + 'deg) scale(' + scale + ')',
                    'transform-origin': 'center'
                })
                .append('svg')
                .attr({
                    width: viewportWidth,
                    height: viewportHeight
                })
                .append('rect')
                .attr({
                    x: x,
                    y: y,
                    width: width,
                    height: height,
                    rx: data.roundEdge,
                    ry: data.roundEdge
                })
                .style({
                    'vector-effect': 'non-scaling-stroke',
                    'stroke-width': data.lineWeight + 'px',
                    'stroke': data.lineColor,
                    'stroke-opacity': data.lineTransparency / 100,
                    'fill': data.fillColor,
                    'fill-opacity': data.showFill === true ? data.shapeTransparency / 100 : 0
                });
        }

        /** this function creates a oval svg   */
        export function createOval(data: BasicShapeData, viewportHeight: number, viewportWidth: number, selectedElement: D3.Selection, degrees: number): void {
            let widthForCircle = (viewportWidth / ShapeFactoryConsts.OvalRadiusConst).toString();
            let heightForCircle = (viewportHeight / ShapeFactoryConsts.OvalRadiusConst).toString();
            let radiusXForCircle = (viewportWidth / (ShapeFactoryConsts.OvalRadiusConst + ShapeFactoryConsts.OvalRadiusConstPadding));
            let radiusYForCircle = (viewportHeight / (ShapeFactoryConsts.OvalRadiusConst + ShapeFactoryConsts.OvalRadiusConstPadding));

            let scale = getScale(viewportWidth, viewportHeight, degrees);

            selectedElement
                .append('div')
                .style({
                    'transform': 'rotate(' + degrees + 'deg) scale(' + scale + ')',
                    'transform-origin': 'center'
                })
                .append('svg')
                .attr({
                    width: viewportWidth,
                    height: viewportHeight
                })
                .append('ellipse')
                .attr({
                    cx: widthForCircle,
                    cy: heightForCircle,
                    rx: radiusXForCircle,
                    ry: radiusYForCircle
                })
                .style({
                    'vector-effect': 'non-scaling-stroke',
                    'stroke-width': data.lineWeight + 'px',
                    'stroke': data.lineColor,
                    'fill': data.fillColor,
                    'fill-opacity': data.showFill === true ? data.shapeTransparency / 100 : 0
                });
        }

        /** this function creates a line svg   */
        export function createLine(data: BasicShapeData, viewportHeight: number, viewportWidth: number, selectedElement: D3.Selection, degrees: number): void {
            let x1, y1, x2, y2;
            let width = (viewportWidth - ShapeFactoryConsts.SmallPaddingConstValue) - ShapeFactoryConsts.SmallPaddingConstValue;
            let height = (viewportHeight - ShapeFactoryConsts.SmallPaddingConstValue) - ShapeFactoryConsts.SmallPaddingConstValue;

            let ratio;
            if (degrees <= 45) {
                ratio = degrees / 90;

                x1 = viewportWidth / 2 + width * ratio;
                y1 = ShapeFactoryConsts.SmallPaddingConstValue;
                x2 = viewportWidth / 2 - width * ratio;
                y2 = (viewportHeight - ShapeFactoryConsts.SmallPaddingConstValue);
            } else if (degrees <= 135) {
                ratio = (degrees - 45) / 90;

                x1 = (viewportWidth - ShapeFactoryConsts.SmallPaddingConstValue);
                y1 = ShapeFactoryConsts.SmallPaddingConstValue + height * ratio;
                x2 = ShapeFactoryConsts.SmallPaddingConstValue;
                y2 = (viewportHeight - ShapeFactoryConsts.SmallPaddingConstValue) - height * ratio;
            } else if (degrees <= 225) {
                ratio = (degrees - 135) / 90;

                x1 = (viewportWidth - ShapeFactoryConsts.SmallPaddingConstValue) - width * ratio;
                y1 = (viewportHeight - ShapeFactoryConsts.SmallPaddingConstValue);
                x2 = ShapeFactoryConsts.SmallPaddingConstValue + width * ratio;
                y2 = ShapeFactoryConsts.SmallPaddingConstValue;
            } else if (degrees <= 315) {
                ratio = (degrees - 225) / 90;

                x1 = ShapeFactoryConsts.SmallPaddingConstValue;
                y1 = (viewportHeight - ShapeFactoryConsts.SmallPaddingConstValue) - height * ratio;
                x2 = (viewportWidth - ShapeFactoryConsts.SmallPaddingConstValue);
                y2 = ShapeFactoryConsts.SmallPaddingConstValue + height * ratio;
            } else if (degrees <= 360) {
                ratio = (degrees - 315) / 90;

                x1 = ShapeFactoryConsts.SmallPaddingConstValue + width * ratio;
                y1 = ShapeFactoryConsts.SmallPaddingConstValue;
                x2 = (viewportWidth - ShapeFactoryConsts.SmallPaddingConstValue) - width * ratio;
                y2 = (viewportHeight - ShapeFactoryConsts.SmallPaddingConstValue);
            }

            // create the inner path with the wanted shape
            selectedElement
                .append('svg')
                .attr({
                    width: viewportWidth,
                    height: viewportHeight
                })
                .append('line')
                .attr({
                    x1: x1,
                    y1: y1,
                    x2: x2,
                    y2: y2,
                })
                .style({
                    'vector-effect': 'non-scaling-stroke',
                    'stroke-width': data.lineWeight + 'px',
                    'stroke-opacity': data.lineTransparency / 100,
                    'stroke': data.lineColor
                });
        }

        /** this function creates a arrow svg   */
        export function createUpArrow(data: BasicShapeData, viewportHeight: number, viewportWidth: number, selectedElement: D3.Selection, degrees: number): void {
            let lineData = [
                { 'x': (viewportWidth * 0.050).toString(), 'y': (viewportHeight * 0.420).toString() },
                { 'x': (viewportWidth * 0.500).toString(), 'y': (viewportHeight * 0.016).toString() },
                { 'x': (viewportWidth * 0.950).toString(), 'y': (viewportHeight * 0.420).toString() },
                { 'x': (viewportWidth * 0.764).toString(), 'y': (viewportHeight * 0.420).toString() },
                { 'x': (viewportWidth * 0.764).toString(), 'y': (viewportHeight * 0.993).toString() },
                { 'x': (viewportWidth * 0.246).toString(), 'y': (viewportHeight * 0.993).toString() },
                { 'x': (viewportWidth * 0.246).toString(), 'y': (viewportHeight * 0.420).toString() },
            ];

            // create the inner path with the wanted shape
            createPathFromArray(data, lineData, selectedElement, viewportHeight, viewportWidth, degrees);
        }

        /** this function creates a triangle svg   */
        export function createTriangle(data: BasicShapeData, viewportHeight: number, viewportWidth: number, selectedElement: D3.Selection, degrees: number): void {
            let lineData = [
                { 'x': '10', 'y': (viewportHeight - ShapeFactoryConsts.SmallPaddingConstValue).toString() },
                { 'x': (viewportWidth / 2).toString(), 'y': '10' },
                {
                    'x': (viewportWidth - ShapeFactoryConsts.SmallPaddingConstValue).toString(),
                    'y': (viewportHeight - ShapeFactoryConsts.SmallPaddingConstValue).toString(),
                },
            ];

            createPathFromArray(data, lineData, selectedElement, viewportHeight, viewportWidth, degrees);
        }

        /** this funcion adds a path to an svg element from an array of points (x,y) */
        function createPathFromArray(data: BasicShapeData, points: Object[], selectedElement: D3.Selection, viewportHeight: number, viewportWidth: number, degrees: number): void {

            let scale = getScale(viewportWidth, viewportHeight, degrees);

            let lineFunction = d3.svg.line()
                .x(function (d) { return d.x; })
                .y(function (d) { return d.y; })
                .interpolate('linear');

            selectedElement
                .append('div')
                .style({
                    'transform': 'rotate(' + degrees + 'deg) scale(' + scale + ')',
                    'transform-origin': 'center'
                })
                .append('svg')
                .attr({
                    width: viewportWidth,
                    height: viewportHeight
                })
                .append('path').attr({
                    d: lineFunction(points) + ' Z',
                })
                .style({
                    'vector-effect': 'non-scaling-stroke',
                    'stroke-width': data.lineWeight + 'px',
                    'stroke': data.lineColor,
                    'fill': data.fillColor,
                    'fill-opacity': data.showFill === true ? data.shapeTransparency / 100 : 0,
                    'stroke-opacity': data.lineTransparency / 100
                });
        }

        function getScale(width: number, height: number, degrees: number): number {
            let originalWidth = width;
            let originalHeight = height;
            let offsetAngle = Math.atan(width / height);
            let originalFactor = Math.sqrt(Math.pow(height, 2) + Math.pow(width, 2));
            let radians = (degrees / 180) * Math.PI;

            if (width >= height) {
                let sum;

                if (degrees < 90) {
                    sum = radians - offsetAngle;
                } else if (degrees < 180) {
                    sum = radians + offsetAngle;
                } else if (degrees < 270) {
                    sum = radians - offsetAngle;
                } else {
                    sum = radians + offsetAngle;
                }

                return (originalHeight / Math.abs(Math.cos(sum))) / originalFactor;
            } else {
                let sum;

                if (degrees < 90) {
                    sum = offsetAngle + radians;
                } else if (degrees < 180) {
                    sum = radians - offsetAngle;
                } else if (degrees < 270) {
                    sum = offsetAngle + radians;
                } else {
                    sum = radians - offsetAngle;
                }

                return (originalWidth / Math.abs(Math.sin(sum))) / originalFactor;
            }
        }
    }
}