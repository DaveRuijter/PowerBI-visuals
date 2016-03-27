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

/// <reference path="../../_references.ts"/>

module powerbitests {
    import VisibilityUtil = powerbi.visuals.visibilityHelper;

    describe("VisibilityUtil tests", () => {
        it("partiallyVisible test", () => {
            let element: JQuery;
            element = powerbitests.helpers.testDom("500", "500");
            element["visible"] = () => { return true; };

            // Element within bounds
            element.css("width", 500);
            expect(VisibilityUtil.partiallyVisible(element)).toBe(true);
            
            // Opacity 0, but non zero height and width
            element.css({ opacity: 0 });
            expect(VisibilityUtil.partiallyVisible(element)).toBe(true);
            
            // Visibility hidden but non-zero height and width
            element.css("visibility", "hidden");
            expect(VisibilityUtil.partiallyVisible(element)).toBe(true);

            // Element out of bounds
            element.css("width", 1500);
            expect(VisibilityUtil.partiallyVisible(element)).toBe(true);

            // Zero height and width
            element.css("width", 0);
            element.css("height", 0);
            expect(VisibilityUtil.partiallyVisible(element)).toBe(false);
        });
    });
}