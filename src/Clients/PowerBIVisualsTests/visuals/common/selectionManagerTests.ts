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
    import SelectionIdBuilder = powerbi.visuals.SelectionIdBuilder;
    import SelectionManager = powerbi.visuals.utility.SelectionManager;

    describe("Selection Manager tests", () => {
        let categoryA = mocks.dataViewScopeIdentity("A");
        let categoryQueryName = "categoryA";
        let categoryColumn: powerbi.DataViewCategoryColumn = {
            source: {
                queryName: categoryQueryName,
                displayName: 'testDisplayName'
            },
            identity: [categoryA],
            values: []
        };

        it("single select - select one", (done) => {
            let sm = new SelectionManager({ hostServices: powerbitests.mocks.createVisualHostServices() });
            let selectionId = SelectionIdBuilder.builder()
                .withCategory(categoryColumn, 0)
                .createSelectionId();

            expect(sm.getSelectionIds().length).toBe(0);

            sm.select(selectionId).then((ids) => {
                expect(ids.length).toBe(1);
                done();
            });
        });

        it("single select - select one then select the same", (done) => {
            let sm = new SelectionManager({ hostServices: powerbitests.mocks.createVisualHostServices() });
            let selectionId = SelectionIdBuilder.builder()
                .withCategory(categoryColumn, 0)
                .createSelectionId();

            expect(sm.getSelectionIds().length).toBe(0);

            sm.select(selectionId).then((ids) => {
                expect(ids.length).toBe(1);
                sm.select(selectionId).then((ids) => {
                    expect(ids.length).toBe(0);
                    done();
                });
            });
        });

        it("single select - select one, then another", (done) => {
            let sm = new SelectionManager({ hostServices: powerbitests.mocks.createVisualHostServices() });
            let selectionId1 = SelectionIdBuilder.builder()
                .withCategory(categoryColumn, 0)
                .createSelectionId();

            let selectionId2 = SelectionIdBuilder.builder()
                .withCategory(categoryColumn, 1)
                .createSelectionId();

            expect(sm.getSelectionIds().length).toBe(0);

            sm.select(selectionId1).then((ids) => {
                expect(ids.length).toBe(1);
                sm.select(selectionId2).then((ids) => {
                    expect(ids.length).toBe(1);
                    done();
                });
            });
        });

        it("single select - select one, then select the same", (done) => {
            let sm = new SelectionManager({ hostServices: powerbitests.mocks.createVisualHostServices() });
            let selectionId1 = SelectionIdBuilder.builder()
                .withCategory(categoryColumn, 0)
                .createSelectionId();

            expect(sm.getSelectionIds().length).toBe(0);

            sm.select(selectionId1).then((ids) => {
                expect(ids.length).toBe(1);
                sm.select(selectionId1).then((ids) => {
                    expect(ids.length).toBe(0);
                    done();
                });
            });
        });

        it("multi select - select one, then another", (done) => {
            let sm = new SelectionManager({ hostServices: powerbitests.mocks.createVisualHostServices() });
            let selectionId1 = SelectionIdBuilder.builder()
                .withCategory(categoryColumn, 0)
                .createSelectionId();

            let selectionId2 = SelectionIdBuilder.builder()
                .withCategory(categoryColumn, 1)
                .createSelectionId();

            expect(sm.getSelectionIds().length).toBe(0);

            sm.select(selectionId1, true).then((ids) => {
                expect(ids.length).toBe(1);
                sm.select(selectionId2, true).then((ids) => {
                    expect(ids.length).toBe(2);
                    done();
                });
            });
        });

        it("multi select - select one, then another & clear", (done) => {
            let sm = new SelectionManager({ hostServices: powerbitests.mocks.createVisualHostServices() });
            let selectionId1 = SelectionIdBuilder.builder()
                .withCategory(categoryColumn, 0)
                .createSelectionId();

            let selectionId2 = SelectionIdBuilder.builder()
                .withCategory(categoryColumn, 1)
                .createSelectionId();

            expect(sm.getSelectionIds().length).toBe(0);

            sm.select(selectionId1, true).then((ids) => {
                expect(ids.length).toBe(1);
                sm.select(selectionId2, true).then((ids) => {
                    expect(ids.length).toBe(2);
                    sm.clear().then(() => {
                        expect(sm.getSelectionIds().length).toBe(0);
                        done();
                    });
                });
            });
        });

        it("single select - when host in drill down", (done) => {
            var mockHostServices = powerbitests.mocks.createVisualHostServices();
            let sm = new SelectionManager({ hostServices: mockHostServices });
            let selectionId = SelectionIdBuilder.builder()
                .withCategory(categoryColumn, 0)
                .createSelectionId();

            expect(sm.getSelectionIds().length).toBe(0);

            sm.select(selectionId).then((ids) => {
                expect(ids.length).toBe(1);
                mockHostServices.shouldRetainSelection = () => true;
                sm.select(selectionId).then((ids) => {
                    expect(ids.length).toBe(1);
                    done();
                });
            });
        });
    });
}