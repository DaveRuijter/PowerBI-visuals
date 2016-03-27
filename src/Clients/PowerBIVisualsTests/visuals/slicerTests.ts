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
    import Slicer = powerbi.visuals.Slicer;
    import DataViewTransform = powerbi.data.DataViewTransform;
    import ValueType = powerbi.ValueType;
    import SelectionId = powerbi.visuals.SelectionId;

    const SelectAllTextKey = 'Select All';
    const SelectedClass = 'selected';
    const CheckedClass = 'checked';

    powerbitests.mocks.setLocale();

    var dataViewMetadata: powerbi.DataViewMetadata = {
        columns: [
            { displayName: "Fruit", properties: { "Category": true }, type: ValueType.fromDescriptor({ text: true }) },
            { displayName: "Price", isMeasure: true }]
    };

    var dataViewCategorical = {
        categories: [{
            source: dataViewMetadata.columns[0],
            values: ["Apple", "Orange", "Kiwi", "Grapes", "Banana"],
            identity: [
                mocks.dataViewScopeIdentity("Apple"),
                mocks.dataViewScopeIdentity("Orange"),
                mocks.dataViewScopeIdentity("Kiwi"),
                mocks.dataViewScopeIdentity("Grapes"),
                mocks.dataViewScopeIdentity("Banana")
            ]
        }],
        values: DataViewTransform.createValueColumns([{
            source: dataViewMetadata.columns[1],
            values: [20, 10, 30, 15, 12]
        }])
    };

    var dataView: powerbi.DataView = {
        metadata: dataViewMetadata,
        categorical: dataViewCategorical
    };

    var interactiveDataViewOptions: powerbi.VisualDataChangedOptions = {
        dataViews: [dataView]
    };

    describe("Slicer", () => {
        it("Slicer_registered_capabilities", () => {
            expect(powerbi.visuals.visualPluginFactory.create().getPlugin("slicer").capabilities).toBe(powerbi.visuals.slicerCapabilities);
        });

        it("Capabilities should include dataViewMapping", () => {
            expect(powerbi.visuals.slicerCapabilities.dataViewMappings).toBeDefined();
            expect(powerbi.visuals.slicerCapabilities.dataViewMappings.length).toBe(1);
        });

        it("Capabilities should have condition", () => {
            expect(powerbi.visuals.slicerCapabilities.dataViewMappings[0].conditions.length).toBe(1);
            expect(powerbi.visuals.slicerCapabilities.dataViewMappings[0].conditions[0][powerbi.visuals.slicerCapabilities.dataRoles[0].name].max).toBe(1);
        });

        it("Capabilities should include dataRoles", () => {
            expect(powerbi.visuals.slicerCapabilities.dataRoles).toBeDefined();
            expect(powerbi.visuals.slicerCapabilities.dataRoles.length).toBe(1);
        });

        it("Capabilities should suppressDefaultTitle", () => {
            expect(powerbi.visuals.slicerCapabilities.suppressDefaultTitle).toBe(true);
        });

        it("Filter property should match calculated", () => {
            expect(powerbi.data.DataViewObjectDescriptors.findFilterOutput(powerbi.visuals.slicerCapabilities.objects)).toEqual(powerbi.visuals.slicerProps.filterPropertyIdentifier);
        });

        it("Sort should be default so the sort UI shows", () => {
            expect(powerbi.visuals.slicerCapabilities.sorting.custom).not.toBeDefined();
            expect(powerbi.visuals.slicerCapabilities.sorting.default).toBeDefined();
        });
    });

    describe("Slicer DOM tests", () => {
        var v: Slicer;
        var element: JQuery;
        var originalRequestAnimationFrameCallback: (callback: Function) => number;

        beforeEach(() => {
            element = powerbitests.helpers.testDom("200", "300");
            v = <Slicer> powerbi.visuals.visualPluginFactory.create().getPlugin("slicer").create();

            v.init({
                element: element,
                host: mocks.createVisualHostServices(),
                style: powerbi.visuals.visualStyles.create(),
                viewport: {
                    height: element.height(),
                    width: element.width()
                }
            });
            originalRequestAnimationFrameCallback = window.requestAnimationFrame;
            window.requestAnimationFrame = (callback: () => void) => {
                callback();
                return 0;
            };
            jasmine.clock().install();
        });

        afterEach(function () {
            window.requestAnimationFrame = originalRequestAnimationFrameCallback;
            jasmine.clock().uninstall();
        });

        xit("Slicer DOM Validation", () => {
            spyOn(powerbi.visuals.valueFormatter, "format").and.callThrough();

            v.onDataChanged(interactiveDataViewOptions);
            jasmine.clock().tick(0);

            expect($(".slicerContainer")).toBeInDOM();
            expect($(".slicerContainer .headerText")).toBeInDOM();
            expect($(".slicerContainer .slicerHeader .clear")).toBeInDOM();
            expect($(".slicerContainer .slicerBody")).toBeInDOM();
            expect($(".slicerContainer .slicerBody .row .slicerText")).toBeInDOM();
            expect($(".slicerText").length).toBe(6);
            expect($(".slicerText").first().text()).toBe(SelectAllTextKey);
            expect($(".slicerText").last().text()).toBe("Banana");

            expect(powerbi.visuals.valueFormatter.format).toHaveBeenCalledWith("Apple", undefined);
            expect(powerbi.visuals.valueFormatter.format).toHaveBeenCalledWith("Orange", undefined);
            expect(powerbi.visuals.valueFormatter.format).toHaveBeenCalledWith("Kiwi", undefined);
            expect(powerbi.visuals.valueFormatter.format).toHaveBeenCalledWith("Grapes", undefined);
            expect(powerbi.visuals.valueFormatter.format).toHaveBeenCalledWith("Banana", undefined);

            // Subsequent update
            var dataView2: powerbi.DataView = {
                metadata: dataViewMetadata,
                categorical: {
                    categories: [{
                        source: dataViewMetadata.columns[0],
                        values: ["Strawberry", "Blueberry", "Blackberry"],
                        identity: [
                            mocks.dataViewScopeIdentity("Strawberry"),
                            mocks.dataViewScopeIdentity("Blueberry"),
                            mocks.dataViewScopeIdentity("Blackberry")
                        ]
                    }],
                    values: DataViewTransform.createValueColumns([{
                        source: dataViewMetadata.columns[1],
                        values: [40, 25, 22]
                    }])
                }
            };

            v.onDataChanged({ dataViews: [dataView2] });
            jasmine.clock().tick(0);

            expect($(".slicerContainer")).toBeInDOM();
            expect($(".slicerContainer .headerText")).toBeInDOM();
            expect($(".slicerContainer .slicerHeader .clear")).toBeInDOM();
            expect($(".slicerContainer .slicerBody")).toBeInDOM();
            expect($(".slicerContainer .slicerBody .row .slicerText")).toBeInDOM();

            expect($(".slicerText").length).toBe(4);
            expect($(".slicerText").first().text()).toBe(SelectAllTextKey);
            expect($(".slicerText").last().text()).toBe("Blackberry");
        });

        it("Validate converter", () => {
            jasmine.clock().tick(0);
            var slicerData = Slicer.converter(dataView, SelectAllTextKey, null);
            expect(slicerData.slicerDataPoints.length).toBe(6);
            var dataViewIdentities = dataView.categorical.categories[0].identity;
            var selectionIds = [
                SelectionId.createWithId(dataViewIdentities[0]),
                SelectionId.createWithId(dataViewIdentities[1]),
                SelectionId.createWithId(dataViewIdentities[2]),
                SelectionId.createWithId(dataViewIdentities[3]),
                SelectionId.createWithId(dataViewIdentities[4])
            ];
            var dataPoints = [
                {
                    value: SelectAllTextKey,
                    identity: SelectionId.createWithMeasure(SelectAllTextKey),
                    selected: false,
                    isSelectAllDataPoint: true
                },
                {
                    value: "Apple",
                    identity: selectionIds[0],
                    selected: false
                },
                {
                    value: "Orange",
                    identity: selectionIds[1],
                    selected: false
                },
                {
                    value: "Kiwi",
                    identity: selectionIds[2],
                    selected: false
                },
                {
                    value: "Grapes",
                    identity: selectionIds[3],
                    selected: false
                },
                {
                    value: "Banana",
                    identity: selectionIds[4],
                    selected: false
                }];

            var slicerSettings = {
                general: {
                    outlineColor: '#808080',
                    outlineWeight: 1
                },
                header: {
                    borderBottomWidth: 1,
                    show: true,
                    outline: 'BottomOnly',
                    fontColor: '#000000',
                    background: '#ffffff',
                    textSize: 10,
                },
                slicerText: {
                    color: '#666666',
                    outline: 'None',
                    background: '#ffffff',
                    textSize: 10,
                },
            };

            let expectedSlicerData = {
                categorySourceName: "Fruit",
                formatString: undefined,
                slicerSettings: slicerSettings,
                slicerDataPoints: dataPoints
            };

            expect(slicerData).toEqual(expectedSlicerData);
        });

        it("Null dataView test", () => {
            v.onDataChanged({ dataViews: [] });
            jasmine.clock().tick(0);

            expect($(".slicerText").length).toBe(0);
        });

        xit("Slicer resize", () => {
            var viewport = {
                height: 200,
                width: 300
            };
            v.onResizing(viewport);
            jasmine.clock().tick(0);

            expect($(".slicerContainer .slicerBody").first().css("height")).toBe("182px");
            expect($(".slicerContainer .slicerBody").first().css("width")).toBe("300px");
            expect($(".slicerContainer .headerText").first().css("width")).toBe("275px");

            // Next Resize
            var viewport2 = {
                height: 150,
                width: 150
            };
            v.onResizing(viewport2);
            jasmine.clock().tick(0);

            expect($(".slicerContainer .slicerBody").first().css("height")).toBe("132px");
            expect($(".slicerContainer .slicerBody").first().css("width")).toBe("150px");
        });

    });

    describe("Slicer Interactivity", () => {
        var v: powerbi.IVisual;
        var element: JQuery;
        var slicerText: JQuery;
        var slicerCheckbox: JQuery;
        var slicerCheckboxInput: JQuery;
        var hostServices: powerbi.IVisualHostServices;
        var clearSelectionSpy: jasmine.Spy;
        var originalRequestAnimationFrameCallback: (callback: Function) => number;

        beforeEach(() => {

            var originalFunc = powerbi.visuals.createInteractivityService;
            powerbi.visuals.createInteractivityService = (host) => {
                var result = originalFunc(host);
                clearSelectionSpy = spyOn(result, "clearSelection");
                clearSelectionSpy.and.callThrough();
                powerbi.visuals.createInteractivityService = originalFunc;
                return result;
            };

            element = powerbitests.helpers.testDom("200", "300");
            v = <Slicer> powerbi.visuals.visualPluginFactory.createMinerva({ dataDotChartEnabled: false, heatMap: false}).getPlugin("slicer").create();
            hostServices = mocks.createVisualHostServices();
            hostServices.canSelect = () => true;

            v.init({
                element: element,
                host: hostServices,
                style: powerbi.visuals.visualStyles.create(),
                viewport: {
                    height: element.height(),
                    width: element.width()
                },
                interactivity: { selection: true }
            });
            originalRequestAnimationFrameCallback = window.requestAnimationFrame;
            window.requestAnimationFrame = (callback: () => void) => {
                callback();
                return 0;
            };
            jasmine.clock().install();

            v.onDataChanged(interactiveDataViewOptions);
            jasmine.clock().tick(0);

            slicerText = $(".slicerText");
            slicerCheckbox = $(".slicerCheckbox");
            slicerCheckboxInput = $(".slicerCheckbox").find("input");

            spyOn(hostServices, "onSelect").and.callThrough();
        });

        afterEach(function () {
            window.requestAnimationFrame = originalRequestAnimationFrameCallback;
            jasmine.clock().uninstall();
        });

        describe("slicer item select", () => {
            xit("by text", () => {
                jasmine.clock().tick(0);
                (<any>slicerText.eq(1)).d3Click(0, 0);

                expect(slicerCheckbox[1].classList).toContain(SelectedClass);
                expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(true);
                expect(slicerCheckbox[2].classList).not.toContain(SelectedClass);
                expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(false);

                expect(hostServices.onSelect).toHaveBeenCalledWith({
                    data:
                    [
                        {
                            data: [
                                interactiveDataViewOptions.dataViews[0].categorical.categories[0].identity[0]
                            ]
                        }
                    ]
                });
            });

            xit("by checkbox", () => {
                jasmine.clock().tick(0);
                (<any>slicerCheckbox.eq(1)).d3Click(0, 0);

                expect(slicerCheckbox[1].classList).toContain(SelectedClass);
                expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(true);
                expect(slicerCheckbox[2].classList).not.toContain(SelectedClass);
                expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(false);

                expect(hostServices.onSelect).toHaveBeenCalledWith({
                    data:
                    [
                        {
                            data: [
                                interactiveDataViewOptions.dataViews[0].categorical.categories[0].identity[0]
                            ]
                        }
                    ]
                });
            });
        });

        xit("slicer item multi-select checkboxes", () => {
            jasmine.clock().tick(0);
            (<any>slicerText.eq(1)).d3Click(0, 0);

            expect(slicerCheckbox[1].classList).toContain(SelectedClass);
            expect(slicerCheckbox[2].classList).not.toContain(SelectedClass);
            expect(d3.select(slicerCheckboxInput[0]).property(CheckedClass)).toBe(false);
            expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(true);
            expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(false);

            (<any>slicerText.last()).d3Click(0, 0);
            expect(slicerCheckbox[5].classList).toContain(SelectedClass);
            expect(d3.select(slicerCheckboxInput[5]).property(CheckedClass)).toBe(true);
        });

        xit("slicer item repeated selection", () => {
            jasmine.clock().tick(0);
            (<any>slicerText.eq(1)).d3Click(0, 0);

            expect(slicerCheckbox[1].classList).toContain(SelectedClass);

            (<any>slicerText.last()).d3Click(0, 0);
            (<any>slicerText.last()).d3Click(0, 0);
                
            expect(slicerCheckbox[5].classList).not.toContain(SelectedClass);

            expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(true);
            expect(d3.select(slicerCheckboxInput[5]).property(CheckedClass)).toBe(false);
        });

        xit("slicer selectAll", () => {
            jasmine.clock().tick(0);

            expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(false);
            expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(false);

            (<any>slicerText.eq(0)).d3Click(0, 0);
            expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(true);
            expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(true);
            expect(d3.select(slicerCheckboxInput[5]).property(CheckedClass)).toBe(true);

            (<any>slicerText.eq(1)).d3Click(0, 0);
            expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(false);
            expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(true);
            expect(d3.select(slicerCheckboxInput[5]).property(CheckedClass)).toBe(true);
        });

        xit("slicer partial select", () => {
            jasmine.clock().tick(0);

            expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(false);
            expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(false);

            (<any>slicerText.eq(0)).d3Click(0, 0);
            expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(true);
            expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(true);
            expect(d3.select(slicerCheckboxInput[5]).property(CheckedClass)).toBe(true);
            var partialSelect = $(".partiallySelected");
            expect(partialSelect.length).toBe(0);

            (<any>slicerText.eq(1)).d3Click(0, 0);
            partialSelect = $(".partiallySelected");
            expect(partialSelect.length).toBe(1);
            expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(false);
            expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(true);
            expect(d3.select(slicerCheckboxInput[5]).property(CheckedClass)).toBe(true);

            (<any>slicerText.eq(0)).d3Click(0, 0);
            expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(false);
            expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(false);
            expect(d3.select(slicerCheckboxInput[5]).property(CheckedClass)).toBe(false);
            var partialSelect = $(".partiallySelected");
            expect(partialSelect.length).toBe(0);

            (<any>slicerText.eq(1)).d3Click(0, 0);
            partialSelect = $(".partiallySelected");
            expect(partialSelect.length).toBe(1);
            expect(d3.select(slicerCheckboxInput[1]).property(CheckedClass)).toBe(true);
            expect(d3.select(slicerCheckboxInput[2]).property(CheckedClass)).toBe(false);
            expect(d3.select(slicerCheckboxInput[5]).property(CheckedClass)).toBe(false);
        });

        xit("slicer clear", () => {
            jasmine.clock().tick(0);
            var clearBtn = $(".clear");

            // Slicer click
            (<any>slicerText.eq(1)).d3Click(0, 0);
            expect(slicerCheckbox[1].classList).toContain(SelectedClass);
            expect(slicerCheckbox[2].classList).not.toContain(SelectedClass);

            (<any>slicerText.last()).d3Click(0, 0);
            expect(slicerCheckbox[5].classList).toContain(SelectedClass);

            /* Slicer clear */
            (<any>clearBtn.first()).d3Click(0, 0);

            expect(slicerCheckbox[0].classList).not.toContain(SelectedClass);
            expect(slicerCheckbox[1].classList).not.toContain(SelectedClass);
            expect(slicerCheckbox[2].classList).not.toContain(SelectedClass);
            expect(slicerCheckbox[3].classList).not.toContain(SelectedClass);
            expect(slicerCheckbox[4].classList).not.toContain(SelectedClass);
            expect(slicerCheckbox[5].classList).not.toContain(SelectedClass);

            expect(hostServices.onSelect).toHaveBeenCalledWith({ data: [] });
        });

        it("slicer loadMoreData noSegment", () => {
            var listViewOptions: powerbi.visuals.ListViewOptions = <powerbi.visuals.ListViewOptions>v["listView"]["options"];
            var loadMoreSpy = spyOn(hostServices, "loadMoreData");
            listViewOptions.loadMoreData();
            expect(loadMoreSpy).not.toHaveBeenCalled();
        });

        it("slicer loadMoreData", () => {
            var metadata: powerbi.DataViewMetadata = {
                columns: dataViewMetadata.columns,
                segment: {}
            };

            var interactiveDataViewOptions: powerbi.VisualDataChangedOptions = {
                dataViews: [{ metadata: metadata, categorical: dataViewCategorical }]
            };
            v.onDataChanged(interactiveDataViewOptions);

            var listViewOptions: powerbi.visuals.ListViewOptions = <powerbi.visuals.ListViewOptions>v["listView"]["options"];
            var loadMoreSpy = spyOn(hostServices, "loadMoreData");
            listViewOptions.loadMoreData();
            expect(loadMoreSpy).toHaveBeenCalled();
        });

        it("slicer loadMoreData already called", () => {
            var metadata: powerbi.DataViewMetadata = {
                columns: dataViewMetadata.columns,
                segment: {}
            };

            var interactiveDataViewOptions: powerbi.VisualDataChangedOptions = {
                dataViews: [{ metadata: metadata, categorical: dataViewCategorical }]
            };
            v.onDataChanged(interactiveDataViewOptions);

            var listViewOptions: powerbi.visuals.ListViewOptions = <powerbi.visuals.ListViewOptions>v["listView"]["options"];
            var loadMoreSpy = spyOn(hostServices, "loadMoreData");
            listViewOptions.loadMoreData();
            listViewOptions.loadMoreData();
            expect(loadMoreSpy.calls.all().length).toBe(1);
        });

        it("Validate scroll position on onDataChanged", () => {
            var interactiveDataViewOptionsWithLoadMore: powerbi.VisualDataChangedOptions = {
                dataViews: [{ metadata: dataViewMetadata, categorical: dataViewCategorical }],
                operationKind: powerbi.VisualDataChangeOperationKind.Append
            };

            dataViewCategorical = {
                categories: [{
                    source: dataViewMetadata.columns[0],
                    values: ["PineApple", "Strawberry", "Mango", "Grapes", "Banana"],
                    identity: [
                        mocks.dataViewScopeIdentity("PineApple"),
                        mocks.dataViewScopeIdentity("Strawberry"),
                        mocks.dataViewScopeIdentity("Mango"),
                        mocks.dataViewScopeIdentity("Grapes"),
                        mocks.dataViewScopeIdentity("Banana"),
                    ]
                }],
                values: DataViewTransform.createValueColumns([{
                    source: dataViewMetadata.columns[1],
                    values: [20, 10, 30, 15, 12]
                }]),
            };
            var interactiveDataViewOptionWithCreate: powerbi.VisualDataChangedOptions = {
                dataViews: [{ metadata: dataViewMetadata, categorical: dataViewCategorical }],
                operationKind: powerbi.VisualDataChangeOperationKind.Create
            };

            var listView = <powerbi.visuals.IListView>v["listView"];
            var renderSpy = spyOn(listView, "render");

            v.onDataChanged(interactiveDataViewOptions);

            jasmine.clock().tick(0);               

            // Loading the same categories should NOT reset the scrollbar
            expect(renderSpy).toHaveBeenCalled();

                // LoadMore should NOT reset the scrollbar
            v.onDataChanged(interactiveDataViewOptionsWithLoadMore);
            jasmine.clock().tick(0); 
            expect(renderSpy).toHaveBeenCalled();

            // OperationKind of create and data with different category identity should reset the scrollbar position
            v.onDataChanged(interactiveDataViewOptionWithCreate);
            jasmine.clock().tick(0);
            expect(renderSpy).toHaveBeenCalled();
        });

        xit('show hide header test', () => {
            jasmine.clock().tick(0);

            expect($(".slicerHeader").css('display')).toBe('block');

            dataView.metadata.objects = { header: { show: false } };
            v.onDataChanged({
                dataViews: [dataView]
            });
            jasmine.clock().tick(0);

            expect($(".slicerHeader").css('display')).toBe('none');
        });

        it('Header outline color test', () => {
            jasmine.clock().tick(0);
            expect($(".headerText").css('border-color')).toBe('rgb(128, 128, 128)');
        });

        xit('background and font slicer text test', () => {
            jasmine.clock().tick(0);

            expect($(".slicerText").css('color')).toBe('rgb(102, 102, 102)');

            dataView.metadata.objects = {
                Rows: {
                    fontColor: { solid: { color: '#f5f5f5' } },
                    background: { solid: { color: '#f6f6f6' } },
                }
            };
            v.onDataChanged({
                dataViews: [dataView]
            });
            jasmine.clock().tick(0);

            expect($(".slicerText").css('color')).toBe('rgb(245, 245, 245)');
            expect($(".slicerText").css('background-color')).toBe('rgb(246, 246, 246)');
        });

        xit('background and font header test', () => {
            jasmine.clock().tick(0);

            expect($(".slicerHeader .headerText").css('color')).toBe('rgb(0, 0, 0)');

            dataView.metadata.objects = {
                header: {
                    show: true,
                    fontColor: { solid: { color: '#f5f5f5' } },
                    background: { solid: { color: '#f6f6f6' } },
                }
            };
            v.onDataChanged({
                dataViews: [dataView]
            });
            jasmine.clock().tick(0);

            expect($(".slicerHeader .headerText").css('color')).toBe('rgb(245, 245, 245)');
            expect($(".slicerHeader .headerText").css('background-color')).toBe('rgb(246, 246, 246)');
        });

        xit('test header border outline', () => {
           jasmine.clock().tick(0);

            expect($(".headerText").css('border-width')).toBe('0px 0px 1px');

            dataView.metadata.objects = { header: { outline: 'None' } };
            v.onDataChanged({
                dataViews: [dataView]
            });
            jasmine.clock().tick(0);

            expect($(".headerText").css('border-width')).toBe('0px');

            dataView.metadata.objects = { header: { outline: 'TopOnly' } };
            v.onDataChanged({
                dataViews: [dataView]
            });
            jasmine.clock().tick(0);

            expect($(".headerText").css('border-width')).toBe('1px 0px 0px');

            dataView.metadata.objects = { header: { outline: 'TopBottom' } };
            v.onDataChanged({
                dataViews: [dataView]
            });
            jasmine.clock().tick(0);

            expect($(".headerText").css('border-width')).toBe('1px 0px');

            dataView.metadata.objects = { header: { outline: 'LeftRight' } };
            v.onDataChanged({
                dataViews: [dataView]
            });
            jasmine.clock().tick(0);

            expect($(".headerText").css('border-width')).toBe('0px 1px');

            dataView.metadata.objects = { header: { outline: 'Frame' } };
            v.onDataChanged({
                dataViews: [dataView]
            });
            jasmine.clock().tick(0);

            expect($(".headerText").css('border-width')).toBe('1px');
        });

        xit('row text size', () => {
            jasmine.clock().tick(0);

            expect($(".slicerText").css('font-size')).toBe('13px');

            dataView.metadata.objects = {
                Rows: {
                    textSize: 14,
                }
            };
            v.onDataChanged({
                dataViews: [dataView]
            });
            jasmine.clock().tick(0);

            expect($(".slicerText").css('font-size')).toBe('19px');
        });

        xit('header text size', () => {
            jasmine.clock().tick(0);

            expect($(".slicerHeader .headerText").css('font-size')).toBe('13px');

            dataView.metadata.objects = {
                header: {
                    show: true,
                    textSize: 14,
                }
            };
            v.onDataChanged({
                dataViews: [dataView]
            });
            jasmine.clock().tick(0);

            expect($(".slicerHeader .headerText").css('font-size')).toBe('19px');
        });
    });
}