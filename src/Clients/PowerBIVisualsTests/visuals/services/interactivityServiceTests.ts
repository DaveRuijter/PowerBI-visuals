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
    import SelectableDataPoint = powerbi.visuals.SelectableDataPoint;
    import SelectionId = powerbi.visuals.SelectionId;
    import ISelectionHandler = powerbi.visuals.ISelectionHandler;

    describe('Interactivity service', () => {
        var host: powerbi.IVisualHostServices;
        var interactivityService: powerbi.visuals.IInteractivityService;
        var selectableDataPoints: SelectableDataPoint[];
        var behavior: MockBehavior;
        var filterPropertyId: powerbi.DataViewObjectPropertyIdentifier;

        beforeEach(() => {
            host = powerbitests.mocks.createVisualHostServices();
            host.canSelect = () => true; // Allows for multiselect behavior by default
            interactivityService = powerbi.visuals.createInteractivityService(host);
            selectableDataPoints = <SelectableDataPoint[]> [
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("0"), mocks.dataViewScopeIdentity("a"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("0"), mocks.dataViewScopeIdentity("b"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("1"), mocks.dataViewScopeIdentity("a"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("1"), mocks.dataViewScopeIdentity("b"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("2"), mocks.dataViewScopeIdentity("a"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("2"), mocks.dataViewScopeIdentity("b"), "queryName") },
            ];
            filterPropertyId = {
                objectName: 'general',
                propertyName: 'selected'
            };
            behavior = new MockBehavior(selectableDataPoints, filterPropertyId);
        });

        it('Basic binding', () => {
            spyOn(behavior, "bindEvents");
            spyOn(behavior, "renderSelection");
            interactivityService.bind(selectableDataPoints, behavior, null);
            expect(behavior.bindEvents).toHaveBeenCalled();
            expect(behavior.verifyCleared()).toBeTruthy();
            expect(behavior.renderSelection).not.toHaveBeenCalled();
            expect(interactivityService.hasSelection()).toBeFalsy();
        });

        it('Binding passes behaviorOptions', () => {
            spyOn(behavior, "bindEvents");
            var arbitraryBehaviorOptions = {
                some: "random",
                collection: "of",
                random: "stuff",
            };
            interactivityService.bind(selectableDataPoints, behavior, arbitraryBehaviorOptions);
            expect(behavior.bindEvents).toHaveBeenCalledWith(arbitraryBehaviorOptions, interactivityService);
        });

        it('Basic selection', () => {
            spyOn(behavior, "renderSelection");
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(0, false);
            expect(behavior.verifySingleSelectedAt(0)).toBeTruthy();
            expect(behavior.renderSelection).toHaveBeenCalledWith(true);
            expect(interactivityService.hasSelection()).toBeTruthy();
        });

        it('Apply selection', () => {
            var newDataPoints = <SelectableDataPoint[]>[
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("0"), mocks.dataViewScopeIdentity("a"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("0"), mocks.dataViewScopeIdentity("b"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("1"), mocks.dataViewScopeIdentity("a"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("1"), mocks.dataViewScopeIdentity("b"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("2"), mocks.dataViewScopeIdentity("a"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdsAndMeasure(mocks.dataViewScopeIdentity("2"), mocks.dataViewScopeIdentity("b"), "queryName") },
            ];
            spyOn(behavior, "renderSelection");
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(0, false);
            expect(behavior.verifySingleSelectedAt(0)).toBeTruthy();
            expect(behavior.renderSelection).toHaveBeenCalledWith(true);
            interactivityService.applySelectionStateToData(newDataPoints);
            expect(newDataPoints[0].selected).toBeTruthy();
            expect(newDataPoints[1].selected).toBeFalsy();
            expect(newDataPoints[2].selected).toBeFalsy();
            expect(newDataPoints[3].selected).toBeFalsy();
            expect(newDataPoints[4].selected).toBeFalsy();
            expect(newDataPoints[5].selected).toBeFalsy();
        });

        it('Clear selection through event', () => {
            spyOn(behavior, "renderSelection");
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(0, false);
            behavior.clear();
            expect(behavior.verifyCleared()).toBeTruthy();
            expect(behavior.renderSelection).toHaveBeenCalledWith(false);
            expect(interactivityService.hasSelection()).toBeFalsy();
        });

        it('Clear selection through service', () => {
            spyOn(behavior, "renderSelection");
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(0, false);
            interactivityService.clearSelection();
            expect(behavior.verifyCleared()).toBeTruthy();
            expect(behavior.renderSelection).toHaveBeenCalledWith(false);
            expect(interactivityService.hasSelection()).toBeFalsy();
        });

        it('Selection sent to host', () => {
            spyOn(host, "onSelect");
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(0, false);
            expect(host.onSelect).toHaveBeenCalledWith({ data: [selectableDataPoints[0].identity.getSelector()] });
        });

        it('PersistPropertiesToHost', () => {
            spyOn(host, "persistProperties");
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndexAndPersist(0, false);
            // Verify that the host is called with the object we expect; the creation is validatd in the next test
            var changes = (<powerbi.visuals.InteractivityService>interactivityService).createPropertiesWithFilter(filterPropertyId);
            expect(host.persistProperties).toHaveBeenCalledWith({ merge: changes });
        });
        
        it('createPropertiesToHost: selecting a dataPoint should result in a VisualObjectInstance', () => {
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(0, false);

            var propertyIdentifier: powerbi.DataViewObjectPropertyIdentifier = {
                objectName: 'general',
                propertyName: 'property'
            };

            var result = (<powerbi.visuals.InteractivityService>interactivityService).createPropertiesWithFilter(propertyIdentifier);
            expect(result.length).toBe(1);
            var firstResult = result[0];
            expect(firstResult.objectName).toBe('general');
            expect(firstResult.properties['property']).toBeDefined();
        });

        it('createPropertiesToHost: no selection should result in empty VisualObjectInstance',() => {
            var propertyIdentifier: powerbi.DataViewObjectPropertyIdentifier = {
                objectName: 'general',
                propertyName: 'property'
            };
            var result = (<powerbi.visuals.InteractivityService>interactivityService).createPropertiesWithFilter(propertyIdentifier);

            expect(result.length).toBe(1);
            var firstResult = result[0];
            expect(firstResult.objectName).toBe('general');
            expect(firstResult.properties['property']).toBeUndefined();
        });

        it('Multiple single selects', () => {
            interactivityService.bind(selectableDataPoints, behavior, null);
            for (var i = 0, ilen = selectableDataPoints.length; i < ilen; i++) {
                behavior.selectIndex(i, false);
                expect(behavior.verifySingleSelectedAt(i)).toBeTruthy();
            }
        });

        it('Single select clears', () => {
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(1, false);
            expect(behavior.verifySingleSelectedAt(1)).toBeTruthy();
            behavior.selectIndex(1, false);
            expect(behavior.verifyCleared()).toBeTruthy();
        });

        it('Basic multiselect', () => {
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(1, true);
            expect(behavior.verifySelectionState([false, true, false, false, false, false])).toBeTruthy();
            behavior.selectIndex(2, true);
            expect(behavior.verifySelectionState([false, true, true, false, false, false])).toBeTruthy();
            behavior.selectIndex(5, true);
            expect(behavior.verifySelectionState([false, true, true, false, false, true])).toBeTruthy();
        });

        it('Multiselect clears', () => {
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(1, true);
            expect(behavior.verifySelectionState([false, true, false, false, false, false])).toBeTruthy();
            behavior.selectIndex(2, true);
            expect(behavior.verifySelectionState([false, true, true, false, false, false])).toBeTruthy();
            behavior.selectIndex(1, true);
            expect(behavior.verifySelectionState([false, false, true, false, false, false])).toBeTruthy();
            behavior.selectIndex(5, true);
            expect(behavior.verifySelectionState([false, false, true, false, false, true])).toBeTruthy();
            behavior.selectIndex(5, true);
            expect(behavior.verifySelectionState([false, false, true, false, false, false])).toBeTruthy();
        });

        it('Single and multiselect', () => {
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(1, false);
            expect(behavior.verifySingleSelectedAt(1)).toBeTruthy();
            behavior.selectIndex(2, true);
            expect(behavior.verifySelectionState([false, true, true, false, false, false])).toBeTruthy();
            behavior.selectIndex(5, true);
            expect(behavior.verifySelectionState([false, true, true, false, false, true])).toBeTruthy();
            behavior.selectIndex(3, false);
            expect(behavior.verifySingleSelectedAt(3)).toBeTruthy();
            behavior.selectIndex(0, true);
            expect(behavior.verifySelectionState([true, false, false, true, false, false])).toBeTruthy();
        });

        it('Multiselect treated as single select when host says selection is invalid', () => {
            host.canSelect = () => false;
            interactivityService.bind(selectableDataPoints, behavior, null);
            behavior.selectIndex(1, true);
            expect(behavior.verifySelectionState([false, true, false, false, false, false])).toBeTruthy();
            behavior.selectIndex(2, true);
            expect(behavior.verifySelectionState([false, false, true, false, false, false])).toBeTruthy();
            behavior.selectIndex(5, true);
            expect(behavior.verifySelectionState([false, false, false, false, false, true])).toBeTruthy();
        });

        it('Legend selection', () => {
            var legendDataPoints = [
                { selected: false, identity: SelectionId.createWithIdAndMeasure(mocks.dataViewScopeIdentity("a"), "queryName") },
                { selected: false, identity: SelectionId.createWithIdAndMeasure(mocks.dataViewScopeIdentity("b"), "queryName") },
            ];
            var legendBehavior = new MockBehavior(legendDataPoints, null);
            interactivityService.bind(selectableDataPoints, behavior, null);
            interactivityService.bind(legendDataPoints, legendBehavior, null, { isLegend: true });

            legendBehavior.selectIndex(0);
            legendBehavior.verifySingleSelectedAt(0);
            behavior.verifySelectionState([true, false, true, false, true, false]);
            expect(interactivityService.hasSelection()).toBeTruthy();
            expect((<powerbi.visuals.InteractivityService>interactivityService).legendHasSelection()).toBeTruthy();

            behavior.selectIndex(1);
            behavior.verifySingleSelectedAt(1);
            legendBehavior.verifyCleared();
            expect(interactivityService.hasSelection()).toBeTruthy();
            expect((<powerbi.visuals.InteractivityService>interactivityService).legendHasSelection()).toBeFalsy();
        });

        it('Slicer selection', () => {
            selectableDataPoints[5].selected = true;
            interactivityService.bind(selectableDataPoints, behavior, null, { overrideSelectionFromData: true });

            // Multiple binds to simulate reloading (should not result in dupes in filter condition).
            selectableDataPoints[5].selected = true;
            interactivityService.bind(selectableDataPoints, behavior, null, { overrideSelectionFromData: true });

            let onSelectSpy = spyOn(host, 'onSelect');

            behavior.selectIndex(0, true);

            expect(behavior.selections()).toEqual([true, false, false, false, false, true]);
            expect(getSelectedIds(interactivityService)).toEqual([
                selectableDataPoints[5].identity,
                selectableDataPoints[0].identity,
            ]);

            expect(host.onSelect).toHaveBeenCalled();
            expect(onSelectSpy.calls.argsFor(0)).toEqual([<powerbi.SelectEventArgs>{
                data: [
                    selectableDataPoints[5].identity.getSelector(),
                    selectableDataPoints[0].identity.getSelector(),
                ]
            }]);
        });
    });

    class MockBehavior implements powerbi.visuals.IInteractiveBehavior {
        private selectableDataPoints: SelectableDataPoint[];
        private selectionHandler: ISelectionHandler;
        private filterPropertyId: powerbi.DataViewObjectPropertyIdentifier;

        constructor(selectableDataPoints: SelectableDataPoint[], filterPropertyId: powerbi.DataViewObjectPropertyIdentifier) {
            this.selectableDataPoints = selectableDataPoints;
            this.filterPropertyId = filterPropertyId;
        }

        public bindEvents(options: any, selectionHandler: ISelectionHandler): void {
            this.selectionHandler = selectionHandler;
        }

        public renderSelection(hasSelection: boolean): void {
            // Stub method to spy on
        }

        public selectIndex(index: number, multiSelect?: boolean): void {
            this.selectionHandler.handleSelection(this.selectableDataPoints[index], !!multiSelect);
        }

        public clear(): void {
            this.selectionHandler.handleClearSelection();
        }

        public selectIndexAndPersist(index: number, multiSelect?: boolean): void {
            this.selectionHandler.handleSelection(this.selectableDataPoints[index], !!multiSelect);
            this.selectionHandler.persistSelectionFilter(this.filterPropertyId);
        }

        public verifyCleared(): boolean {
            var selectableDataPoints = this.selectableDataPoints;
            for (var i = 0, ilen = selectableDataPoints.length; i < ilen; i++) {
                if (selectableDataPoints[i].selected)
                    return false;
            }
            return true;
        }

        public verifySingleSelectedAt(index: number): boolean {
            var selectableDataPoints = this.selectableDataPoints;
            for (var i = 0, ilen = selectableDataPoints.length; i < ilen; i++) {
                var dataPoint = selectableDataPoints[i];
                if (i === index) {
                    if (!dataPoint.selected)
                        return false;
                }
                else if (dataPoint.selected)
                    return false;
            }
            return true;
        }

        public verifySelectionState(selectionState: boolean[]): boolean {
            var selectableDataPoints = this.selectableDataPoints;
            for (var i = 0, ilen = selectableDataPoints.length; i < ilen; i++) {
                if (selectableDataPoints[i].selected !== selectionState[i])
                    return false;
            }
            return true;
        }

        public selections(): boolean[] {
            let selectableDataPoints = this.selectableDataPoints;
            let selections: boolean[] = [];
            for (let dataPoint of selectableDataPoints) {
                selections.push(!!dataPoint.selected);
            }
            return selections;
        }
    }

    function getSelectedIds(interactivityService: powerbi.visuals.IInteractivityService): SelectionId[] {
        // Accessing a private member.
        return interactivityService['selectedIds'];
    }
}