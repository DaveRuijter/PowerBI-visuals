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
    import ObjectEnumerationBuilder = powerbi.visuals.ObjectEnumerationBuilder;
    import VisualObjectInstance = powerbi.VisualObjectInstance;
    import VisualObjectInstanceContainer = powerbi.VisualObjectInstanceContainer;

    describe("ObjectEnumerationBuilder", () => {
        it('Empty creates undefined', () => {
            let enumeration = new ObjectEnumerationBuilder();

            expect(enumeration.complete()).toBeUndefined();
        });

        it('Single instance', () => {
            let enumeration = new ObjectEnumerationBuilder();
            let instance: VisualObjectInstance = {
                objectName: 'myObject',
                selector: null,
                properties: {
                    'myProperty': 123,
                },
            };

            enumeration.pushInstance(instance);

            expect(enumeration.complete()).toEqual({
                instances: [instance],
            });
        });

        it('Multi instance: Same Object is consolidated', () => {
            let enumeration = new ObjectEnumerationBuilder();
            let instance1: VisualObjectInstance = {
                objectName: 'myObject',
                selector: null,
                properties: {
                    'myProperty1': 123,
                },
            };
            let instance2: VisualObjectInstance = {
                objectName: 'myObject',
                selector: null,
                properties: {
                    'myProperty2': 321,
                },
            };

            enumeration
                .pushInstance(instance1)
                .pushInstance(instance2);

            expect(enumeration.complete()).toEqual({
                instances: [{
                    objectName: 'myObject',
                    selector: null,
                    properties: {
                        'myProperty1': 123,
                        'myProperty2': 321,
                    },
                }],
            });
        });

        it('Multi instance: Same Object with same property', () => {
            let enumeration = new ObjectEnumerationBuilder();
            let instance1: VisualObjectInstance = {
                objectName: 'myObject',
                selector: null,
                properties: {
                    'myProperty': 123,
                },
            };
            let instance2: VisualObjectInstance = {
                objectName: 'myObject',
                selector: null,
                properties: {
                    'myProperty': 321,
                },
            };

            enumeration
                .pushInstance(instance1)
                .pushInstance(instance2);

            expect(enumeration.complete()).toEqual({
                instances: [{
                    objectName: 'myObject',
                    selector: null,
                    properties: {
                        'myProperty': 123,
                    },
                }],
            });
        });

        it('Multi instance: Different Object is not consolidated', () => {
            let enumeration = new ObjectEnumerationBuilder();
            let instance1: VisualObjectInstance = {
                objectName: 'myObject',
                selector: null,
                properties: {
                    'myProperty1': 123,
                },
            };
            let instance2: VisualObjectInstance = {
                objectName: 'myObject2',
                selector: null,
                properties: {
                    'myProperty2': 321,
                },
            };

            enumeration
                .pushInstance(instance1)
                .pushInstance(instance2);

            expect(enumeration.complete()).toEqual({
                instances: [
                    instance1,
                    instance2,
                ],
            });
        });

        it('Multi instance: Different selector is not consolidated', () => {
            let enumeration = new ObjectEnumerationBuilder();
            let instance1: VisualObjectInstance = {
                objectName: 'myObject',
                selector: null,
                properties: {
                    'myProperty1': 123,
                },
            };
            let instance2: VisualObjectInstance = {
                objectName: 'myObject',
                selector: { metadata: 'select2' },
                properties: {
                    'myProperty2': 321,
                },
            };

            enumeration
                .pushInstance(instance1)
                .pushInstance(instance2);

            expect(enumeration.complete()).toEqual({
                instances: [
                    instance1,
                    instance2,
                ],
            });
        });

        it('Multi instance: Different containers', () => {
            let enumeration = new ObjectEnumerationBuilder();
            let instance1: VisualObjectInstance = {
                objectName: 'myObject',
                selector: null,
                properties: {
                    'myProperty1': 123,
                },
            };
            let instance2: VisualObjectInstance = {
                objectName: 'myObject',
                selector: null,
                properties: {
                    'myProperty2': 321,
                },
            };
            let container: VisualObjectInstanceContainer = {
                displayName: 'containerName',
            };

            enumeration
                .pushInstance(instance1)
                .pushContainer(container)
                .pushInstance(instance2)
                .popContainer();

            let result = enumeration.complete();
            expect(result).toEqual({
                instances: [
                    instance1,
                    instance2,
                ],
                containers: [
                    container,
                ]
            });
            expect(result.instances[0].containerIdx).toBeUndefined();
            expect(result.instances[1].containerIdx).toBe(0);
        });
    });
}