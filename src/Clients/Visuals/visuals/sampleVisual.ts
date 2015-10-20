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
import SelectionManager = utility.SelectionManager;

    export var cheerMeterProps = {
        dataPoint: {
            defaultColor: <DataViewObjectPropertyIdentifier>{
                objectName: 'dataPoint',
                propertyName: 'defaultColor'
            },
            fill: <DataViewObjectPropertyIdentifier>{
                objectName: 'dataPoint',
                propertyName: 'fill'
            },
        },
    };

    export interface TeamData {
        name: string;
        value: number;
        color: string;
        identity: SelectionId;
    }

    export interface CheerData {
        teamA: TeamData;
        teamB: TeamData;
        background: string;
    }

    interface CheerLayout {
        x1: number;
        x2: number;
        y1: number;
        y2: number;
        fontSize: string;
    }

    export class CheerMeter implements IVisual {
        public static capabilities: VisualCapabilities = {
            dataRoles: [
                {
                    name: 'Category',
                    kind: powerbi.VisualDataRoleKind.Grouping,
                },
                {
                    displayName: 'Noise Measure',
                    name: 'Y',
                    kind: powerbi.VisualDataRoleKind.Measure,
                },
            ],
            dataViewMappings: [{
                categorical: {
                    categories: {
                        for: { in: 'Category' },
                    },
                    values: {
                        select: [{ bind: { to: 'Y' } }]
                    },
                },
            }],
            objects: {
                dataPoint: {
                    displayName: data.createDisplayNameGetter('Visual_DataPoint'),
                    properties: {
                        fill: {
                            displayName: data.createDisplayNameGetter('Visual_Fill'),
                            type: { fill: { solid: { color: true } } }
                        },
                        width: {
                            displayName: '',
                            type: { numeric: true }
                        }
                    }
                },
                general: {
                    displayName: 'General',
                    properties: {
                        fill: {
                            displayName: 'Background color',
                            type: { fill: { solid: { color: true } } }
                        },

                    }
                }
            }
        };

        private static DefaultFontFamily = 'cursive';
        private static DefaultFontColor = 'rgb(165, 172, 175)';
        private static DefaultBackgroundColor = '#243C18';
        private static PaddingBetweenText = 15;

        private textOne: D3.Selection;
        private textTwo: D3.Selection;
        private svg: D3.Selection;
        private isFirstTime: boolean = true;
        private data: CheerData;
        private selectionManager: SelectionManager;

        public static converter(dataView: DataView): CheerData {
            if (!dataView.categorical || !dataView.categorical.categories) return null;
            let cat = dataView.categorical.categories[0];
            if (!cat) return null;
            let catValues = cat.values;
            if (!catValues || !dataView.categorical.values) return null;
            let values = dataView.categorical.values[0].values;
            let objects = dataView.categorical.categories[0].objects;
            let object1 = objects && objects.length > 0 ? objects[0] : undefined;
            let object2 = objects && objects.length > 1 ? objects[1] : undefined;
            let metadataObjects = dataView.metadata.objects;
            let backgroundColor = CheerMeter.DefaultBackgroundColor;
            if (metadataObjects) {
                let general = metadataObjects['general'];
                if (general) {
                    let fill = <Fill>general['fill'];
                    if (fill) {
                        backgroundColor = fill.solid.color;
                    }
                }
            }

            let color1 = DataViewObjects.getFillColor(
                object1,
                cheerMeterProps.dataPoint.fill,
                CheerMeter.DefaultFontColor);

            let color2 = DataViewObjects.getFillColor(
                object2,
                cheerMeterProps.dataPoint.fill,
                CheerMeter.DefaultFontColor);

            var idn1 = SelectionIdBuilder.builder()
                .withCategory(cat, 0)
                .createSelectionId();
            var idn2 = SelectionIdBuilder.builder()
                .withCategory(cat, 1)
                .createSelectionId();

            var data = {
                teamA: {
                    name: catValues[0],
                    value: values[0],
                    color: color1,
                    identity: idn1
                },
                teamB: {
                    name: catValues[1],
                    value: values[1],
                    color: color2,
                    identity: idn2
                },
                background: backgroundColor
            };

            return data;
        }

        public init(options: VisualInitOptions): void {
            this.selectionManager = new SelectionManager({ hostServices: options.host });
            var svg = this.svg = d3.select(options.element.get(0)).append('svg');

            this.textOne = svg.append('text')
                .style('font-family', CheerMeter.DefaultFontFamily);

            this.textTwo = svg.append('text')
                .style('font-family', CheerMeter.DefaultFontFamily);
        }

        public update(options: VisualUpdateOptions) {
            if (!options.dataViews[0]) { return; }
            let data = this.data = CheerMeter.converter(options.dataViews[0]);
            if (!data) return;
            let duration = options.suppressAnimations ? 0 : AnimatorCommon.MinervaAnimationDuration;
            this.draw(data, duration, options.viewport);
        }

        private getRecomendedFontProperties(text1: string, text2: string, parentViewport: IViewport): TextProperties {
            let textProperties: TextProperties = {
                fontSize: '',
                fontFamily: CheerMeter.DefaultFontFamily,
                text: text1 + text2
            };

            let min = 1;
            let max = 1000;
            let i;
            let maxWidth = parentViewport.width;
            let width = 0;

            while (min <= max) {
                i = (min + max) / 2 | 0;

                textProperties.fontSize = i + 'px';
                width = TextMeasurementService.measureSvgTextWidth(textProperties);

                if (maxWidth > width)
                    min = i + 1;
                else if (maxWidth < width)
                    max = i - 1;
                else
                    break;
            }

            textProperties.fontSize = i + 'px';
            width = TextMeasurementService.measureSvgTextWidth(textProperties);
            if (width > maxWidth) {
                i--;
                textProperties.fontSize = i + 'px';
            }

            return textProperties;
        }

        private calculateLayout(data: CheerData, viewport: IViewport): CheerLayout {
            let text1 = data.teamA.name;
            let text2 = data.teamB.name;

            let avaliableViewport: IViewport = {
                height: viewport.height,
                width: viewport.width - CheerMeter.PaddingBetweenText
            };
            let recomendedFontProperties = this.getRecomendedFontProperties(text1, text2, avaliableViewport);

            recomendedFontProperties.text = text1;
            let width1 = TextMeasurementService.measureSvgTextWidth(recomendedFontProperties) | 0;

            recomendedFontProperties.text = text2;
            let width2 = TextMeasurementService.measureSvgTextWidth(recomendedFontProperties) | 0;

            let padding = ((viewport.width - width1 - width2 - CheerMeter.PaddingBetweenText) / 2) | 0;

            recomendedFontProperties.text = text1 + text2;
            let offsetHeight = (TextMeasurementService.measureSvgTextHeight(recomendedFontProperties)) | 0;

            let max = data.teamA.value + data.teamB.value;
            let availableHeight = viewport.height - offsetHeight;
            let y1 = (((max - data.teamA.value) / max) * availableHeight + offsetHeight / 2) | 0;
            let y2 = (((max - data.teamB.value) / max) * availableHeight + offsetHeight / 2) | 0;

            return {
                x1: padding,
                x2: padding + width1 + CheerMeter.PaddingBetweenText,
                y1: y1,
                y2: y2,
                fontSize: recomendedFontProperties.fontSize
            };
        }

        private ensureStartState(layout: CheerLayout, viewport: IViewport) {
            if (this.isFirstTime) {
                this.isFirstTime = false;
                let startY = viewport.height / 2;
                this.textOne.attr(
                    {
                        'x': layout.x1,
                        'y': startY
                    });

                this.textTwo.attr(
                    {
                        'x': layout.x2,
                        'y': startY
                    });
            }
        }

        private clearSelection() {
            this.selectionManager.clear().then(() => {
                this.clearSelectionUI();
            });
        }

        private clearSelectionUI() {
            this.textOne.style('stroke', '#FFF').style('stroke-width', 0);
            this.textTwo.style('stroke', '#FFF').style('stroke-width', 0);
        }

        private updateSelectionUI(ids: SelectionId[]) {
            this.textOne.style('stroke', '#FFF').style('stroke-width', SelectionManager.containsSelection(ids, this.data.teamA.identity) ? '2px' : '0px');
            this.textTwo.style('stroke', '#FFF').style('stroke-width', SelectionManager.containsSelection(ids, this.data.teamB.identity) ? '2px' : '0px');
        }

        private draw(data: CheerData, duration: number, viewport: IViewport) {
            let easeName = 'back';
            let textOne = this.textOne;
            let textTwo = this.textTwo;

            this.svg
                .attr({
                    'height': viewport.height,
                    'width': viewport.width
                })
                .on('click', () => {
                    this.clearSelection();
                })
                .style('background-color', data.background);

            let layout = this.calculateLayout(data, viewport);

            this.ensureStartState(layout, viewport);

            textOne
                .style('font-size', layout.fontSize)
                .style('fill', data.teamA.color)
                .on('click', () => {
                    this.selectionManager.select(data.teamA.identity, d3.event.ctrlKey).then((ids) => {
                        this.updateSelectionUI(ids);
                    });
                    d3.event.stopPropagation();
                })
                .text(data.teamA.name);

            textTwo
                .style('font-size', layout.fontSize)
                .style('fill', data.teamB.color)
                .on('click', () => {
                    this.selectionManager.select(data.teamB.identity, d3.event.ctrlKey).then((ids) => {
                        this.updateSelectionUI(ids);
                    });
                    d3.event.stopPropagation();
                })
                .text(data.teamB.name);

            textOne.transition()
                .duration(duration)
                .ease(easeName)
                .attr({
                    y: layout.y1,
                    x: layout.x1
                });

            textTwo.transition()
                .duration(duration)
                .ease(easeName)
                .attr({
                    y: layout.y2,
                    x: layout.x2
                });
        }

        public destroy(): void {
            this.svg = null;
            this.textOne = this.textTwo = null;
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] {
            let instances: VisualObjectInstance[] = [];
            let data = this.data;
            switch (options.objectName) {
                case 'dataPoint':
                    if (data) {
                        let teams = [data.teamA, data.teamB];

                        for (let i = 0; i < teams.length; i++) {
                            let slice = teams[i];

                            let color = slice.color;
                            let selector = slice.identity;

                            let dataPointInstance: VisualObjectInstance = {
                                objectName: 'dataPoint',
                                displayName: slice.name,
                                selector: selector,
                                properties: {
                                    fill: { solid: { color: color } }
                                },
                            };

                            instances.push(dataPointInstance);
                        };
                    }
                    break;
                case 'general':
                    let general: VisualObjectInstance = {
                        objectName: 'general',
                        displayName: 'General',
                        selector: null,
                        properties: {
                            fill: { solid: { color: data ? data.background : CheerMeter.DefaultBackgroundColor } }
                        }
                    };
                    instances.push(general);
                    break;
            }

            return instances;
        }
    }
}