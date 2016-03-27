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
    import DataViewTransform = powerbi.data.DataViewTransform;
    import ValueType = powerbi.ValueType;
    import PrimitiveType = powerbi.PrimitiveType;
    import DataLabelUtils = powerbi.visuals.dataLabelUtils;
    import ObjectEnumerationBuilder = powerbi.visuals.ObjectEnumerationBuilder;

    powerbitests.mocks.setLocale();

    describe("DataLabelUtils", () => {

        afterEach(() => {
            $(".data-labels").remove();
        });

        describe("dataLabelUtils tests", () => {

            it('display units formatting values : Auto', () => {
                var value: number = 2000000;
                var labelSettings: powerbi.visuals.VisualDataLabelsSettings = DataLabelUtils.getDefaultLabelSettings();
                labelSettings.displayUnits = 0;
                labelSettings.precision = 0;
                var value2 = 1000000;
                var formattersCache = DataLabelUtils.createColumnFormatterCacheManager();
                var formatter = formattersCache.getOrCreate(null, labelSettings, value2);
                var formattedValue = formatter.format(value);
                expect(formattedValue).toBe("2M");
            });

            it('display units formatting values : None', () => {
                var value: number = 20000;
                var labelSettings: powerbi.visuals.VisualDataLabelsSettings = DataLabelUtils.getDefaultLabelSettings();
                labelSettings.displayUnits = 10;
                labelSettings.precision = 0;
                var formattersCache = DataLabelUtils.createColumnFormatterCacheManager();
                var formatter = formattersCache.getOrCreate(null, labelSettings);
                var formattedValue = formatter.format(value);
                expect(formattedValue).toBe("20,000");
            });

            it("display units formatting values : K", () => {
                var value: number = 20000;
                var labelSettings: powerbi.visuals.VisualDataLabelsSettings = DataLabelUtils.getDefaultLabelSettings();
                labelSettings.displayUnits = 10000;
                labelSettings.precision = 0;
                var formattersCache = DataLabelUtils.createColumnFormatterCacheManager();
                var formatter = formattersCache.getOrCreate(null, labelSettings);
                var formattedValue = formatter.format(value);
                expect(formattedValue).toBe("20K");
            });

            it("display units formatting values : M", () => {
                var value: number = 200000;
                var labelSettings: powerbi.visuals.VisualDataLabelsSettings = DataLabelUtils.getDefaultLabelSettings();
                labelSettings.displayUnits = 1000000;
                labelSettings.precision = 1;
                var formattersCache = DataLabelUtils.createColumnFormatterCacheManager();
                var formatter = formattersCache.getOrCreate(null, labelSettings);
                var formattedValue = formatter.format(value);
                expect(formattedValue).toBe("0.2M");
            });

            it("display units formatting values : B", () => {
                var value: number = 200000000000;
                var labelSettings: powerbi.visuals.VisualDataLabelsSettings = DataLabelUtils.getDefaultLabelSettings();
                labelSettings.displayUnits = 1000000000;
                labelSettings.precision = 0;
                var formattersCache = DataLabelUtils.createColumnFormatterCacheManager();
                var formatter = formattersCache.getOrCreate(null, labelSettings);
                var formattedValue = formatter.format(value);
                expect(formattedValue).toBe("200bn");
            });

            it("display units formatting values : T", () => {
                var value: number = 200000000000;
                var labelSettings: powerbi.visuals.VisualDataLabelsSettings = DataLabelUtils.getDefaultLabelSettings();
                labelSettings.displayUnits = 1000000000000;
                labelSettings.precision = 1;
                var formattersCache = DataLabelUtils.createColumnFormatterCacheManager();
                var formatter = formattersCache.getOrCreate(null, labelSettings);
                var formattedValue = formatter.format(value);
                expect(formattedValue).toBe("0.2T");
            });

            it("label formatting - multiple formats", () => {
                var formatCol1 = "#,0.0";
                var formatCol2 = "$#,0.0";
                var value: number = 1545;
                var labelSettings: powerbi.visuals.VisualDataLabelsSettings = DataLabelUtils.getDefaultLabelSettings();
                labelSettings.displayUnits = null;
                labelSettings.precision = 1;

                var formattersCache = DataLabelUtils.createColumnFormatterCacheManager();
                var formatter1 = formattersCache.getOrCreate(formatCol1, labelSettings);
                var formattedValue = formatter1.format(value);

                expect(formattedValue).toBe("1,545.0");

                var formatter2 = formattersCache.getOrCreate(formatCol2, labelSettings);
                formattedValue = formatter2.format(value);

                expect(formattedValue).toBe("$1,545.0");
            });
        });

        describe("Test enumerateCategoryLabels", () => {
            it("test default values", () => {
                let labelSettings = DataLabelUtils.getDefaultPointLabelSettings();

                let enumerationWithColor = new ObjectEnumerationBuilder();
                DataLabelUtils.enumerateCategoryLabels(enumerationWithColor, labelSettings, true);
                let objectsWithColor = enumerationWithColor.complete().instances;

                let enumerationNoColor = new ObjectEnumerationBuilder();
                DataLabelUtils.enumerateCategoryLabels(enumerationNoColor, labelSettings, false);
                let objectsNoColor = enumerationNoColor.complete().instances;

                expect(objectsWithColor[0].properties["show"]).toBe(false);
                expect(objectsNoColor[0].properties["show"]).toBe(false);

                expect(objectsWithColor[0].properties["color"]).toBe(labelSettings.labelColor);
                expect(objectsNoColor[0].properties["color"]).toBeUndefined();
            });

            it("test custom values", () => {
                let labelSettings = DataLabelUtils.getDefaultPointLabelSettings();
                labelSettings.show = true;
                labelSettings.labelColor = "#FF0000";

                let enumerationWithColor = new ObjectEnumerationBuilder();
                DataLabelUtils.enumerateCategoryLabels(enumerationWithColor, labelSettings, true);
                let objectsWithColor = enumerationWithColor.complete().instances;

                expect(objectsWithColor[0].properties["show"]).toBe(true);
                expect(objectsWithColor[0].properties["color"]).toBe("#FF0000");
            });

            it("test category labels objects for donut chart", () => {
                let labelSettings = DataLabelUtils.getDefaultDonutLabelSettings();

                let enumerationWithColor = new ObjectEnumerationBuilder();
                DataLabelUtils.enumerateCategoryLabels(enumerationWithColor, labelSettings, false, true);
                let objectsWithColor = enumerationWithColor.complete().instances;

                expect(objectsWithColor[0].properties["show"]).toBe(labelSettings.showCategory);
            });

            it("test null values", () => {
                let labelSettings = DataLabelUtils.getDefaultPointLabelSettings();
                let donutLabelSettings = DataLabelUtils.getDefaultDonutLabelSettings();

                let enumerationWithColor = new ObjectEnumerationBuilder();
                DataLabelUtils.enumerateCategoryLabels(enumerationWithColor, null, true);
                let objectsWithColor = enumerationWithColor.complete().instances;

                let enumerationCategories = new ObjectEnumerationBuilder();
                DataLabelUtils.enumerateCategoryLabels(enumerationCategories, null, false, true);
                let donutObjectsWithColor = enumerationCategories.complete().instances;

                expect(objectsWithColor[0].properties["show"]).toBe(labelSettings.show);
                expect(objectsWithColor[0].properties["color"]).toBe(labelSettings.labelColor);

                expect(donutObjectsWithColor[0].properties["show"]).toBe(donutLabelSettings.showCategory);
            });
        });
    });

    function columnChartDataLabelsShowValidation(chartType: string, collide: boolean) {
        var visualBuilder: VisualBuilder;

        var dataViewMetadataThreeColumn: powerbi.DataViewMetadataColumn[] = [
            {
                displayName: "col1",
                queryName: "col1",
                type: ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Text)
            },
            {
                displayName: "col2",
                queryName: "col2",
                isMeasure: true,
                type: ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Double)
            },
            {
                displayName: "col3",
                queryName: "col3",
                isMeasure: true,
                type: ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Double)
            }
        ];

        function createMetadata(columns): powerbi.DataViewMetadata {
            var metadata: powerbi.DataViewMetadata = {
                columns: columns,
            };

            metadata.objects = {
                labels: {
                    show: true,
                    labelPrecision: 0,
                    color: {
                        solid: {
                            color: "#FF0000"
                        }
                    }
                }
            };

            return metadata;
        }

        beforeEach(() => {
            if (collide) {
                visualBuilder = new VisualBuilder("100", "100", chartType);
            } else {
                visualBuilder = new VisualBuilder("500", "500", chartType);
            }

            visualBuilder.metadata = createMetadata(dataViewMetadataThreeColumn);
        });
    }

    describe("Stacked Bar Chart show labels validation", () =>
        columnChartDataLabelsShowValidation("barChart", false));
    describe("Clustered Bar Chart show labels validation", () =>
        columnChartDataLabelsShowValidation("clusteredBarChart", false));
    describe("Hundred Percent Stacked Bar Chart show labels validation", () =>
        columnChartDataLabelsShowValidation("hundredPercentStackedBarChart", false));
    describe("Stacked Column Chart show labels validation", () =>
        columnChartDataLabelsShowValidation("columnChart", false));
    describe("Clustered Column Chart show labels validation", () =>
        columnChartDataLabelsShowValidation("clusteredColumnChart", false));
    describe("Hundred Percent Stacked Column Chart show labels validation", () =>
        columnChartDataLabelsShowValidation("hundredPercentStackedColumnChart", false));

    describe("Stacked Bar Chart hide labels validation", () =>
        columnChartDataLabelsShowValidation("barChart", true));
    describe("Clustered Bar Chart hide labels validation", () =>
        columnChartDataLabelsShowValidation("clusteredBarChart", true));
    describe("Hundred Percent Stacked Bar Chart hide labels validation", () =>
        columnChartDataLabelsShowValidation("hundredPercentStackedBarChart", true));
    describe("Stacked Column Chart hide labels validation", () =>
        columnChartDataLabelsShowValidation("columnChart", true));
    describe("Clustered Column Chart hide labels validation", () =>
        columnChartDataLabelsShowValidation("clusteredColumnChart", true));
    describe("Hundred Percent Stacked Column Chart hide labels validation", () =>
        columnChartDataLabelsShowValidation("hundredPercentStackedColumnChart", true));

    class VisualBuilder {
        private visual: powerbi.IVisual;

        private element: JQuery;

        private host: powerbi.IVisualHostServices;

        private style: powerbi.IVisualStyle;

        public metadata: powerbi.DataViewMetadata;

        public values: any[] = [];

        public categoriesValues: any[] = [];

        public categoryIdentities: powerbi.DataViewScopeIdentity[];

        public isIdentity: boolean = false;

        private dataView: powerbi.DataView;

        private height: string;

        private width: string;

        private _pluginName: string;

        public get pluginName(): string {
            return this._pluginName;
        }

        constructor(height: string, width: string, pluginName: string) {
            this.width = width;
            this.height = height;
            this._pluginName = pluginName;

            this.init();
        }

        private init() {
            this.element = powerbitests.helpers.testDom(this.height, this.width);
            this.host = powerbitests.mocks.createVisualHostServices();
            this.style = powerbi.visuals.visualStyles.create();

            this.createVisual();
        }

        public setSize(width: string, height: string) {
            this.width = width;
            this.height = height;

            this.init();
        }

        private createVisual() {
            this.visual = powerbi.visuals.visualPluginFactory.create().getPlugin(this._pluginName).create();

            this.visual.init(this.createInitOptions());
        }

        private createInitOptions(): powerbi.VisualInitOptions {
            return <powerbi.VisualInitOptions> {
                element: this.element,
                host: this.host,
                style: this.style,
                viewport: {
                    height: this.element.height(),
                    width: this.element.width()
                },
                animation: { transitionImmediate: true }
            };
        }

        protected createIdentities() {
            if (!this.isIdentity) {
                this.categoryIdentities = undefined;

                return;
            }

            var categoryIdentities: powerbi.DataViewScopeIdentity[] = [];

            for (var i = 0; i < this.categoriesValues.length; i++) {
                categoryIdentities.push(mocks.dataViewScopeIdentity(this.categoriesValues[i]));
            }

            this.categoryIdentities = categoryIdentities;
        }

        private createDataView() {
            this.dataView = <powerbi.DataView> {
                metadata: this.metadata,
                categorical: {
                    categories: [{
                        source: this.metadata.columns[0],
                        values: this.categoriesValues,
                        identity: this.categoryIdentities
                    }],
                    values: DataViewTransform.createValueColumns([{
                        source: this.metadata.columns[1],
                        values: this.values
                    }])
                }
            };
        }

        public update() {
            this.createIdentities();
            this.createDataView();
        }

        public onDataChanged() {
            this.update();

            this.visual.onDataChanged({
                dataViews: [this.dataView]
            });
        }
    }

    class MapBuilder {
        public datalabelSettings: powerbi.visuals.PointDataLabelsSettings = {
            show: true,
            displayUnits: 2,
            position: powerbi.visuals.PointLabelPosition.Above,
            precision: 2,
            labelColor: "#000000",
            formatterOptions: {}
        };

        public viewPort = {
            height: 150,
            width: 300
        };

        public layout(): powerbi.visuals.ILabelLayout {
            return DataLabelUtils.getMapLabelLayout(this.datalabelSettings);
        }

        public graphicsContext(className: string): D3.Selection {
            return d3.select("body")
                .append("svg")
                .style("position", "absolute")
                .append("g")
                .classed(className, true);
        }

        public build(x: number, y: number, labeltext: string) {
            return {
                x: x,
                y: y,
                radius: 10,
                fill: "#000000",
                stroke: "2",
                strokeWidth: 2,
                selected: true,
                identity: null,
                labeltext: labeltext
            };
        }
    }

    class MapBubbleBuilder extends MapBuilder {
        public buildMapBubble(x: number, y: number, labeltext: string): powerbi.visuals.MapBubble {
            return <powerbi.visuals.MapBubble> this.build(x, y, labeltext);
        }

        public getResult(data: any[], className: string): D3.UpdateSelection {
            return DataLabelUtils.drawDefaultLabelsForDataPointChart(
                data,
                this.graphicsContext(className),
                this.layout(),
                this.viewPort);
        }
    }

    class MapSliceBuilder extends MapBuilder {
        public buildMapSlice(x: number, y: number, labeltext: string, value: any): powerbi.visuals.MapSlice {
            var map: any = this.build(x, y, labeltext);
            map.value = value;

            return <powerbi.visuals.MapSlice> map;
        }

        public getResult(data: any[], className: string) {
            return DataLabelUtils.drawDefaultLabelsForDataPointChart(
                data,
                this.graphicsContext(className),
                this.layout(),
                this.viewPort);
        }
    }
}
