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

    export interface GradientSettings {
        diverging: boolean;
        minColor: any;
        midColor?: any;
        maxColor: any;
        minValue?: number;
        midValue?: number;
        maxValue?: number;
    };

    interface GradientColors {
        minColor: string;
        midColor?: string;
        maxColor: string;
    }

    export module GradientUtils {

        import SQExprBuilder = powerbi.data.SQExprBuilder;
        import DataViewObjectPropertyDefinition = powerbi.data.DataViewObjectPropertyDefinition;
        let DefaultMidColor: string = "#ffffff";

        export function getFillRuleRole(objectDescs: powerbi.data.DataViewObjectDescriptors): string {
            if (!objectDescs)
                return;

            for (let objectName in objectDescs) {
                let objectDesc = objectDescs[objectName];
                for (let propertyName in objectDesc.properties) {
                    let propertyDesc = objectDesc.properties[propertyName];
                    if (propertyDesc.type && propertyDesc.type['fillRule']) {
                        return propertyDesc.rule.inputRole;
                    }
                }
            }
        }

        export function shouldShowGradient(visualConfig): boolean {
            let isShowGradienCard: boolean = visualConfig && visualConfig.query && visualConfig.query.projections && visualConfig.query.projections['Gradient'] ? true : false;
            return isShowGradienCard;
        }

        export function getUpdatedGradientSettings(gradientObject: data.DataViewObjectDefinitions): GradientSettings {
            let gradientSettings: GradientSettings;

            if (gradientObject && !$.isEmptyObject(gradientObject)) {

                gradientSettings = getDefaultGradientSettings();

                for (let propertyName in gradientSettings) {
                    let hasProperty: boolean = (<Object>gradientObject).hasOwnProperty(propertyName);
                    if (hasProperty) {
                        let value: any = gradientObject[propertyName];

                        if (value && value.solid && value.solid.color) {
                            value = value.solid.color;
                        }

                        gradientSettings[propertyName] = value;
                    }
                }
            }

            return gradientSettings;
        }

        export function getGradientMeasureIndex(dataViewCategorical: DataViewCategorical): number {
            if (dataViewCategorical && dataViewCategorical.values && dataViewCategorical.values.grouped) {
                let grouped = dataViewCategorical.values.grouped();
                return DataRoleHelper.getMeasureIndexOfRole(grouped, 'Gradient');
            }
            return -1;
        }

        export function hasGradientRole(dataViewCategorical: DataViewCategorical): boolean {
            let gradientMeasureIndex = getGradientMeasureIndex(dataViewCategorical);
            return gradientMeasureIndex >= 0;
        }

        export function getDefaultGradientSettings(): GradientSettings {

            let colors: GradientColors = getDefaultColors();
            let gradientSettings: GradientSettings = {
                diverging: false,
                minColor: colors.minColor,
                midColor: DefaultMidColor,
                maxColor: colors.maxColor,
                minValue: undefined,
                midValue: undefined,
                maxValue: undefined,
            };

            return gradientSettings;
        }

        export function getDefaultFillRuleDefinition(): DataViewObjectPropertyDefinition {
            return getLinearGradien2FillRuleDefinition();
        }

        export function updateFillRule(propertyName: string, propertyValue: any, definitions: powerbi.data.DataViewObjectDefinitions): void {

            let dataPointProperties: any = definitions["dataPoint"][0].properties;
            let fillRule: FillRuleDefinition = <FillRuleDefinition>dataPointProperties.fillRule;
            let numericValueExpr: data.SQConstantExpr;
            let colorValueExpr: data.SQExpr;

            if (!fillRule) {
                return;
            }

            if ($.isNumeric(propertyValue)) {
                numericValueExpr = propertyValue !== undefined ? SQExprBuilder.double(+propertyValue) : undefined;;
            }

            if (propertyName === "minColor" || propertyName === "midColor" || propertyName === "maxColor") {
                colorValueExpr = getColorExpressionValue(fillRule, propertyName, propertyValue);
            }

            if (propertyName === "minColor") {
                updateMinColor(fillRule, colorValueExpr);
            }
            else if (propertyName === "midColor") {
                updateMidColor(fillRule, colorValueExpr);
            }
            else if (propertyName === "maxColor") {
                updateMaxColor(fillRule, colorValueExpr);
            }
            else if (propertyName === "minValue") {
                updateMinValue(fillRule, numericValueExpr);
            }
            else if (propertyName === "midValue") {
                updateMidValue(fillRule, numericValueExpr);
            }
            else if (propertyName === "maxValue") {
                updateMaxValue(fillRule, numericValueExpr);
            }
            else if (propertyName === "diverging") {
                if (propertyValue) {
                    fillRule = getLinearGradien3FillRuleDefinition(fillRule);
                }
                else {
                    fillRule = getLinearGradien2FillRuleDefinition(fillRule);
                }
                dataPointProperties.fillRule = fillRule;
            }
            else if (propertyName === "revertToDefault") {
                fillRule = this.getDefaultFillRuleDefinition();
                dataPointProperties.fillRule = fillRule;
            }
        }

        export function getGradientSettings(baseFillRule: FillRuleDefinition): GradientSettings {
            if (baseFillRule) {
                return getGradientSettingsFromRule(baseFillRule);
            }
            else {
                return getDefaultGradientSettings();
            }
        }

        export function getFillRule(objectDefinitions: data.DataViewObjectDefinitions): FillRuleDefinition {
            if (objectDefinitions && objectDefinitions["dataPoint"] && objectDefinitions["dataPoint"].length > 0 && objectDefinitions["dataPoint"][0].properties) {
                return <FillRuleDefinition>objectDefinitions["dataPoint"][0].properties['fillRule'];
            }
            return null;
        }

        function getDefaultColors(): GradientColors {

            let dataColors: IDataColorPalette = new powerbi.visuals.DataColorPalette();
            let maxColorInfo: IColorInfo = dataColors.getColorByIndex(0);
            let colors = d3.scale.linear()
                .domain([0, 100])
                .range(["#ffffff", maxColorInfo.value]);
            let maxColor: string = maxColorInfo.value;
            let minColor: string = <any>colors(20);
            let midColor: string = DefaultMidColor;

            return {
                minColor: minColor,
                midColor: midColor,
                maxColor: maxColor,
            };
        }

        export function getGradientSettingsFromRule(fillRule: FillRuleDefinition): GradientSettings {
            let maxColor: string;
            let minColor: string;
            let midColor: string = DefaultMidColor;
            let maxValue: number;
            let midValue: number;
            let minValue: number;
            let diverging: boolean = fillRule.linearGradient3 !== undefined;

            if (fillRule.linearGradient2) {
                let maxColorExpr: any = fillRule.linearGradient2.max.color;
                let minColorExpr: any = fillRule.linearGradient2.min.color;
                let maxValueExpr: any = fillRule.linearGradient2.max.value;
                let minValueExpr: any = fillRule.linearGradient2.min.value;
                maxColor = maxColorExpr.value;
                minColor = minColorExpr.value;
                if (maxValueExpr) {
                    maxValue = <number>maxValueExpr.value;
                }
                if (minValueExpr) {
                    minValue = <number>minValueExpr.value;
                }
            }
            else if (fillRule.linearGradient3) {
                let maxColorExpr: any = fillRule.linearGradient3.max.color;
                let midColorExpr: any = fillRule.linearGradient3.mid.color;
                let minColorExpr: any = fillRule.linearGradient3.min.color;
                let maxValueExpr: any = fillRule.linearGradient3.max.value;
                let midValueExpr: any = fillRule.linearGradient3.mid.value;
                let minValueExpr: any = fillRule.linearGradient3.min.value;
                maxColor = maxColorExpr.value;
                midColor = midColorExpr.value;
                minColor = minColorExpr.value;
                if (maxValueExpr) {
                    maxValue = <number>maxValueExpr.value;
                }
                if (midValueExpr) {
                    midValue = <number>midValueExpr.value;
                }
                if (minValueExpr) {
                    minValue = <number>minValueExpr.value;
                }
            }

            return {
                diverging: diverging,
                minColor: minColor,
                midColor: midColor,
                maxColor: maxColor,
                minValue: minValue,
                midValue: midValue,
                maxValue: maxValue,
            };
        }

        function getLinearGradien2FillRuleDefinition(baseFillRule?: FillRuleDefinition): DataViewObjectPropertyDefinition {
            let gradientSettings: GradientSettings = getGradientSettings(baseFillRule);
            let fillRuleDefinition: FillRuleDefinition = {
                linearGradient2: {
                    max: { color: SQExprBuilder.text(gradientSettings.maxColor) },
                    min: { color: SQExprBuilder.text(gradientSettings.minColor) },
                }
            };

            return fillRuleDefinition;
        }

        function getLinearGradien3FillRuleDefinition(baseFillRule?: FillRuleDefinition): DataViewObjectPropertyDefinition {
            let gradientSettings: GradientSettings = getGradientSettings(baseFillRule);
            let fillRuleDefinition: FillRuleDefinition = {
                linearGradient3: {
                    max: { color: SQExprBuilder.text(gradientSettings.maxColor) },
                    mid: { color: SQExprBuilder.text(gradientSettings.midColor) },
                    min: { color: SQExprBuilder.text(gradientSettings.minColor) },
                }
            };

            return fillRuleDefinition;
        }

        function getDefaultColorExpression(fillRule: FillRuleDefinition, propertyName: string): data.SQExpr {
            let defaultColor: data.SQExpr;
            let defaultFillRule: FillRuleDefinition;

            if (fillRule.linearGradient3) {
                defaultFillRule = getLinearGradien3FillRuleDefinition();
                if (propertyName === "minColor") {
                    defaultColor = defaultFillRule.linearGradient3.min.color;
                }
                else if (propertyName === "midColor") {
                    defaultColor = defaultFillRule.linearGradient3.mid.color;
                }
                else if (propertyName === "maxColor") {
                    defaultColor = defaultFillRule.linearGradient3.max.color;
                }
            }
            else if (fillRule.linearGradient2) {
                defaultFillRule = getLinearGradien2FillRuleDefinition();
                if (propertyName === "minColor") {
                    defaultColor = defaultFillRule.linearGradient2.min.color;
                }
                else if (propertyName === "maxColor") {
                    defaultColor = defaultFillRule.linearGradient2.max.color;
                }
            }

            return defaultColor;
        }

        function getColorExpressionValue(fillRule: FillRuleDefinition, propertyName: string, propertyValue: string): data.SQExpr {
            let colorExpressionValue: data.SQExpr;
            if (propertyValue) {
                colorExpressionValue = SQExprBuilder.text(propertyValue);
            }
            else {
                colorExpressionValue = getDefaultColorExpression(fillRule, propertyName);
            }
            return colorExpressionValue;
        }

        function updateMinColor(fillRule: FillRuleDefinition, colorExpressionValue: data.SQExpr) {
            if (fillRule.linearGradient2) {
                fillRule.linearGradient2.min.color = colorExpressionValue;
            }
            else if (fillRule.linearGradient3) {
                fillRule.linearGradient3.min.color = colorExpressionValue;
            }
        }

        function updateMidColor(fillRule: FillRuleDefinition, colorExpressionValue: data.SQExpr) {
            if (fillRule.linearGradient3) {
                fillRule.linearGradient3.mid.color = colorExpressionValue;
            }
        }

        function updateMaxColor(fillRule: FillRuleDefinition, colorExpressionValue: data.SQExpr) {
            if (fillRule.linearGradient2) {
                fillRule.linearGradient2.max.color = colorExpressionValue;
            }
            else if (fillRule.linearGradient3) {
                fillRule.linearGradient3.max.color = colorExpressionValue;
            }
        }

        function updateMinValue(fillRule: FillRuleDefinition, value: data.SQConstantExpr) {
            if (fillRule.linearGradient2) {
                fillRule.linearGradient2.min.value = value;
            }
            else if (fillRule.linearGradient3) {
                fillRule.linearGradient3.min.value = value;
            }
        }

        function updateMidValue(fillRule: FillRuleDefinition, value: data.SQConstantExpr) {
            if (fillRule.linearGradient3) {
                fillRule.linearGradient3.mid.value = value;
            }
        }

        function updateMaxValue(fillRule: FillRuleDefinition, value: data.SQConstantExpr) {
            if (fillRule.linearGradient2) {
                fillRule.linearGradient2.max.value = value;
            }
            else if (fillRule.linearGradient3) {
                fillRule.linearGradient3.max.value = value;
            }
        }
    };
} 