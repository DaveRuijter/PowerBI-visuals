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
    import KpiImageSize = powerbi.visuals.KpiUtil.KpiImageSize;
    import getKpiImageMetadata = powerbi.visuals.KpiUtil.getKpiImageMetadata;

    export interface CardStyle {
        card: {
            maxFontSize: number;
        };
        label: {
            fontSize: number;
            color: string;
            height: number;
        };
        value: {
            fontSize: number;
            color: string;
            fontFamily: string;
        };
    }

    export interface CardConstructorOptions {
        isScrollable?: boolean;
        displayUnitSystemType?: DisplayUnitSystemType;
        animator?: IGenericAnimator;
    }

    export interface CardFormatSetting {
        showTitle: boolean;
        labelSettings: VisualDataLabelsSettings;
        wordWrap: boolean;
    }

    export class Card extends AnimatedText implements IVisual {
        private static cardClassName: string = 'card';
        private static Label: ClassAndSelector = {
            class: 'label',
            selector: '.label'
        };
        private static Value: ClassAndSelector = {
            class: 'value',
            selector: '.value'
        };
        public static DefaultStyle: CardStyle = {
            card: {
                maxFontSize: 200
            },
            label: {
                fontSize: 16,
                color: '#a6a6a6',
                height: 26
            },
            value: {
                fontSize: 37,
                color: '#333333',
                fontFamily: 'wf_segoe-ui_Semibold'
            }
        };
        private static Caption: ClassAndSelector = {
            class: 'caption',
            selector: '.caption'
        };

        private toolTip: D3.Selection;
        private animationOptions: AnimationOptions;
        private displayUnitSystemType: DisplayUnitSystemType;
        private isScrollable: boolean;
        private graphicsContext: D3.Selection;
        private labelContext: D3.Selection;
        private cardFormatSetting: CardFormatSetting;
        private kpiImage: D3.Selection;

        public constructor(options?: CardConstructorOptions) {
            super(Card.cardClassName);
            this.isScrollable = false;
            this.displayUnitSystemType = DisplayUnitSystemType.WholeUnits;

            if (options) {
                this.isScrollable = !!options.isScrollable;
                if (options.animator)
                    this.animator = options.animator;
                if (options.displayUnitSystemType != null)
                    this.displayUnitSystemType = options.displayUnitSystemType;
            }
        }

        public init(options: VisualInitOptions) {
            debug.assertValue(options, 'options');
            this.animationOptions = options.animation;
            let element = options.element;

            this.kpiImage = d3.select(element.get(0)).append('div');
            let svg = this.svg = d3.select(element.get(0)).append('svg');
            this.graphicsContext = svg.append('g');
            this.currentViewport = options.viewport;
            this.hostServices = options.host;
            this.style = options.style;

            this.updateViewportProperties();

            if (this.isScrollable) {
                svg.attr('class', Card.cardClassName);
                this.labelContext = svg.append('g');
            }
        }

        public onDataChanged(options: VisualDataChangedOptions): void {
            debug.assertValue(options, 'options');

            //Default settings for reset to default
            this.cardFormatSetting = this.getDefaultFormatSettings();

            let dataView = options.dataViews[0];
            let value: any;
            if (dataView) {
                this.getMetaDataColumn(dataView);
                if (dataView.single) {
                    value = dataView.single.value;
                }

                let dataViewMetadata = dataView.metadata;
                if (dataViewMetadata) {
                    let objects: DataViewObjects = dataViewMetadata.objects;
                    if (objects) {
                        let labelSettings = this.cardFormatSetting.labelSettings;

                        labelSettings.labelColor = DataViewObjects.getFillColor(objects, cardProps.labels.color, labelSettings.labelColor);
                        labelSettings.precision = DataViewObjects.getValue(objects, cardProps.labels.labelPrecision, labelSettings.precision);

                        // The precision can't go below 0
                        if (labelSettings.precision != null) {
                            labelSettings.precision = (labelSettings.precision >= 0) ? labelSettings.precision : 0;
                        }

                        labelSettings.displayUnits = DataViewObjects.getValue(objects, cardProps.labels.labelDisplayUnits, labelSettings.displayUnits);
                        this.cardFormatSetting.showTitle = DataViewObjects.getValue(objects, cardProps.cardTitle.show, this.cardFormatSetting.showTitle);
                        this.cardFormatSetting.wordWrap = DataViewObjects.getValue(objects, cardProps.wordWrap.show, this.cardFormatSetting.wordWrap);
                    }
                }
            }

            this.updateInternal(value, true /* suppressAnimations */, true /* forceUpdate */);
        }

        public onResizing(viewport: IViewport): void {
            this.currentViewport = viewport;
            this.updateViewportProperties();
            this.updateInternal(this.value, true /* suppressAnimations */, true /* forceUpdate */);
        }

        private updateViewportProperties() {
            let viewport = this.currentViewport;
            this.svg.attr('width', viewport.width)
                .attr('height', viewport.height);
        }

        public getAdjustedFontHeight(availableWidth: number, textToMeasure: string, seedFontHeight: number) {
            let adjustedFontHeight = super.getAdjustedFontHeight(availableWidth, textToMeasure, seedFontHeight);

            return Math.min(adjustedFontHeight, Card.DefaultStyle.card.maxFontSize);
        }

        public clear(valueOnly: boolean = false) {
            this.svg.select(Card.Value.selector).text('');

            if (!valueOnly)
                this.svg.select(Card.Label.selector).text('');

            super.clear();
        }

        private updateInternal(target: any, suppressAnimations: boolean, forceUpdate: boolean = false) {
            let start = this.value;
            let duration = AnimatorCommon.GetAnimationDuration(this.animator, suppressAnimations);

            if (target === undefined) {
                if (start !== undefined)
                    this.clear();
                return;
            }

            let metaDataColumn = this.metaDataColumn;
            let labelSettings = this.cardFormatSetting.labelSettings;
            let labelStyles = Card.DefaultStyle.label;
            let valueStyles = Card.DefaultStyle.value;
            let isDefaultDisplayUnit = labelSettings.displayUnits === 0;
            let formatter = valueFormatter.create({
                format: this.getFormatString(metaDataColumn),
                value: isDefaultDisplayUnit ? target : labelSettings.displayUnits,
                precision: labelSettings.precision,
                displayUnitSystemType: isDefaultDisplayUnit && labelSettings.precision === 0 ? this.displayUnitSystemType : DisplayUnitSystemType.WholeUnits, // keeps this.displayUnitSystemType as the displayUnitSystemType unless the user changed the displayUnits or the precision
                formatSingleValues: isDefaultDisplayUnit ? true : false,
                allowFormatBeautification: true,
                columnType: metaDataColumn ? metaDataColumn.type : undefined
            });
            let formatSettings = this.cardFormatSetting;
            let width = this.currentViewport.width;
            let height = this.currentViewport.height;
            let translateX = this.getTranslateX(width);
            let translateY = (height - labelStyles.height - valueStyles.fontSize) / 2;

            let statusGraphicInfo = getKpiImageMetadata(metaDataColumn, target, KpiImageSize.Big);
            let columnCaption: string;
            let statusGraphic: string;

            if (statusGraphicInfo) {
                columnCaption = statusGraphicInfo.caption;
                statusGraphic = statusGraphicInfo.statusGraphic;
            }

            if (!columnCaption)
                columnCaption = valueFormatter.format(target, valueFormatter.getFormatString(metaDataColumn, MultiRowCard.formatStringProp));

            if (this.isScrollable) {
                if (!forceUpdate && start === target)
                    return;

                let label: string;

                if (start !== target) {
                    target = formatter.format(target);
                }

                if (metaDataColumn)
                    label = metaDataColumn.displayName;

                if (statusGraphic) {
                    // show KPI
                    this.graphicsContext.selectAll('text').remove();
                    this.displayStatusGraphic(statusGraphic, translateX, translateY, columnCaption, valueStyles);
                } else {
                    // show text value
                    this.kpiImage.selectAll('div').remove();
                    let valueElement = this.graphicsContext
                        .attr('transform', SVGUtil.translate(translateX, this.getTranslateY(valueStyles.fontSize + translateY)))
                        .selectAll('text')
                        .data([target]);

                    valueElement
                        .enter()
                        .append('text')
                        .attr('class', Card.Value.class);

                    valueElement
                        .text((d: any) => d)
                        .style({
                            'font-size': valueStyles.fontSize + 'px',
                            'fill': labelSettings.labelColor,
                            'font-family': valueStyles.fontFamily,
                            'text-anchor': this.getTextAnchor(),
                        });

                    valueElement.call(AxisHelper.LabelLayoutStrategy.clip,
                        width,
                        TextMeasurementService.svgEllipsis);

                    valueElement.exit().remove();
                }

                let labelData = formatSettings.showTitle
                    ? [label]
                    : [];

                let translatedLabelY = this.getTranslateY(valueStyles.fontSize + labelStyles.height + translateY);
                let labelElement = this.labelContext
                    .attr('transform', SVGUtil.translate(translateX, translatedLabelY))
                    .selectAll('text')
                    .data(labelData);

                labelElement
                    .enter()
                    .append('text')
                    .attr('class', Card.Label.class);

                labelElement
                    .text((d) => d)
                    .style({
                        'font-size': labelStyles.fontSize + 'px',
                        'fill': labelStyles.color,
                        'text-anchor': this.getTextAnchor()
                    });

                labelElement.exit().remove();

                let labelElementNode = <SVGTextElement>labelElement.node();
                if (labelElementNode) {
                    if (formatSettings.wordWrap)
                        TextMeasurementService.wordBreak(labelElementNode, width / 2, height - translatedLabelY);
                    else
                        labelElement.call(AxisHelper.LabelLayoutStrategy.clip,
                            width,
                            TextMeasurementService.svgEllipsis);
                }
            }
            else {
                if (statusGraphic) {
                    this.graphicsContext.selectAll('text').remove();
                    this.displayStatusGraphic(statusGraphic, translateX, translateY, columnCaption, valueStyles);
                } else {
                    this.kpiImage.selectAll('div').remove();
                    this.doValueTransition(
                        start,
                        target,
                        this.displayUnitSystemType,
                        this.animationOptions,
                        duration,
                        forceUpdate,
                        formatter);
                }
            }

            this.updateTooltip(target);
            this.value = target;
        }

        private displayStatusGraphic(statusGraphic: string, translateX: number, translateY: number, columnCaption: string, valueStyles) {
            let imageHeight = 40;

            if (statusGraphic) {
                this.kpiImage.style({
                    'display': 'block',
                });

                let captionDiv = this.kpiImage.select('div');
                if (!captionDiv || captionDiv.empty()) {
                    captionDiv = this.kpiImage.append('div');
                    captionDiv.classed(Card.Caption.class, true)
                        .style({
                            'text-align': 'left',
                            'white-space': 'nowrap',
                            'text-overflow': 'ellipsis',
                            'overflow': 'hidden',
                            'position': 'absolute'
                        });
                    captionDiv.append('div')
                        .style({
                            display: 'inline-block',
                            verticalAlign: 'sub'
                        });
                }

                captionDiv.style({
                    'transform': SVGUtil.translateWithPixels((translateX - (imageHeight / 2)), this.getTranslateY(valueStyles.fontSize + translateY) - imageHeight)
                });

                captionDiv.classed(columnCaption, true);
            }
            else {
                this.kpiImage.selectAll('div').remove();
            }
        }

        private updateTooltip(target: number) {
            if (!this.toolTip)
                this.toolTip = this.graphicsContext.append("svg:title");
            this.toolTip.text(target);
        }

        private getDefaultFormatSettings(): CardFormatSetting {
            return {
                showTitle: true,
                labelSettings: dataLabelUtils.getDefaultLabelSettings(/* showLabel: */true, Card.DefaultStyle.value.color, undefined),
                wordWrap: false,
            };
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
            if (!this.cardFormatSetting)
                this.cardFormatSetting = this.getDefaultFormatSettings();

            let formatSettings = this.cardFormatSetting;
            let enumeration = new ObjectEnumerationBuilder();

            switch (options.objectName) {
                case 'cardTitle':
                    enumeration.pushInstance({
                        objectName: 'cardTitle',
                        selector: null,
                        properties: {
                            show: formatSettings.showTitle,
                        },
                    });
                    break;
                case 'labels':
                    let labelSettingOptions: VisualDataLabelsSettingsOptions = {
                        enumeration: enumeration,
                        dataLabelsSettings: formatSettings.labelSettings,
                        show: true,
                        displayUnits: true,
                        precision: true,
                    };
                    dataLabelUtils.enumerateDataLabels(labelSettingOptions);
                    break;
                case 'wordWrap':
                    enumeration.pushInstance({
                        objectName: 'wordWrap',
                        selector: null,
                        properties: {
                            show: formatSettings.wordWrap,
                        },
                    });
                    break;
            }

            return enumeration.complete();
        }
    }
}