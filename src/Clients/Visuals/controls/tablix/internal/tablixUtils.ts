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

/// <reference path="../../../_references.ts"/>

module powerbi.visuals.controls {
    export module HTMLElementUtils {
        export function clearChildren(element: HTMLElement): void {
            if (!element) {
                return;
            }

            while (element.hasChildNodes()) {
                element.removeChild(element.firstChild);
            }
        }

        export function setElementTop(element: HTMLElement, top: number): void {
            element.style.top = top + "px";
        }

        export function setElementLeft(element: HTMLElement, left: number): void {
            element.style.left = left + "px";
        }

        export function setElementHeight(element: HTMLElement, height: number): void {
            if (HTMLElementUtils.isAutoSize(height))
                element.style.height = "auto";
            else
                element.style.height = height + "px";
        }

        export function setElementWidth(element: HTMLElement, width: number): void {
            if (HTMLElementUtils.isAutoSize(width))
                element.style.width = "auto";
            else
                element.style.width = width + "px";
        }

        export function getElementWidth(element: HTMLElement): number {
            return element.offsetWidth;
        }

        export function getElementHeight(element: HTMLElement): number {
            return element.offsetHeight;
        }

        export function isAutoSize(size: number): boolean {
            return size === -1;
        }

        export function getAccumulatedScale(element: HTMLElement): number {
            let scale: number = 1;
            while (element) {
                scale *= HTMLElementUtils.getScale(element);
                element = element.parentElement;
            }

            return scale;
        }
        
        /**
         * Get scale of element, return 1 when not scaled.
         */
        export function getScale(element: any): number {
            element = $(element);

            let str = element.css('-webkit-transform') ||
                element.css('-moz-transform') ||
                element.css('-ms-transform') ||
                element.css('-o-transform') ||
                element.css('transform');

            return (str && (
                str.match(/\d*\.\d*/) && Number(str.match(/\d*\.\d*/)[0]) ||
                str.match(/\d+/) && Number(str.match(/\d+/)[0]))
                ) || 1;
        }
    }
}

module powerbi.visuals.controls.internal {
    import DomFactory = InJs.DomFactory;

    export module TablixUtils {

        export function createTable(): HTMLTableElement {
            return <HTMLTableElement>document.createElement("table");
        }

        export function createDiv(): HTMLDivElement {
            let div: HTMLDivElement = <HTMLDivElement>document.createElement("div");

            // TODO: Fold these into CSS as well combined with the styling done for the different scenarios where div are used.
            let divStyle = div.style;
            divStyle.whiteSpace = "nowrap";
            divStyle.overflow = "hidden";
            divStyle.lineHeight = "normal";

            return div;
        }

        export function appendATagToBodyCell(value: string, cell: controls.ITablixCell): void {
            let element = <HTMLElement>cell.extension.contentHost;
            let atag: HTMLAnchorElement = null;
            if(element.childElementCount === 0) {
                atag = document.createElement('a');
                element.appendChild(atag);
            } else {
                atag = <HTMLAnchorElement>element.children[0];
            }

            atag.href = value;
            atag.target = '_blank';
            atag.title = value;
            atag.innerText = value;
        }

        export function appendImgTagToBodyCell(value: string, cell: controls.ITablixCell): void {
            var element = <HTMLElement>cell.extension.contentHost;
            var contentElement = element.parentElement;
            var imgTag: HTMLImageElement;
            if (element.childElementCount === 0) {
                imgTag = document.createElement('img');
                element.appendChild(imgTag);
            } else {
                imgTag = <HTMLImageElement>element.children[0];
            }
            // set padding for contentElement
            contentElement.style.paddingBottom = '3px';
            contentElement.style.paddingTop = '3px';
            imgTag.src = value;
            imgTag.style.maxHeight = '75px';
            imgTag.style.maxWidth = '100px';
            imgTag.style.height = '100%';
        }

        export function createKpiDom(kpiStatusGraphic: string, kpiValue: string): JQuery {
            debug.assertValue(kpiStatusGraphic, 'kpiStatusGraphic');
            debug.assertValue(kpiValue, 'kpiValue');
            let className: string = KpiUtil.getClassForKpi(kpiStatusGraphic, kpiValue) || '';
            return DomFactory.div()
                .addClass(className)
                .css({
                    'display': 'inline-block',
                    'vertical-align': 'sub'
                });
        }

        export function isValidStatusGraphic(kpiStatusGraphic: string, kpiValue: string): boolean {
            if (!kpiStatusGraphic || kpiValue === undefined) {
                return false;
            }

            return !!KpiUtil.getClassForKpi(kpiStatusGraphic, kpiValue);
        }
    }
}