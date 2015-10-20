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

module powerbi.data {
    /** Defines a selector for content, including data-, metadata, and user-defined repetition. */
    export interface Selector {
        /** Data-bound repetition selection. */
        data?: DataRepetitionSelector[];
	
        /** Metadata-bound repetition selection.  Refers to a DataViewMetadataColumn queryName. */
        metadata?: string;

        /** User-defined repetition selection. */
        id?: string;
    }

    /* tslint:disable:no-unused-expression */
    export type DataRepetitionSelector = DataViewScopeIdentity | DataViewScopeWildcard;
    /* tslint:enable */

    export module Selector {
        import ArrayExtensions = jsCommon.ArrayExtensions;

        export function filterFromSelector(selectors: Selector[], isNot?: boolean): SemanticFilter {
            if (ArrayExtensions.isUndefinedOrEmpty(selectors))
                return;

            let expr: SQExpr;
            for (let i = 0, ilen = selectors.length; i < ilen; i++) {
                let identity = selectors[i];
                let data = identity.data;
                let exprToAdd: SQExpr = undefined;
                if (data && data.length) {
                    for (let j = 0, jlen = data.length; j < jlen; j++) {
                        exprToAdd = SQExprBuilder.and(exprToAdd, (<DataViewScopeIdentity>identity.data[j]).expr);
                    }
                }

                expr = SQExprBuilder.or(expr, exprToAdd);
            }

            if (expr && isNot)
                expr = SQExprBuilder.not(expr);

            return SemanticFilter.fromSQExpr(expr);
        }

        export function matchesData(selector: Selector, identities: DataViewScopeIdentity[]): boolean {
            debug.assertValue(selector, 'selector');
            debug.assertValue(selector.data, 'selector.data');
            debug.assertValue(identities, 'identities');

            let selectorData = selector.data;
            if (selectorData.length !== identities.length)
                return false;

            for (let i = 0, len = selectorData.length; i < len; i++) {
                let dataItem = selector.data[i];
                let selectorDataItem = <DataViewScopeIdentity>dataItem;
                if (selectorDataItem.expr) {
                    if (!DataViewScopeIdentity.equals(selectorDataItem, identities[i]))
                        return false;
                }
                else {
                    if (!DataViewScopeWildcard.matches(<DataViewScopeWildcard>dataItem, identities[i]))
                        return false;
                }
            }

            return true;
        }

        export function matchesKeys(selector: Selector, keysList: SQExpr[][]): boolean {
            debug.assertValue(selector, 'selector');
            debug.assertValue(selector.data, 'selector.data');
            debug.assertValue(keysList, 'keysList');

            let selectorData = selector.data,
                selectorDataLength = selectorData.length;
            if (selectorDataLength !== keysList.length)
                return false;

            for (let i = 0; i < selectorDataLength; i++) {
                let selectorDataItem = selector.data[i],
                    selectorDataExprs: SQExpr[];

                if ((<DataViewScopeIdentity>selectorDataItem).expr) {
                    selectorDataExprs = ScopeIdentityKeyExtractor.run((<DataViewScopeIdentity>selectorDataItem).expr);
                }
                else {
                    selectorDataExprs = (<DataViewScopeWildcard>selectorDataItem).exprs;
                }

                if (!selectorDataExprs)
                    continue;
                if (!SQExprUtils.sequenceEqual(keysList[i], selectorDataExprs))
                    return false;
            }

            return true;
        }

        /** Determines whether two selectors are equal. */
        export function equals(x: Selector, y: Selector): boolean {
            // Normalize falsy to null
            x = x || null;
            y = y || null;

            if (x === y)
                return true;

            if (!x !== !y)
                return false;

            debug.assertValue(x, 'x');
            debug.assertValue(y, 'y');

            if (x.id !== y.id)
                return false;
            if (x.metadata !== y.metadata)
                return false;
            if (!equalsDataArray(x.data, y.data))
                return false;

            return true;
        }

        function equalsDataArray(x: DataRepetitionSelector[], y: DataRepetitionSelector[]): boolean {
            // Normalize falsy to null
            x = x || null;
            y = y || null;

            if (x === y)
                return true;

            if (!x !== !y)
                return false;

            if (x.length !== y.length)
                return false;

            for (let i = 0, len = x.length; i < len; i++) {
                if (!equalsData(x[i], y[i]))
                    return false;
            }

            return true;
        }

        function equalsData(x: DataRepetitionSelector, y: DataRepetitionSelector): boolean {
            if (!(<DataViewScopeIdentity>x).expr && (<DataViewScopeIdentity>y).expr) {
                // TODO: We need to also check wildcard selectors too (once that's supported/figured out).
                return false;
            }

            return DataViewScopeIdentity.equals(<DataViewScopeIdentity>x, <DataViewScopeIdentity>y);
        }

        export function getKey(selector: Selector): string {
            let toStringify: any = {};
            if (selector.data) {
                let data = [];
                for (let i = 0, ilen = selector.data.length; i < ilen; i++) {
                    data.push(selector.data[i].key);
                }
                toStringify.data = data;
            }
            if (selector.metadata)
                toStringify.metadata = selector.metadata;
            if (selector.id)
                toStringify.id = selector.id;
            return JSON.stringify(toStringify);
        }

        export function containsWildcard(selector: Selector): boolean {
            debug.assertValue(selector, 'selector');

            let dataItems = selector.data;
            if (!dataItems)
                return false;

            for (let i = 0, len = dataItems.length; i < len; i++) {
                let wildcard = <DataViewScopeWildcard>dataItems[i];
                if (wildcard.exprs)
                    return true;
            }

            return false;
        }
    }
}