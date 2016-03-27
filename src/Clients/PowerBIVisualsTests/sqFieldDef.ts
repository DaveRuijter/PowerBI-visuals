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

// TODO: this file should be removed.
module powerbi.data {
    /** Represents a simplified table aggregate/column/column aggregate reference within a SQ. */
    export interface SQFieldDef {
        schema: string;
        entity: string;
        column?: string;
        measure?: string;
        aggregate?: data.QueryAggregateFunction;
        entityVar?: string;
    }

    export module SQExprConverter {
        export function asSQFieldDef(sqExpr: SQExpr): SQFieldDef {
            return sqExpr.accept(SQFieldDefinitionBuilder.instance);
        }
    }

    class SQFieldDefinitionBuilder extends powerbi.data.DefaultSQExprVisitor<SQFieldDef> {
        public static instance: SQFieldDefinitionBuilder = new SQFieldDefinitionBuilder();

        public visitColumnRef(expr: SQColumnRefExpr): SQFieldDef {
            let sourceRef: SQFieldDef = expr.source.accept(this);
            if (sourceRef) {
                sourceRef.column = expr.ref;
                return sourceRef;
            }
        }

        public visitMeasureRef(expr: SQMeasureRefExpr): SQFieldDef {
            let sourceRef: SQFieldDef = expr.source.accept(this);
            if (sourceRef) {
                sourceRef.measure = expr.ref;
                return sourceRef;
            }
        }

        public visitAggr(expr: SQAggregationExpr): SQFieldDef {
            let sourceRef: SQFieldDef = expr.arg.accept(this);
            if (sourceRef) {
                sourceRef.aggregate = expr.func;
                return sourceRef;
            }
        }

        public visitEntity(expr: SQEntityExpr): SQFieldDef {
            let fieldDef: SQFieldDef = {
                schema: expr.schema,
                entity: expr.entity
            };
            if (expr.variable)
                fieldDef.entityVar = expr.variable;

            return fieldDef;
        }
    }

    export module SQExprBuilder {
        export function fieldDef(fieldDef: SQFieldDef): SQExpr {
            return wrapAggr(fieldDef)
                || wrapColumn(fieldDef)
                || wrapMeasure(fieldDef)
                || wrapEntity(fieldDef);
        }

        function wrapAggr(fieldDef: SQFieldDef): SQExpr {
            let aggr = fieldDef.aggregate;
            if (aggr !== undefined) {
                let expr = wrapColumn(fieldDef) || wrapEntity(fieldDef);
                if (expr)
                    return SQExprBuilder.aggregate(expr, aggr);

            }
        }

        function wrapColumn(fieldDef: SQFieldDef): SQExpr {
            let column = fieldDef.column;
            if (column) {
                let entityExpr = wrapEntity(fieldDef);
                if (entityExpr)
                    return columnRef(entityExpr, column);
            }
        }

        function wrapMeasure(fieldDef: SQFieldDef): SQExpr {
            let measure = fieldDef.measure;
            if (measure) {
                let entityExpr = wrapEntity(fieldDef);
                if (entityExpr)
                    return measureRef(entityExpr, measure);
            }
        }

        function wrapEntity(fieldDef: SQFieldDef): SQExpr {
            return entity(fieldDef.schema, fieldDef.entity, fieldDef.entityVar);
        }
    }
}