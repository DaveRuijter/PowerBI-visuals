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

module powerbi.data {
    /** Represents common expression patterns for 'field' expressions such as columns, column aggregates, measures, etc. */
    export interface FieldExprPattern {
        column?: FieldExprColumnPattern;
        columnAggr?: FieldExprColumnAggrPattern;
        columnHierarchyLevelVariation?: FieldExprColumnHierarchyLevelVariation;
        entityAggr?: FieldExprEntityAggrPattern;
        hierarchyLevel?: FieldExprHierarchyLevelPattern;
        hierarchy?: FieldExprHierarchyPattern;
        measure?: FieldExprMeasurePattern;
    }

    export interface FieldExprEntityItemPattern {
        schema: string;
        entity: string;
        entityVar?: string;
    }

    export interface FieldExprPropertyPattern extends FieldExprEntityItemPattern {
        name: string;
    }

    export type FieldExprColumnPattern = FieldExprPropertyPattern;

    export interface FieldExprColumnAggrPattern extends FieldExprColumnPattern {
        aggregate: QueryAggregateFunction;
    }

    export module SQExprBuilder {
        export function fieldExpr(fieldExpr: FieldExprPattern): SQExpr {
            return wrapColumnAggr(fieldExpr)
                || wrapColumn(fieldExpr)
                || wrapMeasure(fieldExpr)
                || wrapHierarchyLevel(fieldExpr)
                || wrapHierarchy(fieldExpr)
                || wrapEntityAggr(fieldExpr)
                || wrapPropertyVariationSource(fieldExpr)
                || wrapEntity(fieldExpr);
        }

        function wrapColumnAggr(fieldExpr: FieldExprPattern): SQExpr {
            var aggr = fieldExpr.columnAggr;
            if (aggr) {
                var entityExpr = wrapEntity(fieldExpr.columnAggr);
                return aggregate(columnRef(entityExpr, aggr.name), aggr.aggregate);
            }
        }

        function wrapHierarchyLevel(fieldExpr: FieldExprPattern): SQExpr {
            let hierarchyLevelPattern = fieldExpr.hierarchyLevel;
            if (hierarchyLevelPattern) {
                let hierarchyExpr = hierarchy(wrapEntity(hierarchyLevelPattern), hierarchyLevelPattern.name);
                return hierarchyLevel(hierarchyExpr, hierarchyLevelPattern.level);
            }
        }

        function wrapHierarchy(fieldExpr: FieldExprPattern): SQExpr {
            let hierarchyExprPattern = fieldExpr.hierarchy;
            if (hierarchyExprPattern) {
                let entityExpr = wrapEntity(hierarchyExprPattern);
                return hierarchy(entityExpr, hierarchyExprPattern.name);
            }
        }

        function wrapPropertyVariationSource(fieldExpr: FieldExprPattern): SQExpr {
            let variation = fieldExpr.columnHierarchyLevelVariation;
            if (variation) {
                let entitiyExpr = wrapEntity(variation.source);
                return propertyVariationSource(entitiyExpr, variation.source.name, variation.level.name);
            }
        }

        function wrapColumn(fieldExpr: FieldExprPattern): SQExpr {
            var column = fieldExpr.column;
            if (column) {
                var entityExpr = wrapEntity(fieldExpr.column);
                return columnRef(entityExpr, column.name);
            }
        }

        function wrapMeasure(fieldExpr: FieldExprPattern): SQExpr {
            var measure = fieldExpr.measure;
            if (measure) {
                var entityExpr = wrapEntity(fieldExpr.measure);
                return measureRef(entityExpr, measure.name);
            }
        }

        function wrapEntityAggr(fieldExpr: FieldExprPattern): SQExpr {
            var entityAggregate = fieldExpr.entityAggr;

            if (entityAggregate) {
                var entityExpr = wrapEntity(fieldExpr.entityAggr);
                return aggregate(entityExpr, entityAggregate.aggregate);
            }
        }

        function wrapEntity(fieldExpr: FieldExprPattern): SQExpr {
            let fieldExprEntityItemPattern = FieldExprPattern.toFieldExprEntityItemPattern(fieldExpr);
            return entity(fieldExprEntityItemPattern.schema, fieldExprEntityItemPattern.entity, fieldExprEntityItemPattern.entityVar);
        }
    }

    export interface FieldExprColumnHierarchyLevelVariation {
        source: FieldExprColumnPattern;
        level: FieldExprHierarchyLevelPattern;
        variationName: string;
    }

    export interface FieldExprEntityAggrPattern extends FieldExprEntityItemPattern {
        aggregate: QueryAggregateFunction;
    }

    export interface FieldExprHierarchyLevelPattern extends FieldExprEntityItemPattern {
        level: string;
        name: string;
    }

    export interface FieldExprHierarchyPattern extends FieldExprEntityItemPattern {
        name: string;
    }
    export type FieldExprMeasurePattern = FieldExprPropertyPattern;

    export module SQExprConverter {
        export function asFieldPattern(sqExpr: SQExpr): FieldExprPattern {
            // TODO: adding entity to the FieldExprPattern
            if (sqExpr instanceof data.SQEntityExpr) {
                return {
                    entity: sqExpr.entity,
                    schema: sqExpr.schema,
                };
            }

            return sqExpr.accept(FieldExprPatternBuilder.instance);
        }
    }

    interface FieldExprSourceVariationPattern {
        column: FieldExprColumnPattern;
        variationName: string;
    }

    interface SourceExprPattern {
        entity?: FieldExprEntityItemPattern;
        // TODO: Change FieldExprHierarchyPattern to FieldExprHierarchyLevelPattern
        hierarchy?: FieldExprHierarchyPattern;
        variation?: FieldExprSourceVariationPattern;
    }

    interface HierarchySourceExprPattern {
        hierarchy: FieldExprHierarchyPattern;
        variation?: FieldExprSourceVariationPattern;
    }

    class FieldExprPatternBuilder extends DefaultSQExprVisitor<FieldExprPattern> {
        public static instance: FieldExprPatternBuilder = new FieldExprPatternBuilder();

        public visitColumnRef(expr: SQColumnRefExpr): FieldExprPattern {
            let sourceRef = expr.source.accept(SourceExprPatternBuilder.instance);
            if (!sourceRef)
                return;

            if (sourceRef.entity) {
                let columnRef = <FieldExprColumnPattern>sourceRef.entity;
                columnRef.name = expr.ref;

                return { column: columnRef };
            }
        }

        public visitMeasureRef(expr: SQMeasureRefExpr): FieldExprPattern {
            let sourceRef = expr.source.accept(SourceExprPatternBuilder.instance);
            if (!sourceRef)
                return;

            if (sourceRef.entity) {
                let measureRef = <FieldExprMeasurePattern>sourceRef.entity;
                measureRef.name = expr.ref;

                return { measure: measureRef };
            }
        }

        public visitAggr(expr: SQAggregationExpr): FieldExprPattern {
            let fieldPattern: FieldExprPattern = expr.arg.accept(this);
            if (fieldPattern && fieldPattern.column) {
                let argAggr = <FieldExprColumnAggrPattern>fieldPattern.column;
                argAggr.aggregate = expr.func;
                return { columnAggr: argAggr };
            } else if (fieldPattern && fieldPattern.columnAggr) {
                let argAggr = <FieldExprColumnAggrPattern>fieldPattern.columnAggr;
                argAggr.aggregate = expr.func;
                return { columnAggr: argAggr };
            }

            let sourcePattern = expr.arg.accept(SourceExprPatternBuilder.instance);
            if (sourcePattern && sourcePattern.entity) {
                let argAggr = <FieldExprEntityAggrPattern>sourcePattern.entity;
                argAggr.aggregate = expr.func;

                return { entityAggr: argAggr };
            }
        }

        public visitHierarchy(expr: SQHierarchyExpr): FieldExprPattern {
            let sourcePattern = expr.arg.accept(SourceExprPatternBuilder.instance);

            if (sourcePattern && sourcePattern.entity) {
                let hierarchyRef = <FieldExprHierarchyPattern>(sourcePattern.entity);
                hierarchyRef.name = expr.hierarchy;
                return { hierarchy: hierarchyRef };
            }
        }

        public visitHierarchyLevel(expr: SQHierarchyLevelExpr): FieldExprPattern {
            let hierarchySourceExprPattern: HierarchySourceExprPattern = expr.arg.accept(HierarchyExprPatternBuiler.instance);
            if (!hierarchySourceExprPattern)
                return;

            let hierarchyLevel: FieldExprHierarchyLevelPattern;
            if (hierarchySourceExprPattern.hierarchy) {
                hierarchyLevel = {
                    entity: hierarchySourceExprPattern.hierarchy.entity,
                    schema: hierarchySourceExprPattern.hierarchy.schema,
                    name: hierarchySourceExprPattern.hierarchy.name,
                    level: expr.level,
                };
            }

            if (hierarchySourceExprPattern.variation) {
                return {
                    columnHierarchyLevelVariation: {
                        source: {
                            entity: hierarchySourceExprPattern.variation.column.entity,
                            schema: hierarchySourceExprPattern.variation.column.schema,
                            name: hierarchySourceExprPattern.variation.column.name,
                        },
                        level: hierarchyLevel,
                        variationName: hierarchySourceExprPattern.variation.variationName,
                    }
                };
            }

            return { hierarchyLevel: hierarchyLevel };
        }
    }

    class SourceExprPatternBuilder extends DefaultSQExprVisitor<SourceExprPattern> {
        public static instance: SourceExprPatternBuilder = new SourceExprPatternBuilder();

        public visitEntity(expr: SQEntityExpr): SourceExprPattern {
            let entityRef: FieldExprEntityItemPattern = {
                schema: expr.schema,
                entity: expr.entity
            };
            if (expr.variable)
                entityRef.entityVar = expr.variable;

            return { entity: entityRef };
        }

        public visitPropertyVariationSource(expr: SQPropertyVariationSourceExpr): SourceExprPattern {
            let entityExpr = <SQEntityExpr>expr.arg;

            if (entityExpr instanceof SQEntityExpr) {
                let propertyVariationSource: FieldExprColumnPattern = {
                    schema: entityExpr.schema,
                    entity: entityExpr.entity,
                    name: expr.property,
                };

                if (entityExpr.variable)
                    propertyVariationSource.entityVar = entityExpr.variable;

                return {
                    variation: {
                        column: propertyVariationSource,
                        variationName: expr.name,
                    }
                };
            }
        }
    }

    class HierarchyExprPatternBuiler extends DefaultSQExprVisitor<HierarchySourceExprPattern> {
        public static instance: HierarchyExprPatternBuiler = new HierarchyExprPatternBuiler();

        public visitHierarchy(expr: SQHierarchyExpr): HierarchySourceExprPattern {
            let exprPattern = expr.arg.accept(SourceExprPatternBuilder.instance);
            let hierarchyRef: FieldExprHierarchyPattern;
            let variationRef: FieldExprSourceVariationPattern;

            if (exprPattern.variation) {
                hierarchyRef = {
                    name: expr.hierarchy,
                    schema: exprPattern.variation.column.schema,
                    entity: exprPattern.variation.column.entity,
                };
                variationRef = exprPattern.variation;
            }
            else
                hierarchyRef = {
                    name: expr.hierarchy,
                    schema: exprPattern.entity.schema,
                    entity: exprPattern.entity.entity,
                };

            return {
                hierarchy: hierarchyRef,
                variation: variationRef
            };
        }
    }

    export module FieldExprPattern {
        export function hasFieldExprName(fieldExpr: FieldExprPattern): boolean {
            return (fieldExpr.column ||
                fieldExpr.columnAggr ||
                fieldExpr.measure) !== undefined;
        }

        export function getPropertyName(fieldExpr: FieldExprPattern): string {
            let column = (fieldExpr.column ||
                fieldExpr.columnAggr ||
                fieldExpr.measure ||
                fieldExpr.hierarchy);

            if (column)
                return column.name;
            else if (fieldExpr.hierarchyLevel)
                return fieldExpr.hierarchyLevel.level;
        }

        export function getFieldExprName(fieldExpr: FieldExprPattern): string {
            let name = getPropertyName(fieldExpr);

            if (name)
                return name;

            // In case it is an entity
            return toFieldExprEntityItemPattern(fieldExpr).entity;
        }

        export function toFieldExprEntityItemPattern(fieldExpr: FieldExprPattern): FieldExprEntityItemPattern {
            var field = <FieldExprEntityItemPattern>(fieldExpr.column ||
                fieldExpr.columnAggr ||
                fieldExpr.entityAggr ||
                fieldExpr.hierarchy ||
                fieldExpr.hierarchyLevel ||
                fieldExpr.measure ||
                (fieldExpr.columnHierarchyLevelVariation && fieldExpr.columnHierarchyLevelVariation.source) ||
                fieldExpr); // fieldExpr for entity

            return {
                schema: field.schema,
                entity: field.entity,
                entityVar: field.entityVar,
            };
        }
    }
}