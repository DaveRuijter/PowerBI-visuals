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
    import StringExtensions = jsCommon.StringExtensions;

    module SQExprHierarchyToHierarchyLevelConverter {
        export function convert(sqExpr: SQExpr, federatedSchema: FederatedConceptualSchema): SQExpr[] {
            debug.assertValue(sqExpr, 'sqExpr');
            debug.assertValue(federatedSchema, 'federatedSchema');

            if (sqExpr instanceof SQHierarchyExpr) {
                let hierarchyExpr = <SQHierarchyExpr>sqExpr;

                let conceptualHierarcy = SQExprUtils.getConceptualHierarchy(hierarchyExpr, federatedSchema);
                return _.map(conceptualHierarcy.levels, hierarchyLevel => SQExprBuilder.hierarchyLevel(sqExpr, hierarchyLevel.name));
            }
        }
    }

    module SQExprVariationConverter {
        export function expand(expr: SQExpr, schema: FederatedConceptualSchema): SQExpr[] {
            debug.assertValue(expr, 'sqExpr');
            debug.assertValue(schema, 'federatedSchema');

            let exprs: SQExpr[];
            let conceptualProperty = expr.getConceptualProperty(schema);

            if (conceptualProperty) {
                let column = conceptualProperty.column;
                if (column && column.variations && column.variations.length > 0) {
                    let variations = column.variations;
                    // for SU10, we support only one variation
                    debug.assert(variations.length === 1, "variations.length");
                    let variation = variations[0];

                    let columnExpr = <SQColumnRefExpr>expr;
                    let entityExpr = <SQEntityExpr>columnExpr.source;

                    exprs = [];
                    if (variation.defaultHierarchy) {
                        let hierarchyExpr = SQExprBuilder.hierarchy(
                            SQExprBuilder.propertyVariationSource(
                                SQExprBuilder.entity(entityExpr.schema, entityExpr.entity),
                                variation.name, conceptualProperty.name),
                            variation.defaultHierarchy.name);

                        for (let level of variation.defaultHierarchy.levels)
                            exprs.push(SQExprBuilder.hierarchyLevel(hierarchyExpr, level.name));
                    }
                }
            }

            return exprs;
        }
    }

    export module SQExprUtils {
        /** Returns an array of supported aggregates for a given expr and role. */
        export function getSupportedAggregates(
            expr: SQExpr,
            isGroupingOnly: boolean,
            schema: FederatedConceptualSchema): QueryAggregateFunction[] {
            let emptyList: QueryAggregateFunction[] = [];

            let metadata = getMetadataForUnderlyingType(expr, schema);
            // don't use expr.validate as validate will be using this function and we end up in a recursive loop
            if (!metadata)
                return emptyList;

            let valueType = metadata.type,
                fieldKind = metadata.kind,
                isPropertyIdentity = metadata.idOnEntityKey,
                Agg = QueryAggregateFunction; // alias

            if (!valueType)
                return emptyList;

            // Cannot aggregate on model measures
            if (fieldKind === FieldKind.Measure)
                return emptyList;

            // Cannot aggregate grouping exprs
            if (isGroupingOnly)
                return emptyList;

            if (valueType.numeric || valueType.integer) {
                let aggregates = [Agg.Sum, Agg.Avg, Agg.Min, Agg.Max, Agg.Count, Agg.CountNonNull, Agg.StandardDeviation, Agg.Variance];
                let fieldExpr = SQExprConverter.asFieldPattern(expr);
                let fieldExprItem = FieldExprPattern.toFieldExprEntityItemPattern(fieldExpr);

                let currentSchema = schema.schema(fieldExprItem.schema);
                if (currentSchema.capabilities.supportsMedian)
                    aggregates.push(Agg.Median);
                return aggregates;
            } else if (valueType.text || valueType.bool || valueType.dateTime) {
                if (isPropertyIdentity)
                    return [Agg.CountNonNull];
                return [Agg.Count, Agg.CountNonNull];
            }

            debug.assertFail("Unexpected expr or role.");
            return emptyList;
        }

        export function indexOfExpr(items: SQExpr[], searchElement: SQExpr): number {
            debug.assertValue(items, 'items');
            debug.assertValue(searchElement, 'searchElement');

            for (let i = 0, len = items.length; i < len; i++) {
                if (SQExpr.equals(items[i], searchElement))
                    return i;
            }
            return -1;
        }

        export function sequenceEqual(x: SQExpr[], y: SQExpr[]): boolean {
            debug.assertValue(x, 'x');
            debug.assertValue(y, 'y');

            let len = x.length;
            if (len !== y.length)
                return false;

            for (let i = 0; i < len; i++) {
                if (!SQExpr.equals(x[i], y[i]))
                    return false;
            }

            return true;
        }

        export function uniqueName(namedItems: NamedSQExpr[], expr: SQExpr): string {
            debug.assertValue(namedItems, 'namedItems');

            // Determine all names
            let names: { [name: string]: boolean } = {};
            for (let i = 0, len = namedItems.length; i < len; i++)
                names[namedItems[i].name] = true;

            return StringExtensions.findUniqueName(names, defaultName(expr));
        }

        /** Generates a default expression name  */
        export function defaultName(expr: SQExpr, fallback: string = 'select'): string {
            if (!expr)
                return fallback;

            return expr.accept(SQExprDefaultNameGenerator.instance, fallback);
        }

        /** Gets a value indicating whether the expr is a model measure or an aggregate. */
        export function isMeasure(expr: SQExpr): boolean {
            debug.assertValue(expr, 'expr');

            return expr.accept(IsMeasureVisitor.instance);
        }

        /** Gets a value indicating whether the expr is an AnyValue or equals comparison to AnyValue*/
        export function isAnyValue(expr: SQExpr): boolean {
            debug.assertValue(expr, 'expr');

            return expr.accept(IsAnyValueVisitor.instance);
        }

        /** Gets a value indicating whether the expr is a DefaultValue or equals comparison to DefaultValue*/
        export function isDefaultValue(expr: SQExpr): boolean {
            debug.assertValue(expr, 'expr');

            return expr.accept(IsDefaultValueVisitor.instance);
        }

        export function discourageAggregation(expr: SQExpr, schema: FederatedConceptualSchema): boolean {
            debug.assertValue(expr, 'expr');
            debug.assertValue(schema, 'schema');

            let field = SQExprConverter.asFieldPattern(expr);
            if (!field)
                return;

            let fieldExprItem = FieldExprPattern.toFieldExprEntityItemPattern(field);
            let conceptualSchema = schema.schema(fieldExprItem.schema);
            if (conceptualSchema)
                return conceptualSchema.capabilities && conceptualSchema.capabilities.discourageQueryAggregateUsage;
        }

        export function getKpiStatus(expr: SQExpr, schema: FederatedConceptualSchema): SQExpr {
            let kpi = getKpiStatusProperty(expr, schema);
            if (kpi) {
                let measureExpr = SQExprConverter.asFieldPattern(expr).measure;
                if (measureExpr) {
                    return SQExprBuilder.fieldExpr({
                        measure: {
                            schema: measureExpr.schema,
                            entity: measureExpr.entity,
                            name: kpi.name,
                        }
                    });
                }
            }
        }

        export function getKpiStatusGraphic(expr: SQExpr, schema: FederatedConceptualSchema): string {
            let property = expr.getConceptualProperty(schema);
            if (!property)
                return;

            if (property.measure &&
                property.measure.kpi)
                return property.measure.kpi.statusGraphic;

            let kpi = getKpiStatusProperty(expr, schema);
            if (kpi)
                return kpi.measure.kpi.statusGraphic;
        }

        export function getConceptualHierarchy(sqExpr: SQExpr, federatedSchema: FederatedConceptualSchema): ConceptualHierarchy {
            if (sqExpr instanceof SQHierarchyExpr) {
                let hierarchy = <SQHierarchyExpr>sqExpr;
                let entityExpr = <SQEntityExpr>sqExpr.arg;
                return federatedSchema
                    .schema(entityExpr.schema)
                    .findHierarchy(entityExpr.entity, hierarchy.hierarchy);
            }
        }

        export function getExpr(schema: FederatedConceptualSchema, expr: SQExpr): SQExpr | SQExpr[] {
            return SQExprHierarchyToHierarchyLevelConverter.convert(expr, schema) ||
                SQExprVariationConverter.expand(expr, schema) ||
                expr;
        }

        function getKpiStatusProperty(expr: SQExpr, schema: FederatedConceptualSchema): ConceptualProperty {
            let property = expr.getConceptualProperty(schema);
            if (!property)
                return;

            let kpi = property.kpi;
            if (kpi && kpi.measure.kpi.status === property)
                return kpi;
        }

        function getMetadataForUnderlyingType(expr: SQExpr, schema: FederatedConceptualSchema): SQExprMetadata {
            // Unwrap the aggregate (if the expr has one), and look at the underlying type.
            let metadata = SQExprBuilder.removeAggregate(expr).getMetadata(schema);

            if (!metadata)
                metadata = expr.getMetadata(schema);

            return metadata;
        }

        class SQExprDefaultNameGenerator extends DefaultSQExprVisitorWithArg<string, string> {
            public static instance: SQExprDefaultNameGenerator = new SQExprDefaultNameGenerator();

            public visitEntity(expr: SQEntityExpr): string {
                return expr.entity;
            }

            public visitColumnRef(expr: SQColumnRefExpr): string {
                return expr.source.accept(this) + '.' + expr.ref;
            }

            public visitMeasureRef(expr: SQMeasureRefExpr, fallback: string): string {
                return expr.source.accept(this) + '.' + expr.ref;
            }

            public visitAggr(expr: SQAggregationExpr, fallback: string): string {
                return QueryAggregateFunction[expr.func] + '(' + expr.arg.accept(this) + ')';
            }

            public visitDefault(expr: SQExpr, fallback: string): string {
                return fallback || 'expr';
            }
        }

        class IsMeasureVisitor extends DefaultSQExprVisitor<boolean> {
            public static instance: IsMeasureVisitor = new IsMeasureVisitor();

            public visitMeasureRef(expr: SQMeasureRefExpr): boolean {
                return true;
            }

            public visitAggr(expr: SQAggregationExpr): boolean {
                return true;
            }

            public visitDefault(expr: SQExpr): boolean {
                return false;
            }
        }

        class IsDefaultValueVisitor extends DefaultSQExprVisitor<boolean> {
            public static instance: IsDefaultValueVisitor = new IsDefaultValueVisitor();

            public visitCompare(expr: SQCompareExpr): boolean {
                if (expr.kind !== QueryComparisonKind.Equal)
                    return false;

                return expr.right.accept(this);
            }

            public visitDefaultValue(expr: SQDefaultValueExpr): boolean {
                return true;
            }

            public visitDefault(expr: SQExpr): boolean {
                return false;
            }
        }

        class IsAnyValueVisitor extends DefaultSQExprVisitor<boolean> {
            public static instance: IsAnyValueVisitor = new IsAnyValueVisitor();

            public visitCompare(expr: SQCompareExpr): boolean {
                if (expr.kind !== QueryComparisonKind.Equal)
                    return false;

                return expr.right.accept(this);
            }

            public visitAnyValue(expr: SQAnyValueExpr): boolean {
                return true;
            }

            public visitDefault(expr: SQExpr): boolean {
                return false;
            }
        }
    }
}