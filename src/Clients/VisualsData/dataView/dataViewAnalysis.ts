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

module powerbi {
    export module DataViewAnalysis {
        import ArrayExtensions = jsCommon.ArrayExtensions;
        import QueryProjectionByProperty = powerbi.data.QueryProjectionsByRole;

        export interface ValidateAndReshapeResult {
            dataView?: DataView;
            isValid: boolean;
        }

        /** Reshapes the data view to match the provided schema if possible. If not, returns null */
        export function validateAndReshape(dataView: DataView, dataViewMappings: DataViewMapping[]): ValidateAndReshapeResult {
            if (!dataViewMappings || dataViewMappings.length === 0)
                return { dataView: dataView, isValid: true };

            if (dataView) {
                let dataViewMapping = dataViewMappings[0];

                // Keep the original when possible.
                if (supports(dataView, dataViewMapping))
                    return { dataView: dataView, isValid: true };

                if (dataViewMapping.categorical)
                    return reshapeCategorical(dataView, dataViewMapping);

                if (dataViewMapping.tree)
                    return reshapeTree(dataView, dataViewMapping.tree);

                if (dataViewMapping.single)
                    return reshapeSingle(dataView, dataViewMapping.single);

                if (dataViewMapping.table)
                    return reshapeTable(dataView, dataViewMapping.table);
            }

            return { isValid: false };
        }

        function reshapeCategorical(dataView: DataView, dataViewMapping: DataViewMapping): ValidateAndReshapeResult {
            debug.assertValue(dataViewMapping, 'dataViewMapping');
            //The functionality that used to compare categorical.values.length to schema.values doesn't apply any more, we don't want to use the same logic for re-shaping.
            let categoryRoleMapping = dataViewMapping.categorical;
            let categorical = dataView.categorical;
            if (!categorical)
                return { isValid: false };

            let rowCount;
            if (categoryRoleMapping.rowCount) {
                rowCount = categoryRoleMapping.rowCount.supported;
                if (rowCount && rowCount.max) {
                    let updated: DataViewCategorical;
                    let categories = categorical.categories;
                    let maxRowCount = rowCount.max;
                    let originalLength = undefined;
                    if (categories) {
                        for (let i = 0, len = categories.length; i < len; i++) {
                            let category = categories[i];
                            originalLength = category.values.length;
                            if (maxRowCount !== undefined && originalLength > maxRowCount) {

                                // Row count too large: Trim it to fit.
                                let updatedCategories = ArrayExtensions.range(category.values, 0, maxRowCount - 1);

                                updated = updated || { categories: [] };
                                updated.categories.push({
                                    source: category.source,
                                    values: updatedCategories
                                });
                            }
                        }
                    }

                    if (categorical.values && categorical.values.length > 0 && maxRowCount) {
                        if (!originalLength)
                            originalLength = categorical.values[0].values.length;

                        if (maxRowCount !== undefined && originalLength > maxRowCount) {
                            updated = updated || {};
                            updated.values = data.DataViewTransform.createValueColumns();

                            for (let i = 0, len = categorical.values.length; i < len; i++) {
                                let column = categorical.values[i],
                                    updatedColumn: DataViewValueColumn = {
                                        source: column.source,
                                        values: ArrayExtensions.range(column.values, 0, maxRowCount - 1)
                                    };

                                if (column.min !== undefined)
                                    updatedColumn.min = column.min;
                                if (column.max !== undefined)
                                    updatedColumn.max = column.max;
                                if (column.subtotal !== undefined)
                                    updatedColumn.subtotal = column.subtotal;

                                updated.values.push(updatedColumn);
                            }
                        }
                    }

                    if (updated) {
                        dataView = {
                            metadata: dataView.metadata,
                            categorical: updated,
                        };
                    }
                }
            }

            if (supportsCategorical(dataView, dataViewMapping))
                return { dataView: dataView, isValid: true };

            return null;
        }

        function reshapeSingle(dataView: DataView, singleRoleMapping: DataViewSingleMapping): ValidateAndReshapeResult {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(singleRoleMapping, 'singleRoleMapping');

            if (dataView.single)
                return { dataView: dataView, isValid: true };

            return { isValid: false };
        }

        function reshapeTree(dataView: DataView, treeRoleMapping: DataViewTreeMapping): ValidateAndReshapeResult {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(treeRoleMapping, 'treeRoleMapping');

            // TODO: Need to implement the reshaping of Tree
            let metadata = dataView.metadata;
            if (conforms(countGroups(metadata.columns), treeRoleMapping.depth) /*&& conforms(countMeasures(metadata.columns), treeRoleMapping.aggregates)*/)
                return { dataView: dataView, isValid: true };

            return { isValid: false };
        }

        function reshapeTable(dataView: DataView, tableRoleMapping: DataViewTableMapping): ValidateAndReshapeResult {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(tableRoleMapping, 'tableRoleMapping');

            if (dataView.table)
                return { dataView: dataView, isValid: true };

            return { isValid: false };
        }

        export function countGroups(columns: DataViewMetadataColumn[]): number {
            let count = 0;

            for (let i = 0, len = columns.length; i < len; i++) {
                if (!columns[i].isMeasure)
                    ++count;
            }

            return count;
        }

        export function countMeasures(columns: DataViewMetadataColumn[]): number {
            let count = 0;

            for (let i = 0, len = columns.length; i < len; i++) {
                if (columns[i].isMeasure)
                    ++count;
            }

            return count;
        }

        /** Indicates whether the dataView conforms to the specified schema. */
        export function supports(dataView: DataView, roleMapping: DataViewMapping, usePreferredDataViewSchema?: boolean): boolean {
            if (!roleMapping || !dataView)
                return false;

            if (roleMapping.categorical && !supportsCategorical(dataView, roleMapping.categorical, usePreferredDataViewSchema))
                return false;

            if (roleMapping.tree && !supportsTree(dataView, roleMapping.tree))
                return false;

            if (roleMapping.single && !supportsSingle(dataView.single, roleMapping.single))
                return false;

            if (roleMapping.table && !supportsTable(dataView.table, roleMapping.table, usePreferredDataViewSchema))
                return false;

            return true;
        }

        function supportsCategorical(dataView: DataView, categoryRoleMapping: DataViewCategoricalMapping, usePreferredDataViewSchema?: boolean): boolean {
            debug.assertValue(categoryRoleMapping, 'categoryRoleMapping');

            let dataViewCategorical = dataView.categorical;
            if (!dataViewCategorical)
                return false;

            // TODO: Disabling this implementation isn't right.
            //if (!conforms(countMeasures(dataView.metadata.columns), categoryRoleMapping.values.roles.length))
            //    return false;

            if (categoryRoleMapping.rowCount) {
                let rowCount = categoryRoleMapping.rowCount.supported;
                if (usePreferredDataViewSchema && categoryRoleMapping.rowCount.preferred)
                    rowCount = categoryRoleMapping.rowCount.preferred;

                if (rowCount) {
                    let len: number = 0;
                    if (dataViewCategorical.values && dataViewCategorical.values.length)
                        len = dataViewCategorical.values[0].values.length;
                    else if (dataViewCategorical.categories && dataViewCategorical.categories.length)
                        len = dataViewCategorical.categories[0].values.length;

                    if (!conforms(len, rowCount))
                        return false;
                }
            }

            return true;
        }

        function supportsSingle(dataViewSingle: DataViewSingle, singleRoleMapping: DataViewSingleMapping): boolean {
            debug.assertValue(singleRoleMapping, 'singleRoleMapping');

            if (!dataViewSingle)
                return false;

            return true;
        }

        function supportsTree(dataView: DataView, treeRoleMapping: DataViewTreeMapping): boolean {
            debug.assertValue(treeRoleMapping, 'treeRoleMapping');

            let metadata = dataView.metadata;
            return conforms(countGroups(metadata.columns), treeRoleMapping.depth);
        }

        function supportsTable(dataViewTable: DataViewTable, tableRoleMapping: DataViewTableMapping, usePreferredDataViewSchema?: boolean): boolean {
            debug.assertValue(tableRoleMapping, 'tableRoleMapping');

            if (!dataViewTable)
                return false;

            if (tableRoleMapping.rowCount) {
                let rowCount = tableRoleMapping.rowCount.supported;
                if (usePreferredDataViewSchema && tableRoleMapping.rowCount.preferred)
                    rowCount = tableRoleMapping.rowCount.preferred;

                if (rowCount) {
                    let len: number = 0;
                    if (dataViewTable.rows && dataViewTable.rows.length)
                        len = dataViewTable.rows.length;

                    if (!conforms(len, rowCount))
                        return false;
                }
            }

            return true;
        }

        export function conforms(value: number, range: NumberRange, ignoreMin?: boolean): boolean {
            debug.assertValue(value, 'value');

            if (!range)
                return value === 0;

            if (!ignoreMin && range.min !== undefined && range.min > value)
                return false;

            if (range.max !== undefined && range.max < value)
                return false;

            return true;
        }

        /** Determines the appropriate DataViewMappings for the projections. */
        export function chooseDataViewMappings(projections: QueryProjectionByProperty, mappings: DataViewMapping[]): DataViewMapping[] {
            debug.assertValue(projections, 'projections');
            debug.assertValue(mappings, 'mappings');

            let supportedMappings: DataViewMapping[] = [];

            for (let i = 0, len = mappings.length; i < len; i++) {
                let mapping = mappings[i],
                    mappingConditions = mapping.conditions;

                if (mappingConditions && mappingConditions.length) {
                    for (let j = 0, jlen = mappingConditions.length; j < jlen; j++) {
                        let condition = mappingConditions[j];
                        if (matchesCondition(projections, condition)) {
                            supportedMappings.push(mapping);
                            break;
                        }
                    }
                }
                else {
                    supportedMappings.push(mapping);
                }
            }

            return ArrayExtensions.emptyToNull(supportedMappings);
        }

        function matchesCondition(projections: QueryProjectionByProperty, condition: DataViewMappingCondition): boolean {
            debug.assertValue(projections, 'projections');
            debug.assertValue(condition, 'condition');

            let conditionRoles = Object.keys(condition);
            for (let i = 0, len = conditionRoles.length; i < len; i++) {
                let roleName: string = conditionRoles[i],
                    isDrillable = projections[roleName] && projections[roleName].activeProjectionQueryRef != null,
                    range = condition[roleName];

                let roleCount = getPropertyCount(roleName, projections, isDrillable);
                if (!conforms(roleCount, range))
                    return false;
            }

            return true;
        }

        export function getPropertyCount(roleName: string, projections: QueryProjectionByProperty, useActiveIfAvailable?: boolean): number {
            debug.assertValue(roleName, 'roleName');
            debug.assertValue(projections, 'projections');

            let projectionsForRole = projections[roleName];
            if (projectionsForRole) {
                if (useActiveIfAvailable)
                    return 1;
                return projectionsForRole.all().length;
            }

            return 0;
        }

        export function hasSameCategoryIdentity(dataView1: DataView, dataView2: DataView): boolean {
            if (dataView1
                && dataView2
                && dataView1.categorical
                && dataView2.categorical) {
                let dv1Categories = dataView1.categorical.categories;
                let dv2Categories = dataView2.categorical.categories;
                if (dv1Categories
                    && dv2Categories
                    && dv1Categories.length === dv2Categories.length) {
                    for (let i = 0, len = dv1Categories.length; i < len; i++) {
                        let dv1Identity = dv1Categories[i].identity;
                        let dv2Identity = dv2Categories[i].identity;

                        let dv1Length = getLengthOptional(dv1Identity);
                        if (dv1Length !== getLengthOptional(dv2Identity))
                            return false;

                        for (let j = 0; j < dv1Length; j++) {
                            if (!DataViewScopeIdentity.equals(dv1Identity[j], dv2Identity[j]))
                                return false;
                        }
                    }

                    return true;
                }
            }

            return false;
        }

        function getLengthOptional(identity: DataViewScopeIdentity[]): number {
            if (identity)
                return identity.length;

            return 0;
        }

        export function areMetadataColumnsEquivalent(column1: DataViewMetadataColumn, column2: DataViewMetadataColumn): boolean {
            if (!column1 && !column2)
                return true;

            if (!column1 || !column2)
                return false;

            if (column1.displayName !== column2.displayName)
                return false;

            if (column1.queryName !== column2.queryName)
                return false;

            if (column1.isMeasure !== column2.isMeasure)
                return false;

            if (column1.type !== column2.type)
                return false;

            return true;
        }

        /* Returns true if the metadata columns at the same positions in the array are equivalent. */
        export function isMetadataEquivalent(metadata1: DataViewMetadata, metadata2: DataViewMetadata) {
            if (!metadata1 && !metadata2)
                return true;

            if (!metadata1 || !metadata2)
                return false;

            let previousColumnsLength = metadata1.columns.length;
            let newColumnsLength = metadata2.columns.length;

            if (previousColumnsLength !== newColumnsLength)
                return false;

            for (let i: number = 0; i < newColumnsLength; i++) {
                if (!DataViewAnalysis.areMetadataColumnsEquivalent(metadata1.columns[i], metadata2.columns[i]))
                    return false;
            }

            return true;
        }
    }
}