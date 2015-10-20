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
    import inherit = Prototype.inherit;
    import ArrayExtensions = jsCommon.ArrayExtensions;
    import INumberDictionary = jsCommon.INumberDictionary;

    export interface DataViewTransformApplyOptions {
        prototype: DataView;
        objectDescriptors: DataViewObjectDescriptors;
        dataViewMappings?: DataViewMapping[];
        transforms: DataViewTransformActions;
        colorAllocatorFactory: IColorAllocatorFactory;
    }

    /** Describes the Transform actions to be done to a prototype DataView. */
    export interface DataViewTransformActions {
        /** Describes transform metadata for each semantic query select item, as the arrays align, by index. */
        selects?: DataViewSelectTransform[];

        /** Describes the DataViewObject definitions. */
        objects?: DataViewObjectDefinitions;

        /** Describes the splitting of a single input DataView into multiple DataViews. */
        splits?: DataViewSplitTransform[];

        /** Describes the order of selects (referenced by query index) in each role. */
        projectionOrdering?: DataViewProjectionOrdering;
    }

    export interface DataViewSelectTransform {
        displayName?: string;
        queryName?: string;
        format?: string;
        type?: ValueType;
        roles?: { [roleName: string]: boolean };
        kpiStatusGraphic?: string;
        sort?: SortDirection;
    }

    export interface DataViewSplitTransform {
        selects: INumberDictionary<boolean>;
    }

    export interface DataViewProjectionOrdering {
        [roleName: string]: number[];
    }

    export interface MatrixTransformationContext {
        rowHierarchyRewritten: boolean;
        columnHierarchyRewritten: boolean;
        hierarchyTreesRewritten: boolean;
    }

    interface ValueRewrite<T> {
        from: T;
        to: T;
    }

    interface NumberToNumberMapping {
        [position: number]: number;
    }

    enum CategoricalDataViewTransformation {
        None,
        Pivot,
        SelfCrossJoin,
    }

    // TODO: refactor & focus DataViewTransform into a service with well-defined dependencies.
    export module DataViewTransform {
        export function apply(options: DataViewTransformApplyOptions): DataView[] {
            debug.assertValue(options, 'options');

            // TODO: Flow a context object through to capture errors/warnings about what happens here for better diagnosability.

            let prototype = options.prototype,
                objectDescriptors = options.objectDescriptors,
                dataViewMappings = options.dataViewMappings,
                transforms = options.transforms,
                colorAllocatorFactory = options.colorAllocatorFactory;

            if (!prototype)
                return transformEmptyDataView(objectDescriptors, transforms, colorAllocatorFactory);

            if (!transforms)
                return [prototype];

            let splits = transforms.splits;
            if (ArrayExtensions.isUndefinedOrEmpty(splits)) {
                return [transformDataView(prototype, objectDescriptors, dataViewMappings, transforms, colorAllocatorFactory)];
            }

            let transformedDataviews: DataView[] = [];
            for (let split of splits) {
                let transformed = transformDataView(prototype, objectDescriptors, dataViewMappings, transforms, colorAllocatorFactory, split.selects);
                transformedDataviews.push(transformed);
            }

            return transformedDataviews;
        }

        function transformEmptyDataView(objectDescriptors: DataViewObjectDescriptors, transforms: DataViewTransformActions, colorAllocatorFactory: IColorAllocatorFactory): DataView[] {
            if (transforms && transforms.objects) {
                let emptyDataView: DataView = {
                    metadata: {
                        columns: [],
                    }
                };

                transformObjects(emptyDataView, objectDescriptors, transforms.objects, transforms.selects, colorAllocatorFactory);

                return [emptyDataView];
            }

            return [];
        }

        function transformDataView(
            prototype: DataView,
            objectDescriptors: DataViewObjectDescriptors,
            roleMappings: DataViewMapping[],
            transforms: DataViewTransformActions,
            colorAllocatorFactory: IColorAllocatorFactory,
            selectsToInclude?: INumberDictionary<boolean>): DataView {
            debug.assertValue(prototype, 'prototype');

            let transformed = inherit(prototype);
            transformed.metadata = inherit(prototype.metadata);

            transformed = transformSelects(transformed, roleMappings, transforms.selects, transforms.projectionOrdering, selectsToInclude);
            transformObjects(transformed, objectDescriptors, transforms.objects, transforms.selects, colorAllocatorFactory);

            return transformed;
        }

        function transformSelects(
            dataView: DataView,
            roleMappings: DataViewMapping[],
            selectTransforms: DataViewSelectTransform[],
            projectionOrdering?: DataViewProjectionOrdering,
            selectsToInclude?: INumberDictionary<boolean>): DataView {

            let columnRewrites: ValueRewrite<DataViewMetadataColumn>[] = [];
            if (selectTransforms) {
                dataView.metadata.columns = applyTransformsToColumns(
                    dataView.metadata.columns,
                    selectTransforms,
                    columnRewrites);
            }
            
            // NOTE: no rewrites necessary for Tree (it doesn't reference the columns)
            if (dataView.categorical) {
                dataView.categorical = applyRewritesToCategorical(dataView.categorical, columnRewrites, selectsToInclude);

                // NOTE: This is slightly DSR-specific.
                dataView = pivotIfNecessary(dataView, roleMappings);
            }

            if (dataView.matrix) {
                let matrixTransformationContext: MatrixTransformationContext = {
                    rowHierarchyRewritten: false,
                    columnHierarchyRewritten: false,
                    hierarchyTreesRewritten: false
                };
                dataView.matrix = applyRewritesToMatrix(dataView.matrix, columnRewrites, roleMappings, projectionOrdering, matrixTransformationContext);

                if (shouldPivotMatrix(dataView.matrix, roleMappings))
                    DataViewPivotMatrix.apply(dataView.matrix, matrixTransformationContext);
            }

            if (dataView.table)
                dataView.table = applyRewritesToTable(dataView.table, columnRewrites, roleMappings, projectionOrdering);

            return dataView;
        }

        function applyTransformsToColumns(
            prototypeColumns: DataViewMetadataColumn[],
            selects: DataViewSelectTransform[],
            rewrites: ValueRewrite<DataViewMetadataColumn>[]): DataViewMetadataColumn[] {
            debug.assertValue(prototypeColumns, 'columns');

            if (!selects)
                return prototypeColumns;

            //column may contain undefined entries
            let columns = inherit(prototypeColumns);

            for (let i = 0, len = prototypeColumns.length; i < len; i++) {
                let prototypeColumn = prototypeColumns[i];
                let select = selects[prototypeColumn.index];
                if (!select)
                    continue;

                let column: DataViewMetadataColumn = columns[i] = inherit(prototypeColumn);

                if (select.roles)
                    column.roles = select.roles;
                if (select.type)
                    column.type = select.type;
                column.format = getFormatForColumn(select, column);

                if (select.displayName)
                    column.displayName = select.displayName;
                if (select.queryName)
                    column.queryName = select.queryName;
                if (select.kpiStatusGraphic)
                    column.kpiStatusGraphic = select.kpiStatusGraphic;
                if (select.sort)
                    column.sort = select.sort;

                rewrites.push({
                    from: prototypeColumn,
                    to: column,
                });
            }

            return columns;
        }

        /**
         * Get the column format. Order of precendence is:
         *  1. Select format
         *  2. Column format
         *  3. Default PowerView policy for column type
         */
        function getFormatForColumn(select: DataViewSelectTransform, column: DataViewMetadataColumn): string {
            // TODO: we already copied the select.Format to column.format, we probably don't need this check
            if (select.format)
                return select.format;

            if (column.format)
                return column.format;

            // TODO: deprecate this, default format string logic has been added to valueFormatter
            let type = column.type;
            if (type) {
                if (type.dateTime)
                    return 'd';
                if (type.integer)
                    return 'g';
                if (type.numeric)
                    return '#,0.00';
            }

            return undefined;
        }

        /**
         * 
         * 
         * Note: Exported for testability
         */
        export function upgradeSettingsToObjects(settings: VisualElementSettings, objectDefns?: DataViewObjectDefinitions): DataViewObjectDefinitions {
            if (!settings)
                return;

            if (!objectDefns)
                objectDefns = {};

            for (let propertyKey in settings) {
                let propertyValue = settings[propertyKey],
                    upgradedPropertyKey: string = propertyKey,
                    upgradedPropertyValue = propertyValue,
                    objectName: string = 'general';

                switch (propertyKey) {
                    case 'hasScalarCategoryAxis':
                        // hasScalarCategoryAxis -> categoryAxis.axisType
                        objectName = 'categoryAxis';
                        upgradedPropertyKey = 'axisType';
                        upgradedPropertyValue = SQExprBuilder.text(propertyValue ? axisType.scalar : axisType.categorical);
                        break;

                    case 'Totals':
                        // Totals -> general.totals
                        upgradedPropertyKey = 'totals';
                        upgradedPropertyValue = SQExprBuilder.boolean(!!propertyValue);
                        break;

                    case 'textboxSettings':
                        // textboxSettings.paragraphs -> general.paragraphs
                        upgradedPropertyKey = 'paragraphs';
                        if (propertyValue && propertyValue.paragraphs)
                            upgradedPropertyValue = propertyValue.paragraphs;
                        break;

                    case 'VisualType1':
                        // VisualType1 -> general.visualType1
                        upgradedPropertyKey = 'visualType1';
                        upgradedPropertyValue = SQExprBuilder.text(propertyValue);
                        break;

                    case 'VisualType2':
                        // VisualType2 -> general.visualType2
                        upgradedPropertyKey = 'visualType2';
                        upgradedPropertyValue = SQExprBuilder.text(propertyValue);
                        break;

                    case 'imageVisualSettings':
                        // imageVisualSettings.imageUrl -> general.imageUrl
                        upgradedPropertyKey = 'imageUrl';
                        if (propertyValue && propertyValue.imageUrl)
                            upgradedPropertyValue = SQExprBuilder.text(propertyValue.imageUrl);
                        break;

                    default:
                        // Ignore all other properties.
                        continue;
                }

                setObjectDefinition(objectDefns, objectName, upgradedPropertyKey, upgradedPropertyValue);
            }

            return objectDefns;
        }

        function setObjectDefinition(objects: DataViewObjectDefinitions, objectName: string, propertyName: string, value: any): void {
            debug.assertValue(objects, 'objects');
            debug.assertValue(objectName, 'objectName');
            debug.assertValue(propertyName, 'propertyName');

            let objectContainer: DataViewObjectDefinition[] = objects[objectName];
            if (objectContainer === undefined)
                objectContainer = objects[objectName] = [];

            let object: DataViewObjectDefinition = objectContainer[0];
            if (object === undefined)
                object = objectContainer[0] = { properties: {} };

            object.properties[propertyName] = value;
        }

        function applyRewritesToCategorical(prototype: DataViewCategorical, columnRewrites: ValueRewrite<DataViewMetadataColumn>[], selectsToInclude?: INumberDictionary<boolean>): DataViewCategorical {
            debug.assertValue(prototype, 'prototype');
            debug.assertValue(columnRewrites, 'columnRewrites');

            let categorical = inherit(prototype);

            function override(value: { source?: DataViewMetadataColumn }) {
                let rewrittenSource = findOverride(value.source, columnRewrites);
                if (rewrittenSource) {
                    let rewritten = inherit(value);
                    rewritten.source = rewrittenSource;
                    return rewritten;
                }
            }

            let categories = Prototype.overrideArray(prototype.categories, override);
            if (categories)
                categorical.categories = categories;

            let values = Prototype.overrideArray(prototype.values, override);

            if (values) {
                if (selectsToInclude) {
                    for (let i = values.length - 1; i >= 0; i--) {
                        if (!selectsToInclude[values[i].source.index])
                            values.splice(i, 1);
                    }
                }

                if (values.source) {
                    if (selectsToInclude && !selectsToInclude[values.source.index]) {
                        values.source = undefined;
                    }
                    else {
                        let rewrittenValuesSource = findOverride(values.source, columnRewrites);
                        if (rewrittenValuesSource)
                            values.source = rewrittenValuesSource;
                    }
                }

                categorical.values = values;
                setGrouped(values);
            }

            return categorical;
        }

        function applyRewritesToTable(
            prototype: DataViewTable,
            columnRewrites: ValueRewrite<DataViewMetadataColumn>[],
            roleMappings: DataViewMapping[],
            projectionOrdering: DataViewProjectionOrdering): DataViewTable {
            debug.assertValue(prototype, 'prototype');
            debug.assertValue(columnRewrites, 'columnRewrites');

            // Don't perform this potentially expensive transform unless we actually have a table. 
            // When we switch to lazy per-visual DataView creation, we'll be able to remove this check.
            if (!roleMappings || roleMappings.length !== 1 || !roleMappings[0].table)
                return prototype;

            let table = inherit(prototype);

            // Copy the rewritten columns into the table view
            let override = (metadata: DataViewMetadataColumn) => findOverride(metadata, columnRewrites);
            let columns = Prototype.overrideArray(prototype.columns, override);
            if (columns)
                table.columns = columns;

            if (!projectionOrdering)
                return table;

            let newToOldPositions = createTableColumnPositionMapping(projectionOrdering, columnRewrites);
            if (!newToOldPositions)
                return table;

            // Reorder the columns
            let columnsClone = columns.slice(0);
            let keys = Object.keys(newToOldPositions);
            for (let i = 0, len = keys.length; i < len; i++) {
                let sourceColumn = columnsClone[newToOldPositions[keys[i]]];

                // In the case we've hit the end of our columns array, but still have position reordering keys,
                // there is a duplicate column so we will need to add a new column for the duplicate data
                if (i === columns.length)
                    columns.push(sourceColumn);
                else {
                    debug.assert(i < columns.length, 'The column index is out of range for reordering.');
                    columns[i] = sourceColumn;
                }
            }
            
            // Reorder the rows
            let rows = Prototype.overrideArray(table.rows,
                (row: any[]) => {
                    let newRow: any[] = [];
                    for (let i = 0, len = keys.length; i < len; ++i)
                        newRow[i] = row[newToOldPositions[keys[i]]];

                    return newRow;
                });

            if (rows)
                table.rows = rows;

            return table;
        }

        /** Creates a mapping of new position to original position. */
        function createTableColumnPositionMapping(
            projectionOrdering: DataViewProjectionOrdering,
            columnRewrites: ValueRewrite<DataViewMetadataColumn>[]): NumberToNumberMapping {
            let roles = Object.keys(projectionOrdering);

            debug.assert(roles.length === 1, "Tables should have exactly one role.");

            let role = roles[0],
                originalOrder = _.map(columnRewrites, (rewrite: ValueRewrite<DataViewMetadataColumn>) => rewrite.from.index),
                newOrder = projectionOrdering[role];

            // Optimization: avoid rewriting the table if all columns are in their default positions.
            if (ArrayExtensions.sequenceEqual(originalOrder, newOrder, (x: number, y: number) => x === y))
                return;

            return createOrderMapping(originalOrder, newOrder);
        }

        function applyRewritesToMatrix(
            prototype: DataViewMatrix,
            columnRewrites: ValueRewrite<DataViewMetadataColumn>[],
            roleMappings: DataViewMapping[],
            projectionOrdering: DataViewProjectionOrdering,
            context: MatrixTransformationContext): DataViewMatrix {
            debug.assertValue(prototype, 'prototype');
            debug.assertValue(columnRewrites, 'columnRewrites');

            // Don't perform this potentially expensive transform unless we actually have a matrix. 
            // When we switch to lazy per-visual DataView creation, we'll be able to remove this check.
            if (!roleMappings || roleMappings.length !== 1 || !roleMappings[0].matrix)
                return prototype;

            let matrix = inherit(prototype);

            function override(metadata: DataViewMetadataColumn) {
                return findOverride(metadata, columnRewrites);
            }

            function overrideHierarchy(hierarchy: DataViewHierarchy): DataViewHierarchy {
                let rewrittenHierarchy: DataViewHierarchy = null;

                let newLevels = Prototype.overrideArray(
                    hierarchy.levels,
                    (level: DataViewHierarchyLevel) => {
                        let newLevel: DataViewHierarchyLevel = null;
                        let levelSources = Prototype.overrideArray(level.sources, override);
                        if (levelSources)
                            newLevel = ensureRewritten<DataViewHierarchyLevel>(newLevel, level, h => h.sources = levelSources);

                        return newLevel;
                    });
                if (newLevels)
                    rewrittenHierarchy = ensureRewritten<DataViewHierarchy>(rewrittenHierarchy, hierarchy, r => r.levels = newLevels);

                return rewrittenHierarchy;
            }

            let rows = overrideHierarchy(matrix.rows);
            if (rows) {
                matrix.rows = rows;
                context.rowHierarchyRewritten = true;
            }

            let columns = overrideHierarchy(matrix.columns);
            if (columns) {
                matrix.columns = columns;
                context.columnHierarchyRewritten = true;
            }

            let valueSources = Prototype.overrideArray(matrix.valueSources, override);
            if (valueSources) {
                matrix.valueSources = valueSources;

                // Only need to reorder if we have more than one value source, and they are all bound to the same role
                let matrixValues = <DataViewRoleForMapping>roleMappings[0].matrix.values;
                if (projectionOrdering && valueSources.length > 1 && matrixValues && matrixValues.for) {
                    let columnLevels = columns.levels.length;
                    if (columnLevels > 0) {
                        let newToOldPositions = createMatrixValuesPositionMapping(matrixValues, projectionOrdering, valueSources, columnRewrites);
                        if (newToOldPositions) {
                            let keys = Object.keys(newToOldPositions);
                            let numKeys = keys.length;

                            // Reorder the value columns
                            columns.root = DataViewPivotMatrix.cloneTree(columns.root);
                            if (columnLevels === 1)
                                reorderChildNodes(columns.root, newToOldPositions);
                            else
                                forEachNodeAtLevel(columns.root, columnLevels - 2, (node: DataViewMatrixNode) => reorderChildNodes(node, newToOldPositions));

                            // Reorder the value rows
                            matrix.rows.root = DataViewPivotMatrix.cloneTreeExecuteOnLeaf(matrix.rows.root, (node: DataViewMatrixNode) => {

                                if (!node.values)
                                    return;

                                let newValues: { [id: number]: DataViewTreeNodeValue } = {};

                                let iterations = Object.keys(node.values).length / numKeys;
                                for (let i = 0, len = iterations; i < len; i++) {
                                    let offset = i * numKeys;
                                    for (let keysIndex = 0; keysIndex < numKeys; keysIndex++)
                                        newValues[offset + keysIndex] = node.values[offset + newToOldPositions[keys[keysIndex]]];
                                }

                                node.values = newValues;
                            });

                            context.hierarchyTreesRewritten = true;
                        }
                    }
                }
            }

            return matrix;
        }

        function reorderChildNodes(node: DataViewMatrixNode, newToOldPositions: NumberToNumberMapping): void {
            let keys = Object.keys(newToOldPositions);
            let numKeys = keys.length;
            let children = node.children;

            let childrenClone = children.slice(0);
            for (let i = 0, len = numKeys; i < len; i++) {
                let sourceColumn = childrenClone[newToOldPositions[keys[i]]];

                // In the case we've hit the end of our columns array, but still have position reordering keys,
                // there is a duplicate column so we will need to add a new column for the duplicate data
                if (i === children.length)
                    children.push(sourceColumn);
                else {
                    debug.assert(i < children.length, 'The column index is out of range for reordering.');
                    children[i] = sourceColumn;
                }
            }
        }

        /** Creates a mapping of new position to original position. */
        function createMatrixValuesPositionMapping(
            matrixValues: DataViewRoleForMapping,
            projectionOrdering: DataViewProjectionOrdering,
            valueSources: DataViewMetadataColumn[],
            columnRewrites: ValueRewrite<DataViewMetadataColumn>[]): NumberToNumberMapping {

            let role = matrixValues.for.in;

            function matchValueSource(columnRewrite: ValueRewrite<DataViewMetadataColumn>) {
                for (let i = 0, len = valueSources.length; i < len; i++) {
                    let valueSource = valueSources[i];
                    if (valueSource === columnRewrite.to)
                        return columnRewrite;
                }
            }

            let valueRewrites: ValueRewrite<DataViewMetadataColumn>[] = [];
            for (let i = 0, len = columnRewrites.length; i < len; i++) {
                let columnRewrite = columnRewrites[i];
                if (matchValueSource(columnRewrite))
                    valueRewrites.push(columnRewrite);
            }

            let newOrder = projectionOrdering[role];
            let originalOrder = _.map(valueRewrites, (rewrite: ValueRewrite<DataViewMetadataColumn>) => rewrite.from.index);

            // Optimization: avoid rewriting the matrix if all leaf nodes are in their default positions.
            if (ArrayExtensions.sequenceEqual(originalOrder, newOrder, (x: number, y: number) => x === y))
                return;

            return createOrderMapping(originalOrder, newOrder);
        }

        function createOrderMapping(originalOrder: number[], newOrder: number[]): NumberToNumberMapping {
            let mapping: NumberToNumberMapping = {};
            for (let i = 0, len = newOrder.length; i < len; ++i) {
                let newPosition = newOrder[i];
                mapping[i] = originalOrder.indexOf(newPosition);
            }

            return mapping;
        }

        function forEachNodeAtLevel(node: DataViewMatrixNode, targetLevel: number, callback: (node: DataViewMatrixNode) => void): void {
            if (node.level === targetLevel) {
                callback(node);
                return;
            }

            let children = node.children;
            if (children && children.length > 0) {
                for (let i = 0, ilen = children.length; i < ilen; i++)
                    forEachNodeAtLevel(children[i], targetLevel, callback);
            }
        }

        function findOverride(source: DataViewMetadataColumn, columnRewrites: ValueRewrite<DataViewMetadataColumn>[]): DataViewMetadataColumn {
            for (let i = 0, len = columnRewrites.length; i < len; i++) {
                let columnRewrite = columnRewrites[i];
                if (columnRewrite.from === source)
                    return columnRewrite.to;
            }
        }

        function ensureRewritten<T>(rewritten: T, prototype: T, callback?: (rewritten: T) => void): T {
            if (!rewritten)
                rewritten = inherit(prototype);

            if (callback)
                callback(rewritten);

            return rewritten;
        }

        function transformObjects(
            dataView: DataView,
            objectDescriptors: DataViewObjectDescriptors,
            objectDefinitions: DataViewObjectDefinitions,
            selectTransforms: DataViewSelectTransform[],
            colorAllocatorFactory: IColorAllocatorFactory): void {
            debug.assertValue(dataView, 'dataView');
            debug.assertAnyValue(objectDescriptors, 'objectDescriptors');
            debug.assertAnyValue(objectDefinitions, 'objectDefinitions');
            debug.assertAnyValue(selectTransforms, 'selectTransforms');
            debug.assertValue(colorAllocatorFactory, 'colorAllocatorFactory');

            if (!objectDescriptors)
                return;

            let objectsForAllSelectors = DataViewObjectEvaluationUtils.groupObjectsBySelector(objectDefinitions);
            if (selectTransforms)
                DataViewObjectEvaluationUtils.addDefaultFormatString(objectsForAllSelectors, objectDescriptors, dataView.metadata.columns, selectTransforms);

            let metadataOnce = objectsForAllSelectors.metadataOnce;
            let dataObjects = objectsForAllSelectors.data;
            if (metadataOnce)
                evaluateMetadataObjects(dataView, objectDescriptors, metadataOnce.objects, dataObjects, colorAllocatorFactory);

            let metadataObjects = objectsForAllSelectors.metadata;
            if (metadataObjects) {
                for (let i = 0, len = metadataObjects.length; i < len; i++) {
                    let metadataObject = metadataObjects[i];
                    evaluateMetadataRepetition(dataView, objectDescriptors, metadataObject.selector, metadataObject.objects);
                }
            }

            for (let i = 0, len = dataObjects.length; i < len; i++) {
                let dataObject = dataObjects[i];
                evaluateDataRepetition(dataView, objectDescriptors, dataObject.selector, dataObject.rules, dataObject.objects);
            }

            if (objectsForAllSelectors.userDefined) {
                // TODO: Implement this.
            }
        }

        /** Evaluates and sets properties on the DataView metadata. */
        function evaluateMetadataObjects(
            dataView: DataView,
            objectDescriptors: DataViewObjectDescriptors,
            objectDefns: DataViewNamedObjectDefinition[],
            dataObjects: DataViewObjectDefinitionsForSelectorWithRule[],
            colorAllocatorFactory: IColorAllocatorFactory): void {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(objectDescriptors, 'objectDescriptors');
            debug.assertValue(objectDefns, 'objectDefns');
            debug.assertValue(dataObjects, 'dataObjects');
            debug.assertValue(colorAllocatorFactory, 'colorAllocatorFactory');

            let objects = DataViewObjectEvaluationUtils.evaluateDataViewObjects(objectDescriptors, objectDefns);
            if (objects) {
                dataView.metadata.objects = objects;

                for (let objectName in objects) {
                    let object: DataViewObject = objects[objectName],
                        objectDesc = objectDescriptors[objectName];

                    for (let propertyName in object) {
                        let propertyDesc = objectDesc.properties[propertyName],
                            ruleDesc = propertyDesc.rule;
                        if (!ruleDesc)
                            continue;

                        let definition = createRuleEvaluationInstance(
                            dataView,
                            colorAllocatorFactory,
                            ruleDesc,
                            objectName,
                            object[propertyName],
                            propertyDesc.type);
                        if (!definition)
                            continue;

                        dataObjects.push(definition);
                    }
                }
            }
        }

        function createRuleEvaluationInstance(
            dataView: DataView,
            colorAllocatorFactory: IColorAllocatorFactory,
            ruleDesc: DataViewObjectPropertyRuleDescriptor,
            objectName: string,
            propertyValue: DataViewPropertyValue,
            ruleType: StructuralTypeDescriptor): DataViewObjectDefinitionsForSelectorWithRule {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(colorAllocatorFactory, 'colorAllocatorFactory');
            debug.assertValue(ruleDesc, 'ruleDesc');
            debug.assertValue(propertyValue, 'propertyValue');
            debug.assertValue(ruleType, 'ruleType');

            let ruleOutput = ruleDesc.output;
            if (!ruleOutput)
                return;

            let selectorToCreate = findSelectorForRuleInput(dataView, ruleOutput.selector);
            if (!selectorToCreate)
                return;

            if (ruleType.fillRule)
                return createRuleEvaluationInstanceFillRule(dataView, colorAllocatorFactory, ruleDesc, selectorToCreate, objectName, <FillRule>propertyValue);

            if (ruleType.filter)
                return createRuleEvaluationInstanceFilter(dataView, ruleDesc, selectorToCreate, objectName, <SemanticFilter>propertyValue);
        }

        function createRuleEvaluationInstanceFillRule(
            dataView: DataView,
            colorAllocatorFactory: IColorAllocatorFactory,
            ruleDesc: DataViewObjectPropertyRuleDescriptor,
            selectorToCreate: Selector,
            objectName: string,
            propertyValue: FillRule): DataViewObjectDefinitionsForSelectorWithRule {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(colorAllocatorFactory, 'colorAllocatorFactory');
            debug.assertValue(ruleDesc, 'ruleDesc');
            debug.assertValue(selectorToCreate, 'selectorToCreate');
            debug.assertValue(propertyValue, 'propertyValue');

            let colorAllocator: IColorAllocator;
            if (propertyValue.linearGradient2)
                colorAllocator = createColorAllocatorLinearGradient2(dataView, colorAllocatorFactory, ruleDesc, propertyValue, propertyValue.linearGradient2);
            else if (propertyValue.linearGradient3)
                colorAllocator = createColorAllocatorLinearGradient3(dataView, colorAllocatorFactory, ruleDesc, propertyValue, propertyValue.linearGradient3);

            if (!colorAllocator)
                return;

            let rule = new ColorRuleEvaluation(ruleDesc.inputRole, colorAllocator);
            let fillRuleProperties: DataViewObjectPropertyDefinitions = {};
            fillRuleProperties[ruleDesc.output.property] = {
                solid: { color: rule }
            };

            return {
                selector: selectorToCreate,
                rules: [rule],
                objects: [{
                    name: objectName,
                    properties: fillRuleProperties,
                }]
            };
        }

        function createColorAllocatorLinearGradient2(
            dataView: DataView,
            colorAllocatorFactory: IColorAllocatorFactory,
            ruleDesc: DataViewObjectPropertyRuleDescriptor,
            propertyValueFillRule: FillRule,
            linearGradient2: LinearGradient2): IColorAllocator {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(colorAllocatorFactory, 'colorAllocatorFactory');
            debug.assertValue(ruleDesc, 'ruleDesc');
            debug.assertValue(linearGradient2, 'linearGradient2');

            linearGradient2 = propertyValueFillRule.linearGradient2;
            if (linearGradient2.min.value === undefined ||
                linearGradient2.max.value === undefined) {
                let inputRange = findRuleInputColumnNumberRange(dataView, ruleDesc.inputRole);
                if (!inputRange)
                    return;

                if (linearGradient2.min.value === undefined)
                    linearGradient2.min.value = inputRange.min;
                if (linearGradient2.max.value === undefined)
                    linearGradient2.max.value = inputRange.max;
            }

            return colorAllocatorFactory.linearGradient2(propertyValueFillRule.linearGradient2);
        }

        function createColorAllocatorLinearGradient3(
            dataView: DataView,
            colorAllocatorFactory: IColorAllocatorFactory,
            ruleDesc: DataViewObjectPropertyRuleDescriptor,
            propertyValueFillRule: FillRule,
            linearGradient3: LinearGradient3): IColorAllocator {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(colorAllocatorFactory, 'colorAllocatorFactory');
            debug.assertValue(ruleDesc, 'ruleDesc');
            debug.assertValue(linearGradient3, 'linearGradient3');
            let splitScales: boolean = undefined;
            linearGradient3 = propertyValueFillRule.linearGradient3;
            if (linearGradient3.min.value === undefined ||
                linearGradient3.mid.value === undefined ||
                linearGradient3.max.value === undefined) {
                let inputRange = findRuleInputColumnNumberRange(dataView, ruleDesc.inputRole);
                if (!inputRange)
                    return;

                splitScales =
                    linearGradient3.min.value === undefined &&
                    linearGradient3.max.value === undefined &&
                    linearGradient3.mid.value !== undefined;

                if (linearGradient3.min.value === undefined) {
                    linearGradient3.min.value = inputRange.min;
                }
                if (linearGradient3.max.value === undefined) {
                    linearGradient3.max.value = inputRange.max;
                }
                if (linearGradient3.mid.value === undefined) {
                    let midValue: number = (linearGradient3.max.value + linearGradient3.min.value) / 2;
                    linearGradient3.mid.value = midValue;
                }
            }

            return colorAllocatorFactory.linearGradient3(propertyValueFillRule.linearGradient3, splitScales);
        }

        function createRuleEvaluationInstanceFilter(
            dataView: DataView,
            ruleDesc: DataViewObjectPropertyRuleDescriptor,
            selectorToCreate: Selector,
            objectName: string,
            propertyValue: SemanticFilter): DataViewObjectDefinitionsForSelectorWithRule {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(ruleDesc, 'ruleDesc');
            debug.assertValue(selectorToCreate, 'selectorToCreate');
            debug.assertValue(propertyValue, 'propertyValue');

            // NOTE: This is not right -- we ought to discover the keys from an IN operator expression to avoid additional dependencies.
            if (!dataView.categorical || !dataView.categorical.categories || dataView.categorical.categories.length !== 1)
                return;
            let identityFields = dataView.categorical.categories[0].identityFields;
            if (!identityFields)
                return;

            let scopeIds = SQExprConverter.asScopeIdsContainer(propertyValue, identityFields);
            if (!scopeIds)
                return;

            let rule = new FilterRuleEvaluation(scopeIds);
            let properties: DataViewObjectPropertyDefinitions = {};
            properties[ruleDesc.output.property] = rule;

            return {
                selector: selectorToCreate,
                rules: [rule],
                objects: [{
                    name: objectName,
                    properties: properties,
                }]
            };
        }

        function evaluateDataRepetition(
            dataView: DataView,
            objectDescriptors: DataViewObjectDescriptors,
            selector: Selector,
            rules: RuleEvaluation[],
            objectDefns: DataViewNamedObjectDefinition[]): void {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(objectDescriptors, 'objectDescriptors');
            debug.assertValue(selector, 'selector');
            debug.assertAnyValue(rules, 'rules');
            debug.assertValue(objectDefns, 'objectDefns');

            let containsWildcard = Selector.containsWildcard(selector);

            let dataViewCategorical = dataView.categorical;
            if (dataViewCategorical) {
                let selectorMatched: boolean = false;

                // 1) Match against categories
                selectorMatched = evaluateDataRepetitionCategoricalCategory(dataViewCategorical, objectDescriptors, selector, rules, containsWildcard, objectDefns) || selectorMatched;

                // 2) Match against valueGrouping
                selectorMatched = evaluateDataRepetitionCategoricalValueGrouping(dataViewCategorical, objectDescriptors, selector, rules, containsWildcard, objectDefns) || selectorMatched;

                if (selectorMatched)
                    return;
            }

            // If we made it here, nothing matched the selector.  Consider capturing this in a diagnostics/context object to help debugging.
        }

        function evaluateDataRepetitionCategoricalCategory(
            dataViewCategorical: DataViewCategorical,
            objectDescriptors: DataViewObjectDescriptors,
            selector: Selector,
            rules: RuleEvaluation[],
            containsWildcard: boolean,
            objectDefns: DataViewNamedObjectDefinition[]): boolean {
            debug.assertValue(dataViewCategorical, 'dataViewCategorical');
            debug.assertValue(objectDescriptors, 'objectDescriptors');
            debug.assertValue(selector, 'selector');
            debug.assertAnyValue(rules, 'rules');
            debug.assertValue(containsWildcard, 'containsWildcard');
            debug.assertValue(objectDefns, 'objectDefns');

            if (!dataViewCategorical.categories || dataViewCategorical.categories.length === 0)
                return;

            let targetColumn = findSelectedCategoricalColumn(dataViewCategorical, selector);
            if (!targetColumn)
                return;

            let identities = targetColumn.identities,
                foundMatch: boolean,
                matchedRules: { rule: RuleEvaluation; inputValues: any[] }[];

            if (!identities)
                return;

            debug.assert(targetColumn.column.values.length === identities.length, 'Column length mismatch');

            for (let i = 0, len = identities.length; i < len; i++) {
                let identity = identities[i];

                if (containsWildcard || Selector.matchesData(selector, [identity])) {
                    // Set the context for any rules.
                    if (rules) {
                        if (!matchedRules)
                            matchedRules = matchRulesToDataViewCategorical(rules, dataViewCategorical);

                        for (let ruleIdx = 0, ruleLen = matchedRules.length; ruleIdx < ruleLen; ruleIdx++) {
                            let matchedRule = matchedRules[ruleIdx];
                            matchedRule.rule.setContext(identity, matchedRule.inputValues ? matchedRule.inputValues[i] : undefined);
                        }
                    }

                    let objects = DataViewObjectEvaluationUtils.evaluateDataViewObjects(objectDescriptors, objectDefns);
                    if (objects) {
                        // TODO: This mutates the DataView -- the assumption is that prototypal inheritance has already occurred.  We should
                        // revisit this, likely when we do lazy evaluation of DataView.
                        if (!targetColumn.column.objects) {
                            targetColumn.column.objects = [];
                            targetColumn.column.objects.length = len;
                        }
                        targetColumn.column.objects[i] = objects;
                    }

                    if (!containsWildcard)
                        return true;

                    foundMatch = true;
                }
            }

            return foundMatch;
        }

        function matchRulesToDataViewCategorical(
            rules: RuleEvaluation[],
            dataViewCategorical: DataViewCategorical): { rule: RuleEvaluation; inputValues: any[] }[] {
            let result: { rule: RuleEvaluation; inputValues: any[] }[] = [];

            for (let i = 0, len = rules.length; i < len; i++) {
                let rule = rules[i],
                    inputColumn = findRuleInputCategoricalColumn(dataViewCategorical, rule.inputRole);

                let inputValues: any[];
                if (inputColumn)
                    inputValues = inputColumn.values;

                result.push({
                    rule: rule,
                    inputValues: inputValues,
                });
            }

            return result;
        }

        function evaluateDataRepetitionCategoricalValueGrouping(
            dataViewCategorical: DataViewCategorical,
            objectDescriptors: DataViewObjectDescriptors,
            selector: Selector,
            rules: RuleEvaluation[],
            containsWildcard: boolean,
            objectDefns: DataViewNamedObjectDefinition[]): boolean {
            debug.assertValue(dataViewCategorical, 'dataViewCategorical');
            debug.assertValue(objectDescriptors, 'objectDescriptors');
            debug.assertValue(selector, 'selector');
            debug.assertAnyValue(rules, 'rules');
            debug.assertValue(containsWildcard, 'containsWildcard');
            debug.assertValue(objectDefns, 'objectDefns');

            let dataViewCategoricalValues = dataViewCategorical.values;
            if (!dataViewCategoricalValues || !dataViewCategoricalValues.identityFields)
                return;

            if (!Selector.matchesKeys(selector, [dataViewCategoricalValues.identityFields]))
                return;

            let valuesGrouped = dataViewCategoricalValues.grouped();
            if (!valuesGrouped)
                return;

            let foundMatch: boolean;
            for (let i = 0, len = valuesGrouped.length; i < len; i++) {
                let valueGroup = valuesGrouped[i];
                let selectorMetadata = selector.metadata;
                let valuesInGroup = valueGroup.values;
                if (containsWildcard || Selector.matchesData(selector, [valueGroup.identity])) {
                    let objects = DataViewObjectEvaluationUtils.evaluateDataViewObjects(objectDescriptors, objectDefns);
                    if (objects) {
                        // TODO: This mutates the DataView -- the assumption is that prototypal inheritance has already occurred.  We should
                        // revisit this, likely when we do lazy evaluation of DataView.

                        if (selectorMetadata) {
                            for (let j = 0, jlen = valuesInGroup.length; j < jlen; j++) {
                                let valueColumn = valuesInGroup[j],
                                    valueSource = valueColumn.source;
                                if (valueSource.queryName === selectorMetadata) {
                                    let valueSourceOverwrite = Prototype.inherit(valueSource);
                                    valueSourceOverwrite.objects = objects;
                                    valueColumn.source = valueSourceOverwrite;

                                    foundMatch = true;
                                    break;
                                }
                            }
                        }
                        else {
                            valueGroup.objects = objects;
                            setGrouped(dataViewCategoricalValues, valuesGrouped);

                            foundMatch = true;
                        }
                    }

                    if (!containsWildcard)
                        return true;
                }
            }

            return foundMatch;
        }

        function evaluateMetadataRepetition(
            dataView: DataView,
            objectDescriptors: DataViewObjectDescriptors,
            selector: Selector,
            objectDefns: DataViewNamedObjectDefinition[]): void {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(objectDescriptors, 'objectDescriptors');
            debug.assertValue(selector, 'selector');
            debug.assertValue(objectDefns, 'objectDefns');

            // TODO: This mutates the DataView -- the assumption is that prototypal inheritance has already occurred.  We should
            // revisit this, likely when we do lazy evaluation of DataView.
            let columns = dataView.metadata.columns,
                metadataId = selector.metadata;
            for (let i = 0, len = columns.length; i < len; i++) {
                let column = columns[i];
                if (column.queryName === metadataId) {
                    let objects = DataViewObjectEvaluationUtils.evaluateDataViewObjects(objectDescriptors, objectDefns);
                    if (objects)
                        column.objects = objects;
                }
            }
        }

        /** Attempts to find a column that can possibly match the selector. */
        function findSelectedCategoricalColumn(dataViewCategorical: DataViewCategorical, selector: Selector) {
            debug.assertValue(dataViewCategorical.categories[0], 'dataViewCategorical.categories[0]');

            let categoricalColumn = dataViewCategorical.categories[0];
            if (!categoricalColumn.identityFields)
                return;
            if (!Selector.matchesKeys(selector, [categoricalColumn.identityFields]))
                return;

            let identities = categoricalColumn.identity,
                targetColumn: DataViewCategoricalColumn = categoricalColumn;

            let selectedMetadataId = selector.metadata;
            if (selectedMetadataId) {
                let valueColumns = dataViewCategorical.values;
                if (valueColumns) {
                    for (let i = 0, len = valueColumns.length; i < len; i++) {
                        let valueColumn = valueColumns[i];
                        if (valueColumn.source.queryName === selectedMetadataId) {
                            targetColumn = valueColumn;
                            break;
                        }
                    }
                }
            }

            return {
                column: targetColumn,
                identities: identities,
            };
        }

        function findSelectorForRuleInput(dataView: DataView, selectorRoles: string[]): Selector {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(selectorRoles, 'selectorRoles');

            if (selectorRoles.length !== 1)
                return;

            let dataViewCategorical = dataView.categorical;
            if (!dataViewCategorical)
                return;

            let categories = dataViewCategorical.categories;
            if (!categories || categories.length !== 1)
                return;

            let categoryColumn = categories[0],
                categoryRoles = categoryColumn.source.roles,
                categoryIdentityFields = categoryColumn.identityFields;
            if (!categoryRoles || !categoryIdentityFields || !categoryRoles[selectorRoles[0]])
                return;

            return { data: [DataViewScopeWildcard.fromExprs(categoryIdentityFields)] };
        }

        function findRuleInputCategoricalColumn(dataViewCategorical: DataViewCategorical, inputRole: string): DataViewCategoricalColumn {
            debug.assertValue(dataViewCategorical, 'dataViewCategorical');

            return findRuleInputInCategoricalColumns(dataViewCategorical.values, inputRole) ||
                findRuleInputInCategoricalColumns(dataViewCategorical.categories, inputRole);
        }

        function findRuleInputInCategoricalColumns(columns: DataViewCategoricalColumn[], inputRole: string): DataViewCategoricalColumn {
            debug.assertAnyValue(columns, 'columns');

            if (!columns)
                return;

            for (let i = 0, len = columns.length; i < len; i++) {
                let column = columns[i],
                    roles = column.source.roles;
                if (!roles || !roles[inputRole])
                    continue;

                return column;
            }
        }

        /** Attempts to find the value range for the single column with the given inputRole. */
        function findRuleInputColumnNumberRange(dataView: DataView, inputRole: string): NumberRange {
            debug.assertValue(dataView, 'dataView');
            debug.assertValue(inputRole, 'inputRole');

            // NOTE: This implementation currently only supports categorical DataView, becuase that's the
            // only scenario that has custom colors, as of this writing.  This would be rewritten to be more generic
            // as required, when needed.
            let dataViewCategorical = dataView.categorical;
            if (!dataViewCategorical)
                return;

            let values = dataViewCategorical.values;
            if (!values)
                return;

            for (let i = 0, len = values.length; i < len; i++) {
                let valueCol = values[i],
                    valueColRoles = valueCol.source.roles;

                if (!valueColRoles || !valueColRoles[inputRole])
                    continue;

                let min = valueCol.min;
                if (min === undefined)
                    min = valueCol.minLocal;
                if (min === undefined)
                    continue;

                let max = valueCol.max;
                if (max === undefined)
                    max = valueCol.maxLocal;
                if (max === undefined)
                    continue;

                return { min: min, max: max };
            }
        }

        export function createTransformActions(
            queryMetadata: QueryMetadata,
            visualElements: VisualElement[],
            objectDescs: DataViewObjectDescriptors,
            objectDefns: DataViewObjectDefinitions): DataViewTransformActions {
            debug.assertAnyValue(queryMetadata, 'queryMetadata');
            debug.assertAnyValue(visualElements, 'visualElements');
            debug.assertAnyValue(objectDescs, 'objectDescs');
            debug.assertAnyValue(objectDefns, 'objectDefns');

            if ((!queryMetadata || ArrayExtensions.isUndefinedOrEmpty(queryMetadata.Select)) &&
                ArrayExtensions.isUndefinedOrEmpty(visualElements) &&
                !objectDefns)
                return;

            let transforms: DataViewTransformActions = {};
            if (queryMetadata) {
                let querySelects = queryMetadata.Select;
                if (querySelects) {
                    let transformSelects: DataViewSelectTransform[] = transforms.selects = [];
                    for (let i = 0, len = querySelects.length; i < len; i++) {
                        let selectMetadata = querySelects[i],
                            selectTransform = toTransformSelect(selectMetadata, i);
                        transformSelects.push(selectTransform);

                        if (selectTransform.format && objectDescs) {
                            debug.assert(!!selectTransform.queryName, 'selectTransform.queryName should be defined (or defaulted).');

                            let formatStringProp = DataViewObjectDescriptors.findFormatString(objectDescs);
                            if (formatStringProp) {
                                // Select Format strings are migrated into objects
                                if (!objectDefns)
                                    objectDefns = {};

                                DataViewObjectDefinitions.ensure(
                                    objectDefns,
                                    formatStringProp.objectName,
                                    { metadata: selectTransform.queryName })
                                .properties[formatStringProp.propertyName] = SQExprBuilder.text(selectTransform.format);
                            }
                        }
                    }
                }
            }

            if (visualElements) {
                let visualElementsLength = visualElements.length;
                if (visualElementsLength > 1)
                    transforms.splits = [];

                for (let i = 0; i < visualElementsLength; i++) {
                    let visualElement = visualElements[i];

                    if (visualElement.Settings && i === 0)
                        objectDefns = upgradeSettingsToObjects(visualElement.Settings, objectDefns);

                    if (visualElement.DataRoles) {
                        if (!transforms.selects)
                            transforms.selects = [];
                        if (!transforms.projectionOrdering)
                            transforms.projectionOrdering = {};

                        populateDataRoles(visualElement.DataRoles, transforms.selects, transforms.projectionOrdering);
                    }

                    if (transforms.splits)
                        transforms.splits.push(populateSplit(visualElement.DataRoles));
                }
            }

            if (objectDefns)
                transforms.objects = objectDefns;

            return transforms;
        }

        function toTransformSelect(select: SelectMetadata, index: number): DataViewSelectTransform {
            debug.assertValue(select, 'select');

            let result: DataViewSelectTransform = {};

            if (select.Restatement)
                result.displayName = select.Restatement;

            if (select.Name)
                result.queryName = select.Name;
            else if (!result.queryName)
                result.queryName = '$select' + index;

            if (select.Format)
                result.format = select.Format;

            if (select.Type) {
                result.type = describeDataType(select.Type, ConceptualDataCategory[select.DataCategory]);
            }

            if (select.kpiStatusGraphic) {
                result.kpiStatusGraphic = select.kpiStatusGraphic;
            }

            return result;
        }

        function describeDataType(type?: SemanticType, category?: string): ValueType {
            type = (type || SemanticType.None);

            let primitiveType = PrimitiveType.Null;
            switch (type) {
                case SemanticType.String:
                    primitiveType = PrimitiveType.Text;
                    break;
                case SemanticType.Number:
                    primitiveType = PrimitiveType.Double;
                    break;
                case SemanticType.Integer:
                    primitiveType = PrimitiveType.Integer;
                    break;
                case SemanticType.Boolean:
                    primitiveType = PrimitiveType.Boolean;
                    break;
                case SemanticType.Date:
                    primitiveType = PrimitiveType.Date;
                    break;
                case SemanticType.DateTime:
                    primitiveType = PrimitiveType.DateTime;
                    break;
                case SemanticType.Time:
                    primitiveType = PrimitiveType.Time;
                    break;
                case SemanticType.Year:
                    primitiveType = PrimitiveType.Integer;
                    debug.assert(!category || category === 'Year', 'Unexpected category for Year type.');
                    category = 'Year';
                    break;
                case SemanticType.Month:
                    primitiveType = PrimitiveType.Integer;
                    debug.assert(!category || category === 'Month', 'Unexpected category for Month type.');
                    category = 'Month';
                    break;
            }

            return ValueType.fromPrimitiveTypeAndCategory(primitiveType, category);
        }

        function populateDataRoles(
            roles: DataRole[],
            selects: DataViewSelectTransform[],
            projectionOrdering: DataViewProjectionOrdering): void {
            debug.assertValue(roles, 'roles');
            debug.assertValue(selects, 'selects');
            debug.assertValue(projectionOrdering, 'projectionOrdering');

            for (let i = 0, len = roles.length; i < len; i++) {
                let role = roles[i],
                    roleProjection = role.Projection,
                    roleName = role.Name;

                // Update the select
                let select = selects[roleProjection];
                if (select === undefined) {
                    fillArray(selects, roleProjection);
                    select = selects[roleProjection] = {};
                }

                let selectRoles = select.roles;
                if (select.roles === undefined)
                    selectRoles = select.roles = {};

                selectRoles[roleName] = true;

                // Update the projectionOrdering
                let projectionOrderingForRole = projectionOrdering[roleName];
                if (projectionOrderingForRole === undefined)
                    projectionOrderingForRole = projectionOrdering[roleName] = [];
                projectionOrderingForRole.push(roleProjection);
            }
        }

        function fillArray(selects: DataViewSelectTransform[], length: number): void {
            debug.assertValue(selects, 'selects');
            debug.assertValue(length, 'length');

            for (let i = selects.length; i < length; i++)
                selects[i] = {};
        }

        function populateSplit(roles: DataRole[]): DataViewSplitTransform {
            debug.assertAnyValue(roles, 'roles');

            let selects: INumberDictionary<boolean> = {};
            let split: DataViewSplitTransform = {
                selects: selects
            };

            if (roles) {
                for (let i = 0, len = roles.length; i < len; i++) {
                    let role = roles[i];
                    selects[role.Projection] = true;
                }
            }

            return split;
        }

        export function createValueColumns(
            values: DataViewValueColumn[] = [],
            valueIdentityFields?: SQExpr[],
            source?: DataViewMetadataColumn): DataViewValueColumns {
            let result = <DataViewValueColumns>values;
            setGrouped(<DataViewValueColumns>values);

            if (valueIdentityFields)
                result.identityFields = valueIdentityFields;

            if (source)
                result.source = source;

            return result;
        }

        function setGrouped(values: DataViewValueColumns, groupedResult?: DataViewValueColumnGroup[]): void {
            values.grouped = groupedResult
                ? () => groupedResult
                : () => groupValues(values);
        }

        /** Group together the values with a common identity. */
        function groupValues(values: DataViewValueColumn[]): DataViewValueColumnGroup[] {
            debug.assertValue(values, 'values');

            let groups: DataViewValueColumnGroup[] = [],
                currentGroup: DataViewValueColumnGroup;

            for (let i = 0, len = values.length; i < len; i++) {
                let value = values[i];

                if (!currentGroup || currentGroup.identity !== value.identity) {
                    currentGroup = {
                        values: []
                    };

                    if (value.identity) {
                        currentGroup.identity = value.identity;

                        let source = value.source;
                        // allow null, which will be formatted as (Blank).
                        if (source.groupName !== undefined)
                            currentGroup.name = source.groupName;
                        else if (source.displayName)
                            currentGroup.name = source.displayName;
                    }

                    groups.push(currentGroup);
                }

                currentGroup.values.push(value);
            }

            return groups;
        }

        function pivotIfNecessary(dataView: DataView, dataViewMappings: DataViewMapping[]): DataView {
            debug.assertValue(dataView, 'dataView');

            let transformedDataView: DataView;
            switch (determineCategoricalTransformation(dataView.categorical, dataViewMappings)) {
                case CategoricalDataViewTransformation.Pivot:
                    transformedDataView = DataViewPivotCategorical.apply(dataView);
                    break;

                case CategoricalDataViewTransformation.SelfCrossJoin:
                    transformedDataView = DataViewSelfCrossJoin.apply(dataView);
                    break;
            }

            return transformedDataView || dataView;
        }

        function determineCategoricalTransformation(categorical: DataViewCategorical, dataViewMappings: DataViewMapping[]): CategoricalDataViewTransformation {
            if (!categorical || ArrayExtensions.isUndefinedOrEmpty(dataViewMappings))
                return;

            let categories = categorical.categories;
            if (!categories || categories.length !== 1)
                return;

            let values = categorical.values;
            if (ArrayExtensions.isUndefinedOrEmpty(values))
                return;

            if (values.grouped().some(vg => !!vg.identity))
                return;

            // If we made it here, the DataView has a single category and no valueGrouping.
            let categoryRoles = categories[0].source.roles;

            for (let i = 0, len = dataViewMappings.length; i < len; i++) {
                let roleMappingCategorical = dataViewMappings[i].categorical;
                if (!roleMappingCategorical)
                    continue;

                if (!hasRolesGrouped(categoryRoles, <DataViewGroupedRoleMapping>roleMappingCategorical.values))
                    continue;

                // If we made it here, the DataView's single category has the value grouping role.
                let categoriesMapping = roleMappingCategorical.categories;
                let hasCategoryRole =
                    hasRolesBind(categoryRoles, <DataViewRoleBindMappingWithReduction>categoriesMapping) ||
                    hasRolesFor(categoryRoles, <DataViewRoleForMappingWithReduction>categoriesMapping);

                if (hasCategoryRole)
                    return CategoricalDataViewTransformation.SelfCrossJoin;

                return CategoricalDataViewTransformation.Pivot;
            }
        }

        function shouldPivotMatrix(matrix: DataViewMatrix, dataViewMappings: DataViewMapping[]): boolean {
            if (!matrix || ArrayExtensions.isUndefinedOrEmpty(dataViewMappings))
                return;

            let rowLevels = matrix.rows.levels;
            if (rowLevels.length < 1)
                return;

            let rows = matrix.rows.root.children;
            if (!rows || rows.length === 0)
                return;

            let rowRoles = rowLevels[0].sources[0].roles;

            for (let i = 0, len = dataViewMappings.length; i < len; i++) {
                let roleMappingMatrix = dataViewMappings[i].matrix;
                if (!roleMappingMatrix)
                    continue;

                if (!hasRolesFor(rowRoles, <DataViewRoleForMappingWithReduction>roleMappingMatrix.rows) &&
                    hasRolesFor(rowRoles, <DataViewRoleForMappingWithReduction>roleMappingMatrix.columns)) {
                    return true;
                }
            }
        }

        function hasRolesBind(roles: { [name: string]: boolean }, roleMapping: DataViewRoleBindMappingWithReduction): boolean {
            if (roles && roleMapping && roleMapping.bind)
                return roles[roleMapping.bind.to];
        }

        function hasRolesFor(roles: { [name: string]: boolean }, roleMapping: DataViewRoleForMappingWithReduction): boolean {
            if (roles && roleMapping && roleMapping.for)
                return roles[roleMapping.for.in];
        }

        function hasRolesGrouped(roles: { [name: string]: boolean }, roleMapping: DataViewGroupedRoleMapping): boolean {
            if (roles && roleMapping && roleMapping.group)
                return roles[roleMapping.group.by];
        }
    }
}
