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
    export interface TreemapAnimationOptions extends IAnimationOptions {
        viewModel: TreemapData;
        nodes: D3.Layout.GraphNode[];
        highlightNodes: D3.Layout.GraphNode[];
        majorLabeledNodes: D3.Layout.GraphNode[];
        minorLabeledNodes: D3.Layout.GraphNode[];
        shapeGraphicsContext: D3.Selection;
        labelGraphicsContext: D3.Selection;
        layout: ITreemapLayout;
        labelSettings: VisualDataLabelsSettings;
    }

    export interface TreemapAnimationResult extends IAnimationResult {
        shapes: D3.UpdateSelection;
        highlightShapes: D3.UpdateSelection;
        majorLabels: D3.UpdateSelection;
        minorLabels: D3.UpdateSelection;
    }

    export type ITreemapAnimator = IAnimator<IAnimatorOptions, TreemapAnimationOptions, TreemapAnimationResult>;

    export class WebTreemapAnimator extends BaseAnimator<IAnimatorOptions, TreemapAnimationOptions, TreemapAnimationResult> implements ITreemapAnimator {
        previousViewModel: TreemapData;

        constructor(options?: IAnimatorOptions) {
            super(options);
        }

        public animate(options: TreemapAnimationOptions): TreemapAnimationResult {
            let result: TreemapAnimationResult = {
                failed: true,
                shapes: null,
                highlightShapes: null,
                majorLabels: null,
                minorLabels: null,
            };

            let viewModel = options.viewModel;
            let previousViewModel = this.previousViewModel;

            if (!previousViewModel) {
                // This is the initial drawing of the chart, which has no special animation for now.
            }
            else if (viewModel.hasHighlights && !previousViewModel.hasHighlights) {
                result = this.animateNormalToHighlighted(options);
            }
            else if (viewModel.hasHighlights && previousViewModel.hasHighlights) {
                result = this.animateHighlightedToHighlighted(options);
            }
            else if (!viewModel.hasHighlights && previousViewModel.hasHighlights) {
                result = this.animateHighlightedToNormal(options);
            }

            this.previousViewModel = viewModel;
            return result;
        }

        private animateNormalToHighlighted(options: TreemapAnimationOptions): TreemapAnimationResult {
            let hasSelection = false;
            let hasHighlights = true;

            let shapes = this.animateDefaultShapes(options.shapeGraphicsContext, options.nodes, hasSelection, hasHighlights, options.layout);

            let highlightShapes = options.shapeGraphicsContext.selectAll('.' + Treemap.HighlightNodeClassName)
                .data(options.highlightNodes, (d: TreemapNode) => d.key + "highlight");

            highlightShapes.enter().append('rect')
                .attr('class', options.layout.highlightShapeClass)
                .attr(options.layout.shapeLayout); // Start using the normal shape layout

            highlightShapes
                .style("fill", (d: TreemapNode) => Treemap.getFill(d, true))
                .style("fill-opacity", (d: TreemapNode) => Treemap.getFillOpacity(d, hasSelection, hasHighlights, true))
                .transition()
                .duration(this.animationDuration)
                .attr(options.layout.highlightShapeLayout); // Animate to the highlighted positions

            highlightShapes.exit().remove();

            let majorLabels = this.animateDefaultMajorLabels(options.labelGraphicsContext, options.majorLabeledNodes, options.labelSettings, options.layout);
            let minorLabels = this.animateDefaultMinorLabels(options.labelGraphicsContext, options.minorLabeledNodes, options.labelSettings, options.layout);

            return {
                failed: false,
                shapes: shapes,
                highlightShapes: highlightShapes,
                majorLabels: majorLabels,
                minorLabels: minorLabels,
            };
        }

        private animateHighlightedToHighlighted(options: TreemapAnimationOptions): TreemapAnimationResult {
            let hasSelection = false;
            let hasHighlights = true;

            let shapes = this.animateDefaultShapes(options.shapeGraphicsContext, options.nodes, hasSelection, hasHighlights, options.layout);

            options.shapeGraphicsContext.selectAll('.' + Treemap.HighlightNodeClassName)
                .data(options.highlightNodes, (d: TreemapNode) => d.key + "highlight");

            let highlightShapes = this.animateDefaultHighlightShapes(options.shapeGraphicsContext, options.highlightNodes, hasSelection, hasHighlights, options.layout);
            
            let majorLabels = this.animateDefaultMajorLabels(options.labelGraphicsContext, options.majorLabeledNodes, options.labelSettings, options.layout);
            let minorLabels = this.animateDefaultMinorLabels(options.labelGraphicsContext, options.minorLabeledNodes, options.labelSettings, options.layout);

            return {
                failed: false,
                shapes: shapes,
                highlightShapes: highlightShapes,
                majorLabels: majorLabels,
                minorLabels: minorLabels,
            };
        }

        private animateHighlightedToNormal(options: TreemapAnimationOptions): TreemapAnimationResult {
            let hasSelection = options.interactivityService ? options.interactivityService.hasSelection() : false;

            let shapes = options.shapeGraphicsContext.selectAll('.' + Treemap.TreemapNodeClassName)
                .data(options.nodes, (d: TreemapNode) => d.key);

            shapes.enter().append('rect')
                .attr('class', options.layout.shapeClass);

            shapes
                .transition()
                .duration(this.animationDuration)
                .style("fill", (d: TreemapNode) => Treemap.getFill(d, false))
                .style("fill-opacity", (d: TreemapNode) => ColumnUtil.getFillOpacity(d.selected, false, d.selected, !d.selected))
                .attr(options.layout.shapeLayout)
                .transition()
                .duration(0)
                .delay(this.animationDuration)
                .style("fill-opacity", (d: TreemapNode) => ColumnUtil.getFillOpacity(d.selected, false, hasSelection, false));

            shapes.exit().remove();

            let highlightShapes = options.shapeGraphicsContext.selectAll('.' + Treemap.HighlightNodeClassName)
                .data(options.nodes, (d: TreemapNode) => d.key + "highlight");

            highlightShapes.enter().append('rect')
                .attr('class', options.layout.highlightShapeClass);

            highlightShapes
                .style("fill", (d: TreemapNode) => Treemap.getFill(d, true))
                .style("fill-opacity", (d: TreemapNode) => ColumnUtil.getFillOpacity(d.selected, true, d.selected, !d.selected))
                .transition()
                .duration(this.animationDuration)
                .attr(hasSelection ? options.layout.zeroShapeLayout : options.layout.shapeLayout) // Animate to the normal shape layout or zero shape layout depending on whether we have a selection or not
                .remove();

            highlightShapes.exit().remove();

            let majorLabels = this.animateDefaultMajorLabels(options.labelGraphicsContext, options.majorLabeledNodes, options.labelSettings, options.layout);
            let minorLabels = this.animateDefaultMinorLabels(options.labelGraphicsContext, options.minorLabeledNodes, options.labelSettings, options.layout);

            return {
                failed: false,
                shapes: shapes,
                highlightShapes: highlightShapes,
                majorLabels: majorLabels,
                minorLabels: minorLabels,
            };
        }

        private animateDefaultShapes(context: D3.Selection, nodes: D3.Layout.GraphNode[], hasSelection: boolean, hasHighlights: boolean, layout: ITreemapLayout): D3.UpdateSelection {
            let isHighlightShape = false;
            let shapes = context.selectAll('.' + Treemap.TreemapNodeClassName)
                .data(nodes, (d: TreemapNode) => d.key);

            shapes.enter().append('rect')
                .attr('class', layout.shapeClass);

            shapes
                .transition()
                .duration(this.animationDuration)
                .style("fill", (d: TreemapNode) => Treemap.getFill(d, isHighlightShape))
                .style("fill-opacity", (d: TreemapNode) => Treemap.getFillOpacity(d, hasSelection, hasHighlights, isHighlightShape))
                .attr(layout.shapeLayout);

            shapes.exit().remove();

            return shapes;
        }

        private animateDefaultHighlightShapes(context: D3.Selection, nodes: D3.Layout.GraphNode[], hasSelection: boolean, hasHighlights: boolean, layout: ITreemapLayout): D3.UpdateSelection {
            let isHighlightShape = true;
            let highlightShapes = context.selectAll('.' + Treemap.HighlightNodeClassName)
                .data(nodes, (d) => d.key + "highlight");

            highlightShapes.enter().append('rect')
                .attr('class', layout.highlightShapeClass);

            highlightShapes
                .transition()
                .duration(this.animationDuration)
                .style("fill", (d: TreemapNode) => Treemap.getFill(d, isHighlightShape))
                .style("fill-opacity", (d: TreemapNode) => Treemap.getFillOpacity(d, hasSelection, hasHighlights, isHighlightShape))
                .attr(layout.highlightShapeLayout);

            highlightShapes.exit().remove();
            return highlightShapes;
        }

        private animateDefaultMajorLabels(context: D3.Selection, nodes: D3.Layout.GraphNode[], labelSettings: VisualDataLabelsSettings, layout: ITreemapLayout): D3.UpdateSelection {
            let labels = context
                .selectAll('.' + Treemap.MajorLabelClassName)
                .data(nodes, (d: TreemapNode) => d.key);

            labels.enter().append('text')
                .attr('class', layout.majorLabelClass);

            labels
                .text(layout.majorLabelText)
                .style('fill', () => labelSettings.labelColor)
                .transition()
                .duration(this.animationDuration)
                .attr(layout.majorLabelLayout);

            labels.exit().remove();

            return labels;
        }

        private animateDefaultMinorLabels(context: D3.Selection, nodes: D3.Layout.GraphNode[], labelSettings: VisualDataLabelsSettings, layout: ITreemapLayout): D3.UpdateSelection {
            let labels = context
                .selectAll('.' + Treemap.MinorLabelClassName)
                .data(nodes, (d: TreemapNode) => d.key);

            labels.enter().append('text')
                .attr('class', layout.minorLabelClass);

            labels
                .text(layout.minorLabelText)
                .style('fill', () => labelSettings.labelColor)
                .transition()
                .duration(this.animationDuration)
                .attr(layout.minorLabelLayout);

            labels.exit().remove();

            return labels;
        }
    }
}