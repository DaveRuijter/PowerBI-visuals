//-/// <reference path="../_references.ts"/>

module powerbi.visuals {
    import SelectionManager = utility.SelectionManager;

    export interface FishModel {
        x: number;
        y: number;
        color: string;
        size: number;
        tooltip: string[];
        paused: boolean;
        speed: number;
        body: D3.Selection;
        svgShapeId: string;
        selector: SelectionId;
    }

    export interface DecorationModel {
        x: number;
        y: number;
        size: number;
        tooltip: string;
        svgShapeId: string;
        svgElement: D3.Selection;
    }

    export class Aquarium implements IVisual {

        public static capabilities: VisualCapabilities = {
            dataRoles: [
                {
                    name: 'Category',
                    displayName: 'Fish',
                    kind: VisualDataRoleKind.Grouping
                },
                {
                    name: 'Series',
                    displayName: 'Fish Size',
                    kind: VisualDataRoleKind.Measure
                }],
            dataViewMappings: [{
                categories: {
                    for: { in: 'Category' },
                    dataReductionAlgorithm: { top: {} }
                },
                values: {
                    select: [{ bind: { to: 'Series' } }]
                },
                conditions: [
                { 'Category': { max: 1 }, 'Series': { max: 2 } },
                ],
            }],
            objects: {
                general: {
                    displayName: 'Options',
                    properties: {
                        maxSize: {
                            type: { numeric: true },
                            displayName: 'Max Size'
                        },
                        triangleFish: {
                            type: { bool: true },
                            displayName: 'Triangle Fish'
                        }
                    },
                }
            },
        };

        private colors: IDataColorPalette;
        private fish: { [id: string]: FishModel; } = {};
        private viewPort: IViewport;        
        private svg: D3.Selection;
        private mainGroupElement: D3.Selection;
        private decorations: DecorationModel[] = [];
        private reed1: DecorationModel;
        private reed2: DecorationModel;
        private selectedFish: FishModel = null;
        private tooltipcontainer;
        private selectionManager: SelectionManager;
        private stopTimer: boolean = false;

        private aquariumWidth: number = 1000;
        private aquariumHeight: number = 1000;

        private fishTypes: [string] = ['round-fishy', 'triangle-fishy'];

        private static VisualClassName = 'aquarium';

        public init(options: VisualInitOptions): void {
            this.colors = options.style.colorPalette.dataColors;
            this.selectionManager = new SelectionManager({ hostServices: options.host });

            var element = options.element;
            var svg = this.svg = d3.select(element.get(0))
                .append('svg')
                .attr('style', 'background-color:#abcdef;')
                .classed(Aquarium.VisualClassName, true);

            this.tooltipcontainer = $('<div class="tooltip-container" style="display: none;"><div class="arrow top left"></div><div class="tooltip-content-container"><div class="tooltip-row"><div class="tooltip-title-cell" id="series"></div></div><div class="tooltip-row"><div class="tooltip-title-cell" id="label"></div><div class="tooltip-value-cell" id="value"></div></div></div></div>');
            element.append(this.tooltipcontainer);            

            var defs = svg.append('defs');
            var roundFishy = defs.append('g').attr('id', 'round-fishy');
            roundFishy.append('path').attr('d', 'M -177,150 L -283,92 L -283,308 L -177,230	A 190,175 180 1 0 -177, 150 Z').attr('style', 'stroke:#433F3C; stroke-width:12'); //fill:#E0700E; 
            roundFishy.append('circle').attr('cx', '65').attr('cy', '138').attr('r', '70').attr('style', 'fill:white; stroke:#433F3C; stroke-width:12');
            roundFishy.append('circle').attr('cx', '75').attr('cy', '130').attr('r', '40').attr('style', 'fill:#433F3C; stroke:#433F3C; stroke-width:12');
            roundFishy.append('circle').attr('cx', '89').attr('cy', '112').attr('r', '20').attr('style', 'fill:white; stroke:white; stroke-width:12');

            var triangleFishy = defs.append('g').attr('id', 'triangle-fishy');
            triangleFishy.append('path').attr('d', 'M -177,150 L -283,92 L -283,308 L -177,230 L -177,490 L 193,190 L -177, -110 Z').attr('style', 'stroke:#433F3C; stroke-width:12'); //fill:#E0700E; 
            triangleFishy.append('circle').attr('cx', '65').attr('cy', '138').attr('r', '70').attr('style', 'fill:white; stroke:#433F3C; stroke-width:12');
            triangleFishy.append('circle').attr('cx', '75').attr('cy', '130').attr('r', '40').attr('style', 'fill:#433F3C; stroke:#433F3C; stroke-width:12');
            triangleFishy.append('circle').attr('cx', '89').attr('cy', '112').attr('r', '20').attr('style', 'fill:white; stroke:white; stroke-width:12');

            var reed = defs.append('g').attr('id', 'reed');
            reed.append('path').attr('style', 'stroke:#534B48; fill:#006E3A; stroke-width:6').attr('d', 'M114.43965148925781,-17.836837768554688c20.44805908203125,-85.779052734375,48.7747802734375,-117.4727783203125,30.03509521484375,-163.6903076171875s-71.989501953125,-92.864990234375,-100.6168212890625,-132.1533203125s-84.077880859375,-132.08612060546875,-93.10821533203125,-184.7149658203125s21.73486328125,-149.710693359375,34.540283203125,-223.75994873046875s17.45208740234375,48.0499267578125,13.51495361328125,148.6729736328125s55.0174560546875,139.6170654296875,120.139892578125,210.24407958984375s111.7686767578125,125.52093505859375,85.59930419921875,204.2371826171875s-84.91241455078125,179.61480712890625,-90.1044921875,141.164306640625Z');
            reed.append('path').attr('style', 'stroke:#534B48; fill:#28903A; stroke-width:6').attr('d', 'M186.52381896972656,-34.35681915283203c-16.0870361328125,-53.51904296875,-37.96246337890625,-163.650390625,-43.5506591796875,-219.25482177734375s-2.00079345703125,-129.58349609375,-28.53289794921875,-193.72503662109375s14.06134033203125,-172.40170288085938,-3.00372314453125,-234.2723388671875s-45.325927734375,-104.57376098632812,-66.076904296875,-180.20941162109375s100.6195068359375,163.68280029296875,108.1256103515625,259.8017578125s4.10247802734375,39.044403076171875,3.0025634765625,142.6658935546875s10.140380859375,47.303619384765625,28.5333251953125,186.21728515625s40.96484375,338.2679443359375,1.502685546875,238.77667236328125Z');
            reed.append('path').attr('style', 'stroke:#534B48; fill:#004F26; stroke-width:6').attr('d', 'M60.37778091430664,-10.328531265258789c-47.280029296875,-78.24652099609375,-94.27294921875,-263.0853271484375,-90.104736328125,-354.411865234375s59.7353515625,-244.06985473632812,91.60626220703125,-274.8194580078125s-19.64837646484375,245.58934020996094,-33.03839111328125,315.3660888671875s49.837890625,246.51153564453125,31.536865234375,313.865234375Z');
            reed.append('path').attr('style', 'stroke:#534B48; fill:#28903A; stroke-width:6').attr('d', 'M85.9068603515625,-11.830490112304688c0.37860107421875,-0.7492674589157104,1.1295166015625,-2.25103759765625,1.50177001953125,-3.00347900390625s36.43194580078125,-180.66912841796875,36.0418701171875,-241.781005859375s-46.2655029296875,-170.63031005859375,-12.01373291015625,-225.26177978515625s81.49969482421875,-178.17620849609375,141.163818359375,-222.25827026367188s-139.2791748046875,327.9466857910156,-156.1815185546875,444.5165710449219s-38.55157470703125,303.08831787109375,-10.51220703125,247.7879638671875Z');

            var pebble = defs.append('g').attr('id', 'pebble');
            pebble.append('ellipse')
                .attr('cx', 0)
                .attr('cy', 0)
                .attr('rx', 30)
                .attr('ry', 20)
                .attr('style', 'fill:#D2CFC0; stroke:#534B48; stroke-width:3');
                            
            this.mainGroupElement = svg.append('g');

            //make pebbles
            for (var i = 0; i < 50; i++) {
                //add pebbles across the entire width of the aquarium with a bit of X and Y randomness
                //splice them in random order so they overlap randomly
                this.decorations.splice(Math.random() * this.decorations.length,0, {
                        x: (i * 20) +((Math.random() * 10) - 5),
                        y: ((i%2) * 30) + 930 + (Math.random() * 40),
                        size: 1,
                        svgShapeId: "pebble",
                        svgElement: null,
                        tooltip: null
                    });
            }
            
            //add reeds to support up to 2 series of data
            this.reed1 = {
                x: 40,
                y: 1000,
                size: 0.4,
                tooltip: null,
                svgShapeId: "reed",
                svgElement: null
            };
            this.decorations.push(this.reed1);
            this.reed2 = {
                x: 840,
                y: 1000,
                size: 0.8,
                tooltip: null,
                svgShapeId: "reed",
                svgElement: null
            };
            this.decorations.push(this.reed2);

            d3.timer(() => {            
                this.draw(this.viewPort);
                return this.stopTimer;
            });
        }

        public onResizing(viewport: IViewport) { /* This API will be depricated */ }

        /** Notifies the visual that it is being destroyed, and to do any cleanup necessary (such as unsubscribing event handlers). */
        public destroy() {
            this.stopTimer = true;
        }

        public updateModel(dataView: DataView, colors: IDataColorPalette) {         
            if(!dataView) return;
               
            var table = dataView.table;
            var catDv: DataViewCategorical = dataView.categorical;
            if (!table || !catDv || !catDv.values || !catDv.categories){
                this.reed1.size = this.reed2.size = 0;
                for (var key in this.fish) {
                    this.removeFish(key);
                }
                return;  
            }
            var cat = catDv.categories[0];

            /*
            if (dataView.metadata && dataView.metadata.objects) {
                var objects = dataView.metadata.objects;
                var maxSize = DataViewObjects.getValue<number>(objects, { objectName: 'general', propertyName: 'maxSize' });
                var triangleFish = DataViewObjects.getValue<boolean>(objects, { objectName: 'general', propertyName: 'triangleFish' });
            }*/

            //work out the max of each series and the max of all series
            var tableMax = 1;
            var seriesMax: { [id: number]: number; } = {};
            //for (var series = 1; series <= 2 && series < table.columns.length; series++) {
            for (var series = 0, lenC = catDv.values.length; series <= 2 && series < lenC; series++){
                //max of this series
                seriesMax[series] = 0;
                //for (var row = 0, len = table.rows.length; row < len; row++) {
                //catDv.values[series].length
                for (var row = 0, len = catDv.values[series].values.length; row < len; row++) {
                    //var value = table.rows[row][series];
                    var value = catDv.values[series].values[row];
                    if (value) {
                        seriesMax[series] = Math.max(seriesMax[series], value);
                    }
                }

                //max of table
                tableMax = Math.max(tableMax, seriesMax[series]);
            }                        

            if(seriesMax[0] && !isNaN(seriesMax[0])){
                this.reed1.size = (seriesMax[0] / tableMax) * 0.8;
            }
            else
                this.reed1.size = 0;
                
            if(seriesMax[1] && !isNaN(seriesMax[1])){
                this.reed2.size = (seriesMax[1] / tableMax) * 0.8;
            }
            else
                this.reed2.size = 0;
            //this.reed2.size = (seriesMax[2] / tableMax) * 0.8;

            //update our dataset with the new dataset
            var updatedFishIds: { [id: string]: string; } = {};
            //for (var row = 0, len = table.rows.length; row < len; row++) {
            //    for (var series = 1; series <= table.columns.length; series++) {
            for (var series = 0, countSeries = catDv.values.length; series < countSeries; series++) {
                for (var row = 0, lenRows = catDv.values[0].values.length; row < lenRows; row++) {
                    //var value = table.rows[row][series];
                    var value = catDv.values[series].values[row];
                    //there might be null values which need to be ignored
                    if (!isNaN(value)) {
                        //create a unique identifier per series
                        //var label = table.rows[row][0];
                        var label = catDv.categories[0].values[series];
                        var fishId = label + ", series " + series + "_" + row;
                        var valueFormated = "$" + value.toFixed(2).replace(/\B(?=(\d{3})+\b)/g, ",");                    
                        var labelColour = colors.getColorByIndex(row).value;
                        //var seriesLabel = dataView.metadata.columns[series].displayName;
                        var seriesLabel = catDv.values[series].source.displayName;
                        var tooltip = [
                            seriesLabel,
                            label,
                            valueFormated
                        ];
                        var selector = SelectionId.createWithId(cat.identity[row]);

                        updatedFishIds[fishId] = fishId;
                        this.addOrUpdateFish(fishId, value / tableMax, labelColour, tooltip, selector, this.fishTypes[(series)%2]);
                    }
                }          
            }

            //cull out the supurfulous fish
            for (var key in this.fish) {
                if (!updatedFishIds[key]) {
                    this.removeFish(key);
                }
            }

        }

        private addOrUpdateFish(fishId: string, value: number, colour: string, tooltip: string[], selector: SelectionId, fishType: string) {
            var fishy = this.fish[fishId];
            if (fishy == null) {
                fishy = this.fish[fishId] = {
                    x: 80 + Math.random() * 840,
                    y: 80 + Math.random() * 760,
                    color: colour,
                    tooltip: tooltip,
                    size: 0,
                    speed: 0,
                    paused: false,
                    body: null,
                    svgShapeId: fishType,
                    selector: selector
                };
            }
            fishy.tooltip = tooltip;
            fishy.size = value * 0.3;                          
            fishy.speed = value * (fishy.speed < 0 ? -1 : 1); //don't change their direction            
            if (fishy.size < 0) fishy.speed /= 2; // #ded
        }

        private removeFish(fishId: string) {
            this.fish[fishId].body.remove();
            delete this.fish[fishId];
        }

        public update(options: VisualUpdateOptions) {
            if (!options.dataViews && !options.dataViews[0]) return;

            this.updateModel(options.dataViews[0], this.colors);
            this.viewPort = options.viewport;
            this.draw(this.viewPort);
        }

        public moveFishy(fishy: FishModel) {
            if (!fishy) {
                return;
            }

            //if we are the selected fish, don't move
            if (this.selectedFish === fishy) {
                return;
            }

            if (fishy.paused) {
                return;
            }

            //float to the top if #ded
            if (fishy.size < 0 && fishy.y > 40) {                
                fishy.y -= 0.5;                
            }

            //move
            fishy.x -= fishy.speed;

            //turn around if we hit the edge
            if (fishy.x < 80 || fishy.x > this.aquariumWidth-80) {
                fishy.speed *= -1;
            }
        }

        private setTooltip(fishy: FishModel) {
            if (fishy == null) {
                this.tooltipcontainer.hide(20);
                return;
            }
            this.tooltipcontainer.find('#series').text(fishy.tooltip[0]);
            this.tooltipcontainer.find('#label').text(fishy.tooltip[1]);
            this.tooltipcontainer.find('#value').text(fishy.tooltip[2]);
            var mouse = d3.mouse(this.svg.node());
            this.tooltipcontainer.css({ top: mouse[1] - 15, left: mouse[0] + 15 });
            this.tooltipcontainer.show();
        }

        private selectFish(fishy: FishModel) {
            this.selectedFish = fishy;
            this.selectionManager.select(fishy.selector);
        }

        private resetSelection() {
            this.selectedFish = null;
            //raise event to other charts
            this.selectionManager.clear();
        }

        private drawFishy(fishy: FishModel, scaleX: number, scaleY: number) {
            var x = fishy.x * scaleX;
            var y = fishy.y * scaleY;
            var size = fishy.size * scaleX;

            //create all the parts if this is the first time we are drawing the fish
            if (fishy.body == null) {
                fishy.body = this.mainGroupElement
                    .append('use')
                    .attr('xlink:href', '#' + fishy.svgShapeId)
                    .on('click', () => {
                        //double select to unselect
                        if (this.selectedFish === fishy) {
                            this.resetSelection();
                        } else {
                            this.selectFish(fishy);
                        }

                        d3.event.stopPropagation();                        
                    })
                    .on('mousemove', () =>
                    {
                        fishy.paused = true;
                        this.setTooltip(fishy);
                    })
                    .on('mouseout', () =>
                    {
                        fishy.paused = false;
                        this.setTooltip(null);
                    });
            }
            
            //otherwise just update positions
            var transform = 'translate(' + x + ',' + y + ')';
            transform += ' scale(' + size + ')';
            if (fishy.speed > 0) {
                transform += ' scale(-1, 1)';
            }
            fishy.body.attr('transform', transform);
            fishy.body.attr('style', 'fill:' + fishy.color);            

            var faded = false;
            if (this.selectedFish != null && this.selectedFish !== fishy) {
                faded = true;
            }

            if (faded) {
                fishy.body.style('opacity', 0.5);
            } else {
                fishy.body.style('opacity', 1);
            }
        }

        private drawDecoration(decoration: DecorationModel, scaleX: number, scaleY:number) {
            var x = decoration.x * scaleX;
            var y = decoration.y * scaleY;
            var size = decoration.size * scaleX;

            //bail out if anything is wrong
            if (!x || !y/* || !size*/) {
                return;
            }

            if (decoration.svgElement == null) {
                decoration.svgElement = this.mainGroupElement.append('use').attr('xlink:href', '#' + decoration.svgShapeId);
            }
            
            var transform = 'translate(' + x + ',' + y + ')';
            transform += ' scale(' + size + ')';
            decoration.svgElement.attr('transform', transform);
        }

        public draw(viewport: IViewport) {
            if (!viewport)
                return;

            this.svg
                .attr({
                    'height': viewport.height,
                    'width': viewport.width
                })
                .on('click', () => { this.resetSelection(); });

            var scaleX = viewport.width / this.aquariumWidth;
            var scaleY = viewport.height / this.aquariumHeight;

            //decorations
            if (this.decorations != null) {
                for (var i = 0, len = this.decorations.length; i < len; i++) {
                    this.drawDecoration(this.decorations[i], scaleX, scaleY);
                }
            }

            //fish
            for (var fishId in this.fish) {
                this.moveFishy(this.fish[fishId]);
                this.drawFishy(this.fish[fishId], scaleX, scaleY);
            }
            
        }
    }
}