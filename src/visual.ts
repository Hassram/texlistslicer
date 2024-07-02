/*
*  Power BI Visual CLI
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
"use strict";

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;
import FilterAction = powerbi.FilterAction;
import { IAdvancedFilter, AdvancedFilter, BasicFilter, IBasicFilter } from "powerbi-models";

import { Selection as d3Selection, select as d3Select } from "d3-selection";

import { TextFilterSettingsModel } from "./settings";

import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import clearButtonSVG from "./clearButton"
import searchButtonSVG from "./searchButton"

const pxToPt = 0.75,
  fontPxAdjSml = 20,
  fontPxAdjStd = 24,
  fontPxAdjLrg = 26;


export class Visual implements IVisual {

  private target: HTMLElement;
  private searchUi: d3Selection<HTMLDivElement, any, any, any>;
  private buttonGroup: d3Selection<HTMLDivElement, any, any, any>;
  private searchBox: d3Selection<HTMLTextAreaElement, any, any, any>;

  private clearButtonSVG: d3Selection<SVGSVGElement, any, any, any>
  private searchButtonSVG: d3Selection<SVGSVGElement, any, any, any>
  private column: powerbi.DataViewMetadataColumn;
  private host: powerbi.extensibility.visual.IVisualHost;
  private events: IVisualEventService;
  private formattingSettingsService: FormattingSettingsService;
  private formattingSettings: TextFilterSettingsModel;
  private localizationManager: ILocalizationManager;
  private viewHeight: number


  constructor(options: VisualConstructorOptions) {
    this.events = options.host.eventService;
    this.target = options.element;
    this.viewHeight = 100



    this.searchUi = d3Select(this.target)
      .append("div")
      // .style("margin", "2px")          
      // .style("border", "4px solid black")
      .classed("my-div", true);



    this.searchBox = this.searchUi
      .append("textarea")
      .attr("aria-label", "Enter your search")
      .attr("type", "text")
      .attr("name", "search-field")
      .attr("autofocus", true)
      .attr("tabindex", 0)
      .classed("accessibility-compliant", true)
      .classed("searchUi", true)
      .style("resize", "none")
      .classed("border-on-focus", true)
      .style("height", "10% !important");

    this.buttonGroup = this.searchUi
      .append("div")
      .classed("button-group", true)



    this.searchButtonSVG = this.buttonGroup
      .append("svg")
      .attr('fill','black')
      .attr("width", 32)
      .attr("height", 32);

    this.searchButtonSVG
      .html(searchButtonSVG)

    this.clearButtonSVG = this.buttonGroup
      .append("svg")
      .attr("width", 32)
      .attr('fill', this.searchBox.property("value") == "" ? 'grey' : 'black')
      .attr("height", 32);

    this.clearButtonSVG
      .html(clearButtonSVG)



    // custom visuals require 2 tabs to get to the first focusable element (by design)
    // event listener below is designed to overcome this limitation
    window.addEventListener('focus', (event) => {
      // focus entered from the parent window
      if (event.target === window) {
        this.searchBox.node().focus();
      }
    })

    // focus input after clear button
    this.clearButtonSVG.on("keydown", event => {
      if (event.key === "Tab") {
        event.preventDefault();
        this.searchBox.node()?.focus();
        event.stopPropagation();
      }
    })

    this.searchBox.on("keydown", (event) => {
      if (event.key === "Enter") {
        this.performSearch(this.searchBox.property("value"));
      }
    });

    // these click handlers also handle "Enter" key press with keyboard navigation
    // this.searchButton
    //   .on("click", () => this.performSearch(this.searchBox.property("value")));
    this.clearButtonSVG
      .on("click", () => this.clearSearch());


    this.clearButtonSVG
      .on("click", () => this.clearSearch());

    this.searchButtonSVG
      .on('click', () => this.performSearch(this.searchBox.property("value")))

    d3Select(this.target)
      .on("contextmenu", (event) => {
        const
          mouseEvent: MouseEvent = event,
          selectionManager = options.host.createSelectionManager();
        selectionManager.showContextMenu({}, {
          x: mouseEvent.clientX,
          y: mouseEvent.clientY
        });
        mouseEvent.preventDefault();
      });

    this.localizationManager = options.host.createLocalizationManager()
    this.formattingSettingsService = new FormattingSettingsService(this.localizationManager);

    this.host = options.host;
  }

  public getFormattingModel(): powerbi.visuals.FormattingModel {
    // removes border color
    if (this.formattingSettings?.textBox.enableBorder.value === false) {
      this.formattingSettings.removeBorderColor();
    }
    const model = this.formattingSettingsService.buildFormattingModel(this.formattingSettings);

    return model;
  }

  public update(options: VisualUpdateOptions) {
    const width: number = options.viewport.width;
    this.viewHeight = options.viewport.height - 32 - 20;
    this.searchUi
      .attr("height", '{height}px !important')

    this.events.renderingStarted(options);
    this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(TextFilterSettingsModel, options.dataViews);
    const metadata = options.dataViews && options.dataViews[0] && options.dataViews[0].metadata;
    const newColumn = metadata && metadata.columns && metadata.columns[0];
    let searchText = "";
    this.updateUiSizing();

    // We had a column, but now it is empty, or it has changed.
    if (options.dataViews && options.dataViews.length > 0 && this.column && (!newColumn || this.column.queryName !== newColumn.queryName)) {
      this.performSearch("");

      // Well, it hasn't changed, then lets try to load the existing search text.
    } else if (options?.jsonFilters?.length > 0) {
      searchText = `${(<IAdvancedFilter[]>options.jsonFilters).map((f) => f.conditions.map((c) => c.value)).join(" ")}`;
    }

    this.searchBox.property("value", searchText);
    this.column = newColumn;

    this.events.renderingFinished(options);

  }

  /**
   * Ensures that the UI is sized according to the specified properties (or defaults, if not overridden).
   */
  private updateUiSizing() {
    const
      textBox = this.formattingSettings?.textBox,
      fontSize = textBox.font.fontSize.value,
      fontScaleSml = Math.floor((fontSize / pxToPt) + fontPxAdjSml),
      fontScaleStd = Math.floor((fontSize / pxToPt) + fontPxAdjStd),
      fontScaleLrg = Math.floor((fontSize / pxToPt) + fontPxAdjLrg);
    this.searchUi
      .style('height', this.viewHeight + "px")
      .style("margin", 5)
      .style("padding", 5)
      .style('font-size', `${fontSize}pt`)
      .style('font-family', textBox.font.fontFamily.value);
    this.searchBox
      .attr('placeholder', this.localizationManager.getDisplayName(textBox.placeholderTextKey))
      .style('width', `calc(100% - ${fontScaleStd}px)`)
      .style('padding-right', `${fontScaleStd}px`)
      .style('border-style', textBox.enableBorder.value && 'solid' || 'none')
      .style('border-color', textBox.borderColor.value.value)
      .style('font-size', `${fontSize}pt`)
      .style('color', textBox.textColor.value.value);
  }


  private parseStringToList(input: string): string[] {
    const normalizedInput = input.replace(/\s+/g, ',');
    const list = normalizedInput.split(',');
    const filteredList = list.map(item => item.trim())
      .filter(trimmedItem => trimmedItem !== '');

    return filteredList;
  }



  /** 
   * Perfom search/filtering in a column
   * @param {string} text - text to filter on
   */
  public performSearch(text: string) {

    this.clearButtonSVG
      .attr('fill', this.searchBox.property("value") == "" ? 'grey' : 'black')
    if (this.column) {
      const isBlank = !text.trim()
      // console.log("text entered is:", text, "which is:", isBlank)
      const target = {
        table: this.column.queryName.substr(0, this.column.queryName.indexOf(".")),
        column: this.column.queryName.substr(this.column.queryName.indexOf(".") + 1)
      };

      const filter: any = null;
      const action = FilterAction.merge;

      const basicFilter: IBasicFilter = {
        $schema: 'https://powerbi.com/product/schema#basic',
        target,
        operator: "In",
        values: this.parseStringToList(text),
        filterType: 1
      };


      if (this.parseStringToList(text).length > 0) {

        this.host.applyJsonFilter({
          $schema: 'https://powerbi.com/product/schema#basic',
          target,
          operator: "In",
          values: this.parseStringToList(text),
          filterType: 1
        }, "general", "filter", action);
      } else {
        this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
      }

      this.searchBox.property("value", text);
    }


  }

  public clearSearch() {
    this.clearButtonSVG
    .attr('fill','Gray')

    this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
    // console.log("Clear button was hit XX")

  }
}