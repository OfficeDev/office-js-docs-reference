| Class | Fields | Description |
|:---|:---|:---|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type: Excel.ChartAxisType, group?: Excel.ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-getitem-member(1))|Returns the specific axis identified by type and group.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-color-member)|HTML color code representing the color of borders in the chart.|
||[lineStyle](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-linestyle-member)|Represents the line style of the border.|
||[weight](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-weight-member)|Represents weight of the border, in points.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[border](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-border-member)|Represents the border format of chart area, which includes color, linestyle, and weight.|
||[font](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-font-member)|Represents the font attributes (font name, font size, color, etc.) for the current object.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#excel-excel-chartformatstring-font-member)|Represents the font attributes, such as font name, font size, and color of a chart characters object.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-height-member)|Specifies the height, in points, of the legend on the chart.|
||[left](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-left-member)|Specifies the left value, in points, of the legend on the chart.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-legendentries-member)|Represents a collection of legendEntries in the legend.|
||[overlay](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-overlay-member)|Specifies if the chart legend should overlap with the main body of the chart.|
||[position](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-position-member)|Specifies the position of the legend on the chart.|
||[showShadow](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-showshadow-member)|Specifies if the legend has a shadow on the chart.|
||[top](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-top-member)|Specifies the top of a chart legend.|
||[visible](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-visible-member)|Specifies if the chart legend is visible.|
||[width](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-width-member)|Specifies the width, in points, of the legend on the chart.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|||
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getcount-member(1))|Returns the number of legend entries in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getitemat-member(1))|Returns a legend entry at the given index.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-items-member)|Gets the loaded child items in this collection.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-clear-member(1))|Clears the line format of a chart element.|
||[lineStyle](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-linestyle-member)|Represents the line style.|
||[weight](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-weight-member)|Represents weight of the line, in points.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[dataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-datalabel-member)|Returns the data label of a chart point.|
||[format](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-format-member)|Encapsulates the format properties chart point.|
||[hasDataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-hasdatalabel-member)|Represents whether a data point has a data label.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerbackgroundcolor-member)|HTML color code representation of the marker background color of a data point (e.g., #FF0000 represents Red).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerforegroundcolor-member)|HTML color code representation of the marker foreground color of a data point (e.g., #FF0000 represents Red).|
||[markerSize](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markersize-member)|Represents marker size of a data point.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerstyle-member)|Represents marker style of a chart data point.|
||[value](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-value-member)|Returns the value of a chart point.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-border-member)|Represents the border format of a chart data point, which includes color, style, and weight information.|
||[fill](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-fill-member)|Represents the fill format of a chart, which includes background formatting information.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[fill](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-fill-member)|Represents the fill format of a chart series, which includes background formatting information.|
||[line](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-line-member)|Represents line formatting.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-add-member(1))|Add a new series to the collection.|
||[count](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-count-member)|Returns the number of series in the collection.|
||[getCount()](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getcount-member(1))|Returns the number of series in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getitemat-member(1))|Retrieves a series based on its position in the collection.|
||[items](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-items-member)|Gets the loaded child items in this collection.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-getsubstring-member(1))|Get the substring of a chart title.|
||[height](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-height-member)|Returns the height, in points, of the chart title.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-horizontalalignment-member)|Specifies the horizontal alignment for chart title.|
||[left](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-left-member)|Specifies the distance, in points, from the left edge of chart title to the left edge of chart area.|
||[overlay](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-overlay-member)|Specifies if the chart title will overlay the chart.|
||[position](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-position-member)|Represents the position of chart title.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-setformula-member(1))|Sets a string value that represents the formula of chart title using A1-style notation.|
||[showShadow](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-showshadow-member)|Represents a boolean value that determines if the chart title has a shadow.|
||[text](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-text-member)|Specifies the chart's title text.|
||[textOrientation](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-textorientation-member)|Specifies the angle to which the text is oriented for the chart title.|
||[top](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-top-member)|Specifies the distance, in points, from the top edge of chart title to the top of chart area.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-verticalalignment-member)|Specifies the vertical alignment of chart title.|
||[visible](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-visible-member)|Specifies if the chart title is visibile.|
||[width](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-width-member)|Specifies the width, in points, of the chart title.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-border-member)|Represents the border format of chart title, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-font-member)|Represents the font attributes (such as font name, font size, and color) for an object.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[format](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-format-member)|Represents the formatting of a chart trendline.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel.ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-add-member(1))|Adds a new trendline to trendline collection.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getcount-member(1))|Returns the number of trendlines in the collection.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getitem-member(1))|Gets a trendline object by index, which is the insertion order in the items array.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-items-member)|Gets the loaded child items in this collection.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#excel-excel-charttrendlineformat-line-member)|Represents chart line formatting.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-delete-member(1))|Deletes the custom property.|
||[key](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-key-member)|The key of the custom property.|
||[type](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-type-member)|The type of the value used for the custom property.|
||[value](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-value-member)|The value of the custom property.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-add-member(1))|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-deleteall-member(1))|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getcount-member(1))|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitem-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitemornullobject-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-items-member)|Gets the loaded child items in this collection.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#excel-excel-dataconnectioncollection-refreshall-member(1))|Refreshes all the data connections in the collection.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-author-member)|The author of the workbook.|
||[category](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-category-member)|The category of the workbook.|
||[comments](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-comments-member)|The comments of the workbook.|
||[company](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-company-member)|The company of the workbook.|
||[creationDate](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-creationdate-member)|Gets the creation date of the workbook.|
||[custom](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-custom-member)|Gets the collection of custom properties of the workbook.|
||[keywords](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-keywords-member)|The keywords of the workbook.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-lastauthor-member)|Gets the last author of the workbook.|
||[manager](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-manager-member)|The manager of the workbook.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-revisionnumber-member)|Gets the revision number of the workbook.|
||[subject](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-subject-member)|The subject of the workbook.|
||[title](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-title-member)|The title of the workbook.|
|[FunctionResult](/javascript/api/excel/excel.functionresult)|[error](/javascript/api/excel/excel.functionresult#excel-excel-functionresult-error-member)|Error value (such as "#DIV/0") representing the error.|
||[value](/javascript/api/excel/excel.functionresult#excel-excel-functionresult-value-member)|The value of function evaluation.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[arrayValues](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-arrayvalues-member)|Returns an object containing values and types of the named item.|
||[delete()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-delete-member(1))|Deletes the given name.|
||[formula](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-formula-member)|The formula of the named item.|
||[getRange()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrange-member(1))|Returns the range object that is associated with the name.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrangeornullobject-member(1))|Returns the range object that is associated with the name.|
||[name](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-name-member)|The name of the object.|
||[scope](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-scope-member)|Specifies if the name is scoped to the workbook or to a specific worksheet.|
||[type](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-type-member)|Specifies the type of the value returned by the name's formula.|
||[value](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-value-member)|Represents the value computed by the name's formula.|
||[visible](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-visible-member)|Specifies if the object is visible.|
||[worksheet](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheet-member)|Returns the worksheet on which the named item is scoped to.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheetornullobject-member)|Returns the worksheet to which the named item is scoped.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-types-member)|Represents the types for each item in the named item array|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-values-member)|Represents the values of each item in the named item array.|
|[Range](/javascript/api/excel/excel.range)|[documentReference](/javascript/api/excel/excel.range#excel-excel-range-documentreference-member)|Represents the document reference target for the hyperlink.|
||[screenTip](/javascript/api/excel/excel.range#excel-excel-range-screentip-member)|Represents the string displayed when hovering over the hyperlink.|
||[textToDisplay](/javascript/api/excel/excel.range#excel-excel-range-texttodisplay-member)|Represents the string that is displayed in the top left most cell in the range.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[cellAddresses](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-celladdresses-member)|Represents the cell addresses of the `RangeView`.|
||[columnCount](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-columncount-member)|The number of visible columns.|
||[formulas](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulas-member)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulaslocal-member)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulasr1c1-member)|Represents the formula in R1C1-style notation.|
||[getRange()](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-getrange-member(1))|Gets the parent range associated with the current `RangeView`.|
||[index](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-index-member)|Returns a value that represents the index of the `RangeView`.|
||[numberFormat](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-numberformat-member)|Represents Excel's number format code for the given cell.|
||[rowCount](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-rowcount-member)|The number of visible rows.|
||[rows](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-rows-member)|Represents a collection of range views associated with the range.|
||[text](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-text-member)|Text values of the specified range.|
||[valueTypes](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuetypes-member)|Represents the type of data of each cell.|
||[values](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-values-member)|Represents the raw values of the specified range view.|
|[Style](/javascript/api/excel/excel.style)|[borders](/javascript/api/excel/excel.style#excel-excel-style-borders-member)|A collection of four border objects that represent the style of the four borders.|
||[fill](/javascript/api/excel/excel.style#excel-excel-style-fill-member)|The fill of the style.|
||[font](/javascript/api/excel/excel.style#excel-excel-style-font-member)|A `Font` object that represents the font of the style.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-add-member(1))|Adds a new style to the collection.|
||[items](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-items-member)|Gets the loaded child items in this collection.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member)|Occurs when data in cells changes on a specific table.|
||[onSelectionChanged](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onselectionchanged-member)|Occurs when the selection changes on a specific table.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-add-member(1))|Adds a new column to the table.|
||[count](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-count-member)|Returns the number of columns in the table.|
||[getCount()](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getcount-member(1))|Gets the number of columns in the table.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitem-member(1))|Gets a column object by name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemat-member(1))|Gets a column based on its position in the collection.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemornullobject-member(1))|Gets a column object by name or ID.|
||[items](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-items-member)|Gets the loaded child items in this collection.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-address-member)|Gets the range address that represents the selected area of the table on a specific worksheet.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-isinsidetable-member)|Specifies if the selection is inside a table.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-tableid-member)|Gets the ID of the table in which the selection changed.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the selection changed.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-freezeat-member(1))|Sets the frozen cells in the active worksheet view.|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-freezecolumns-member(1))|Freeze the first column or columns of the worksheet in place.|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-freezerows-member(1))|Freeze the top row or rows of the worksheet in place.|
||[getLocation()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getlocation-member(1))|Gets a range that describes the frozen cells in the active worksheet view.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getlocationornullobject-member(1))|Gets a range that describes the frozen cells in the active worksheet view.|
||[unfreeze()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-unfreeze-member(1))|Removes all frozen panes in the worksheet.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-worksheetid-member)|Gets the ID of the worksheet that is added to the workbook.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-address-member)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changetype-member)|Gets the change type that represents how the changed event is triggered.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changetype-member)|Gets the change type that represents how the changed event is triggered.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-source-member)|Gets the source of the event.|
||[tableId](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-tableid-member)|Gets the ID of the table in which the data changed.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the data changed.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-worksheetid-member)|Gets the ID of the worksheet that is activated.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet that is deactivated.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-worksheetid-member)|Gets the ID of the worksheet that is deleted from the workbook.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-address-member)|Gets the range address that represents the selected area of a specific worksheet.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the selection changed.|
