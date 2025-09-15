| Class | Fields | Description |
|:---|:---|:---|
|[Application](/.application)|[calculate(calculationType: Excel.CalculationType)](/.application#excel-javascript/api/excel/-application-calculate-member(1))|Recalculate all currently opened workbooks in Excel.|
||[calculationMode](/.application#excel-javascript/api/excel/-application-calculationmode-member)|Returns the calculation mode used in the workbook, as defined by the constants in `Excel.CalculationMode`.|
|[Binding](/.binding)|[getRange()](/.binding#excel-javascript/api/excel/-binding-getrange-member(1))|Returns the range represented by the binding.|
||[getTable()](/.binding#excel-javascript/api/excel/-binding-gettable-member(1))|Returns the table represented by the binding.|
||[getText()](/.binding#excel-javascript/api/excel/-binding-gettext-member(1))|Returns the text represented by the binding.|
||[id](/.binding#excel-javascript/api/excel/-binding-id-member)|Represents the binding identifier.|
||[type](/.binding#excel-javascript/api/excel/-binding-type-member)|Returns the type of the binding.|
|[BindingCollection](/.bindingcollection)|[count](/.bindingcollection#excel-javascript/api/excel/-bindingcollection-count-member)|Returns the number of bindings in the collection.|
||[getItem(id: string)](/.bindingcollection#excel-javascript/api/excel/-bindingcollection-getitem-member(1))|Gets a binding object by ID.|
||[getItemAt(index: number)](/.bindingcollection#excel-javascript/api/excel/-bindingcollection-getitemat-member(1))|Gets a binding object based on its position in the items array.|
||[items](/.bindingcollection#excel-javascript/api/excel/-bindingcollection-items-member)|Gets the loaded child items in this collection.|
|[Chart](/.chart)|[axes](/.chart#excel-javascript/api/excel/-chart-axes-member)|Represents chart axes.|
||[dataLabels](/.chart#excel-javascript/api/excel/-chart-datalabels-member)|Represents the data labels on the chart.|
||[delete()](/.chart#excel-javascript/api/excel/-chart-delete-member(1))|Deletes the chart object.|
||[format](/.chart#excel-javascript/api/excel/-chart-format-member)|Encapsulates the format properties for the chart area.|
||[height](/.chart#excel-javascript/api/excel/-chart-height-member)|Specifies the height, in points, of the chart object.|
||[left](/.chart#excel-javascript/api/excel/-chart-left-member)|The distance, in points, from the left side of the chart to the worksheet origin.|
||[legend](/.chart#excel-javascript/api/excel/-chart-legend-member)|Represents the legend for the chart.|
||[name](/.chart#excel-javascript/api/excel/-chart-name-member)|Specifies the name of a chart object.|
||[series](/.chart#excel-javascript/api/excel/-chart-series-member)|Represents either a single series or collection of series in the chart.|
||[setData(sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/.chart#excel-javascript/api/excel/-chart-setdata-member(1))|Resets the source data for the chart.|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/.chart#excel-javascript/api/excel/-chart-setposition-member(1))|Positions the chart relative to cells on the worksheet.|
||[title](/.chart#excel-javascript/api/excel/-chart-title-member)|Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.|
||[top](/.chart#excel-javascript/api/excel/-chart-top-member)|Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
||[width](/.chart#excel-javascript/api/excel/-chart-width-member)|Specifies the width, in points, of the chart object.|
|[ChartAreaFormat](/.chartareaformat)|[fill](/.chartareaformat#excel-javascript/api/excel/-chartareaformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/.chartareaformat#excel-javascript/api/excel/-chartareaformat-font-member)|Represents the font attributes (font name, font size, color, etc.) for the current object.|
|[ChartAxes](/.chartaxes)|[categoryAxis](/.chartaxes#excel-javascript/api/excel/-chartaxes-categoryaxis-member)|Represents the category axis in a chart.|
||[seriesAxis](/.chartaxes#excel-javascript/api/excel/-chartaxes-seriesaxis-member)|Represents the series axis of a 3-D chart.|
||[valueAxis](/.chartaxes#excel-javascript/api/excel/-chartaxes-valueaxis-member)|Represents the value axis in an axis.|
|[ChartAxis](/.chartaxis)|[format](/.chartaxis#excel-javascript/api/excel/-chartaxis-format-member)|Represents the formatting of a chart object, which includes line and font formatting.|
||[majorGridlines](/.chartaxis#excel-javascript/api/excel/-chartaxis-majorgridlines-member)|Returns an object that represents the major gridlines for the specified axis.|
||[majorUnit](/.chartaxis#excel-javascript/api/excel/-chartaxis-majorunit-member)|Represents the interval between two major tick marks.|
||[maximum](/.chartaxis#excel-javascript/api/excel/-chartaxis-maximum-member)|Represents the maximum value on the value axis.|
||[minimum](/.chartaxis#excel-javascript/api/excel/-chartaxis-minimum-member)|Represents the minimum value on the value axis.|
||[minorGridlines](/.chartaxis#excel-javascript/api/excel/-chartaxis-minorgridlines-member)|Returns an object that represents the minor gridlines for the specified axis.|
||[minorUnit](/.chartaxis#excel-javascript/api/excel/-chartaxis-minorunit-member)|Represents the interval between two minor tick marks.|
||[title](/.chartaxis#excel-javascript/api/excel/-chartaxis-title-member)|Represents the axis title.|
|[ChartAxisFormat](/.chartaxisformat)|[font](/.chartaxisformat#excel-javascript/api/excel/-chartaxisformat-font-member)|Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.|
||[line](/.chartaxisformat#excel-javascript/api/excel/-chartaxisformat-line-member)|Specifies chart line formatting.|
|[ChartAxisTitle](/.chartaxistitle)|[format](/.chartaxistitle#excel-javascript/api/excel/-chartaxistitle-format-member)|Specifies the formatting of the chart axis title.|
||[text](/.chartaxistitle#excel-javascript/api/excel/-chartaxistitle-text-member)|Specifies the axis title.|
||[visible](/.chartaxistitle#excel-javascript/api/excel/-chartaxistitle-visible-member)|Specifies if the axis title is visible.|
|[ChartAxisTitleFormat](/.chartaxistitleformat)|[font](/.chartaxistitleformat#excel-javascript/api/excel/-chartaxistitleformat-font-member)|Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.|
|[ChartCollection](/.chartcollection)|[add(type: Excel.ChartType, sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/.chartcollection#excel-javascript/api/excel/-chartcollection-add-member(1))|Creates a new chart.|
||[count](/.chartcollection#excel-javascript/api/excel/-chartcollection-count-member)|Returns the number of charts in the worksheet.|
||[getItem(name: string)](/.chartcollection#excel-javascript/api/excel/-chartcollection-getitem-member(1))|Gets a chart using its name.|
||[getItemAt(index: number)](/.chartcollection#excel-javascript/api/excel/-chartcollection-getitemat-member(1))|Gets a chart based on its position in the collection.|
||[items](/.chartcollection#excel-javascript/api/excel/-chartcollection-items-member)|Gets the loaded child items in this collection.|
|[ChartDataLabelFormat](/.chartdatalabelformat)|[fill](/.chartdatalabelformat#excel-javascript/api/excel/-chartdatalabelformat-fill-member)|Represents the fill format of the current chart data label.|
||[font](/.chartdatalabelformat#excel-javascript/api/excel/-chartdatalabelformat-font-member)|Represents the font attributes (such as font name, font size, and color) for a chart data label.|
|[ChartDataLabels](/.chartdatalabels)|[format](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-format-member)|Specifies the format of chart data labels, which includes fill and font formatting.|
||[position](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-position-member)|Value that represents the position of the data label.|
||[separator](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-separator-member)|String representing the separator used for the data labels on a chart.|
||[showBubbleSize](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-showbubblesize-member)|Specifies if the data label bubble size is visible.|
||[showCategoryName](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-showcategoryname-member)|Specifies if the data label category name is visible.|
||[showLegendKey](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-showlegendkey-member)|Specifies if the data label legend key is visible.|
||[showPercentage](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-showpercentage-member)|Specifies if the data label percentage is visible.|
||[showSeriesName](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-showseriesname-member)|Specifies if the data label series name is visible.|
||[showValue](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-showvalue-member)|Specifies if the data label value is visible.|
|[ChartFill](/.chartfill)|[clear()](/.chartfill#excel-javascript/api/excel/-chartfill-clear-member(1))|Clears the fill color of a chart element.|
||[setSolidColor(color: string)](/.chartfill#excel-javascript/api/excel/-chartfill-setsolidcolor-member(1))|Sets the fill formatting of a chart element to a uniform color.|
|[ChartFont](/.chartfont)|[bold](/.chartfont#excel-javascript/api/excel/-chartfont-bold-member)|Represents the bold status of font.|
||[color](/.chartfont#excel-javascript/api/excel/-chartfont-color-member)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[italic](/.chartfont#excel-javascript/api/excel/-chartfont-italic-member)|Represents the italic status of the font.|
||[name](/.chartfont#excel-javascript/api/excel/-chartfont-name-member)|Font name (e.g., "Calibri")|
||[size](/.chartfont#excel-javascript/api/excel/-chartfont-size-member)|Size of the font (e.g., 11)|
||[underline](/.chartfont#excel-javascript/api/excel/-chartfont-underline-member)|Type of underline applied to the font.|
|[ChartGridlines](/.chartgridlines)|[format](/.chartgridlines#excel-javascript/api/excel/-chartgridlines-format-member)|Represents the formatting of chart gridlines.|
||[visible](/.chartgridlines#excel-javascript/api/excel/-chartgridlines-visible-member)|Specifies if the axis gridlines are visible.|
|[ChartGridlinesFormat](/.chartgridlinesformat)|[line](/.chartgridlinesformat#excel-javascript/api/excel/-chartgridlinesformat-line-member)|Represents chart line formatting.|
|[ChartLegend](/.chartlegend)|[format](/.chartlegend#excel-javascript/api/excel/-chartlegend-format-member)|Represents the formatting of a chart legend, which includes fill and font formatting.|
||[overlay](/.chartlegend#excel-javascript/api/excel/-chartlegend-overlay-member)|Specifies if the chart legend should overlap with the main body of the chart.|
||[position](/.chartlegend#excel-javascript/api/excel/-chartlegend-position-member)|Specifies the position of the legend on the chart.|
||[visible](/.chartlegend#excel-javascript/api/excel/-chartlegend-visible-member)|Specifies if the chart legend is visible.|
|[ChartLegendFormat](/.chartlegendformat)|[fill](/.chartlegendformat#excel-javascript/api/excel/-chartlegendformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/.chartlegendformat#excel-javascript/api/excel/-chartlegendformat-font-member)|Represents the font attributes such as font name, font size, and color of a chart legend.|
|[ChartLineFormat](/.chartlineformat)|[clear()](/.chartlineformat#excel-javascript/api/excel/-chartlineformat-clear-member(1))|Clears the line format of a chart element.|
||[color](/.chartlineformat#excel-javascript/api/excel/-chartlineformat-color-member)|HTML color code representing the color of lines in the chart.|
|[ChartPoint](/.chartpoint)|[format](/.chartpoint#excel-javascript/api/excel/-chartpoint-format-member)|Encapsulates the format properties chart point.|
||[value](/.chartpoint#excel-javascript/api/excel/-chartpoint-value-member)|Returns the value of a chart point.|
|[ChartPointFormat](/.chartpointformat)|[fill](/.chartpointformat#excel-javascript/api/excel/-chartpointformat-fill-member)|Represents the fill format of a chart, which includes background formatting information.|
|[ChartPointsCollection](/.chartpointscollection)|[count](/.chartpointscollection#excel-javascript/api/excel/-chartpointscollection-count-member)|Returns the number of chart points in the series.|
||[getItemAt(index: number)](/.chartpointscollection#excel-javascript/api/excel/-chartpointscollection-getitemat-member(1))|Retrieve a point based on its position within the series.|
||[items](/.chartpointscollection#excel-javascript/api/excel/-chartpointscollection-items-member)|Gets the loaded child items in this collection.|
|[ChartSeries](/.chartseries)|[format](/.chartseries#excel-javascript/api/excel/-chartseries-format-member)|Represents the formatting of a chart series, which includes fill and line formatting.|
||[name](/.chartseries#excel-javascript/api/excel/-chartseries-name-member)|Specifies the name of a series in a chart.|
||[points](/.chartseries#excel-javascript/api/excel/-chartseries-points-member)|Returns a collection of all points in the series.|
|[ChartSeriesCollection](/.chartseriescollection)|[count](/.chartseriescollection#excel-javascript/api/excel/-chartseriescollection-count-member)|Returns the number of series in the collection.|
||[getItemAt(index: number)](/.chartseriescollection#excel-javascript/api/excel/-chartseriescollection-getitemat-member(1))|Retrieves a series based on its position in the collection.|
||[items](/.chartseriescollection#excel-javascript/api/excel/-chartseriescollection-items-member)|Gets the loaded child items in this collection.|
|[ChartSeriesFormat](/.chartseriesformat)|[fill](/.chartseriesformat#excel-javascript/api/excel/-chartseriesformat-fill-member)|Represents the fill format of a chart series, which includes background formatting information.|
||[line](/.chartseriesformat#excel-javascript/api/excel/-chartseriesformat-line-member)|Represents line formatting.|
|[ChartTitle](/.charttitle)|[format](/.charttitle#excel-javascript/api/excel/-charttitle-format-member)|Represents the formatting of a chart title, which includes fill and font formatting.|
||[overlay](/.charttitle#excel-javascript/api/excel/-charttitle-overlay-member)|Specifies if the chart title will overlay the chart.|
||[text](/.charttitle#excel-javascript/api/excel/-charttitle-text-member)|Specifies the chart's title text.|
||[visible](/.charttitle#excel-javascript/api/excel/-charttitle-visible-member)|Specifies if the chart title is visible.|
|[ChartTitleFormat](/.charttitleformat)|[fill](/.charttitleformat#excel-javascript/api/excel/-charttitleformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/.charttitleformat#excel-javascript/api/excel/-charttitleformat-font-member)|Represents the font attributes (such as font name, font size, and color) for an object.|
|[NamedItem](/.nameditem)|[getRange()](/.nameditem#excel-javascript/api/excel/-nameditem-getrange-member(1))|Returns the range object that is associated with the name.|
||[name](/.nameditem#excel-javascript/api/excel/-nameditem-name-member)|The name of the object.|
||[type](/.nameditem#excel-javascript/api/excel/-nameditem-type-member)|Specifies the type of the value returned by the name's formula.|
||[value](/.nameditem#excel-javascript/api/excel/-nameditem-value-member)|Represents the value computed by the name's formula.|
||[visible](/.nameditem#excel-javascript/api/excel/-nameditem-visible-member)|Specifies if the object is visible.|
|[NamedItemCollection](/.nameditemcollection)|[getItem(name: string)](/.nameditemcollection#excel-javascript/api/excel/-nameditemcollection-getitem-member(1))|Gets a `NamedItem` object using its name.|
||[items](/.nameditemcollection#excel-javascript/api/excel/-nameditemcollection-items-member)|Gets the loaded child items in this collection.|
|[Range](/.range)|[address](/.range#excel-javascript/api/excel/-range-address-member)|Specifies the range reference in A1-style.|
||[addressLocal](/.range#excel-javascript/api/excel/-range-addresslocal-member)|Represents the range reference for the specified range in the language of the user.|
||[cellCount](/.range#excel-javascript/api/excel/-range-cellcount-member)|Specifies the number of cells in the range.|
||[clear(applyTo?: Excel.ClearApplyTo)](/.range#excel-javascript/api/excel/-range-clear-member(1))|Clear range values and formatting, such as fill and border.|
||[columnCount](/.range#excel-javascript/api/excel/-range-columncount-member)|Specifies the total number of columns in the range.|
||[columnIndex](/.range#excel-javascript/api/excel/-range-columnindex-member)|Specifies the column number of the first cell in the range.|
||[delete(shift: Excel.DeleteShiftDirection)](/.range#excel-javascript/api/excel/-range-delete-member(1))|Deletes the cells associated with the range.|
||[format](/.range#excel-javascript/api/excel/-range-format-member)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.|
||[formulas](/.range#excel-javascript/api/excel/-range-formulas-member)|Represents the formula in A1-style notation.|
||[formulasLocal](/.range#excel-javascript/api/excel/-range-formulaslocal-member)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.|
||[getBoundingRect(anotherRange: Range \| string)](/.range#excel-javascript/api/excel/-range-getboundingrect-member(1))|Gets the smallest range object that encompasses the given ranges.|
||[getCell(row: number, column: number)](/.range#excel-javascript/api/excel/-range-getcell-member(1))|Gets the range object containing the single cell based on row and column numbers.|
||[getColumn(column: number)](/.range#excel-javascript/api/excel/-range-getcolumn-member(1))|Gets a column contained in the range.|
||[getEntireColumn()](/.range#excel-javascript/api/excel/-range-getentirecolumn-member(1))|Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").|
||[getEntireRow()](/.range#excel-javascript/api/excel/-range-getentirerow-member(1))|Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").|
||[getIntersection(anotherRange: Range \| string)](/.range#excel-javascript/api/excel/-range-getintersection-member(1))|Gets the range object that represents the rectangular intersection of the given ranges.|
||[getLastCell()](/.range#excel-javascript/api/excel/-range-getlastcell-member(1))|Gets the last cell within the range.|
||[getLastColumn()](/.range#excel-javascript/api/excel/-range-getlastcolumn-member(1))|Gets the last column within the range.|
||[getLastRow()](/.range#excel-javascript/api/excel/-range-getlastrow-member(1))|Gets the last row within the range.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/.range#excel-javascript/api/excel/-range-getoffsetrange-member(1))|Gets an object which represents a range that's offset from the specified range.|
||[getRow(row: number)](/.range#excel-javascript/api/excel/-range-getrow-member(1))|Gets a row contained in the range.|
||[insert(shift: Excel.InsertShiftDirection)](/.range#excel-javascript/api/excel/-range-insert-member(1))|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space.|
||[numberFormat](/.range#excel-javascript/api/excel/-range-numberformat-member)|Represents Excel's number format code for the given range.|
||[rowCount](/.range#excel-javascript/api/excel/-range-rowcount-member)|Returns the total number of rows in the range.|
||[rowIndex](/.range#excel-javascript/api/excel/-range-rowindex-member)|Returns the row number of the first cell in the range.|
||[select()](/.range#excel-javascript/api/excel/-range-select-member(1))|Selects the specified range in the Excel UI.|
||[text](/.range#excel-javascript/api/excel/-range-text-member)|Text values of the specified range.|
||[valueTypes](/.range#excel-javascript/api/excel/-range-valuetypes-member)|Specifies the type of data in each cell.|
||[values](/.range#excel-javascript/api/excel/-range-values-member)|Represents the raw values of the specified range.|
||[worksheet](/.range#excel-javascript/api/excel/-range-worksheet-member)|The worksheet containing the current range.|
|[RangeBorder](/.rangeborder)|[color](/.rangeborder#excel-javascript/api/excel/-rangeborder-color-member)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").|
||[sideIndex](/.rangeborder#excel-javascript/api/excel/-rangeborder-sideindex-member)|Constant value that indicates the specific side of the border.|
||[style](/.rangeborder#excel-javascript/api/excel/-rangeborder-style-member)|One of the constants of line style specifying the line style for the border.|
||[weight](/.rangeborder#excel-javascript/api/excel/-rangeborder-weight-member)|Specifies the weight of the border around a range.|
|[RangeBorderCollection](/.rangebordercollection)|[count](/.rangebordercollection#excel-javascript/api/excel/-rangebordercollection-count-member)|Number of border objects in the collection.|
||[getItem(index: Excel.BorderIndex)](/.rangebordercollection#excel-javascript/api/excel/-rangebordercollection-getitem-member(1))|Gets a border object using its name.|
||[getItemAt(index: number)](/.rangebordercollection#excel-javascript/api/excel/-rangebordercollection-getitemat-member(1))|Gets a border object using its index.|
||[items](/.rangebordercollection#excel-javascript/api/excel/-rangebordercollection-items-member)|Gets the loaded child items in this collection.|
|[RangeFill](/.rangefill)|[clear()](/.rangefill#excel-javascript/api/excel/-rangefill-clear-member(1))|Resets the range background.|
||[color](/.rangefill#excel-javascript/api/excel/-rangefill-color-member)|HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")|
|[RangeFont](/.rangefont)|[bold](/.rangefont#excel-javascript/api/excel/-rangefont-bold-member)|Represents the bold status of the font.|
||[color](/.rangefont#excel-javascript/api/excel/-rangefont-color-member)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[italic](/.rangefont#excel-javascript/api/excel/-rangefont-italic-member)|Specifies the italic status of the font.|
||[name](/.rangefont#excel-javascript/api/excel/-rangefont-name-member)|Font name (e.g., "Calibri").|
||[size](/.rangefont#excel-javascript/api/excel/-rangefont-size-member)|Font size.|
||[underline](/.rangefont#excel-javascript/api/excel/-rangefont-underline-member)|Type of underline applied to the font.|
|[RangeFormat](/.rangeformat)|[borders](/.rangeformat#excel-javascript/api/excel/-rangeformat-borders-member)|Collection of border objects that apply to the overall range.|
||[fill](/.rangeformat#excel-javascript/api/excel/-rangeformat-fill-member)|Returns the fill object defined on the overall range.|
||[font](/.rangeformat#excel-javascript/api/excel/-rangeformat-font-member)|Returns the font object defined on the overall range.|
||[horizontalAlignment](/.rangeformat#excel-javascript/api/excel/-rangeformat-horizontalalignment-member)|Represents the horizontal alignment for the specified object.|
||[verticalAlignment](/.rangeformat#excel-javascript/api/excel/-rangeformat-verticalalignment-member)|Represents the vertical alignment for the specified object.|
||[wrapText](/.rangeformat#excel-javascript/api/excel/-rangeformat-wraptext-member)|Specifies if Excel wraps the text in the object.|
|[Table](/.table)|[columns](/.table#excel-javascript/api/excel/-table-columns-member)|Represents a collection of all the columns in the table.|
||[delete()](/.table#excel-javascript/api/excel/-table-delete-member(1))|Deletes the table.|
||[getDataBodyRange()](/.table#excel-javascript/api/excel/-table-getdatabodyrange-member(1))|Gets the range object associated with the data body of the table.|
||[getHeaderRowRange()](/.table#excel-javascript/api/excel/-table-getheaderrowrange-member(1))|Gets the range object associated with the header row of the table.|
||[getRange()](/.table#excel-javascript/api/excel/-table-getrange-member(1))|Gets the range object associated with the entire table.|
||[getTotalRowRange()](/.table#excel-javascript/api/excel/-table-gettotalrowrange-member(1))|Gets the range object associated with the totals row of the table.|
||[id](/.table#excel-javascript/api/excel/-table-id-member)|Returns a value that uniquely identifies the table in a given workbook.|
||[name](/.table#excel-javascript/api/excel/-table-name-member)|Name of the table.|
||[rows](/.table#excel-javascript/api/excel/-table-rows-member)|Represents a collection of all the rows in the table.|
||[showHeaders](/.table#excel-javascript/api/excel/-table-showheaders-member)|Specifies if the header row is visible.|
||[showTotals](/.table#excel-javascript/api/excel/-table-showtotals-member)|Specifies if the total row is visible.|
||[style](/.table#excel-javascript/api/excel/-table-style-member)|Constant value that represents the table style.|
|[TableCollection](/.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/.tablecollection#excel-javascript/api/excel/-tablecollection-add-member(1))|Creates a new table.|
||[count](/.tablecollection#excel-javascript/api/excel/-tablecollection-count-member)|Returns the number of tables in the workbook.|
||[getItem(key: string)](/.tablecollection#excel-javascript/api/excel/-tablecollection-getitem-member(1))|Gets a table by name or ID.|
||[getItemAt(index: number)](/.tablecollection#excel-javascript/api/excel/-tablecollection-getitemat-member(1))|Gets a table based on its position in the collection.|
||[items](/.tablecollection#excel-javascript/api/excel/-tablecollection-items-member)|Gets the loaded child items in this collection.|
|[TableColumn](/.tablecolumn)|[delete()](/.tablecolumn#excel-javascript/api/excel/-tablecolumn-delete-member(1))|Deletes the column from the table.|
||[getDataBodyRange()](/.tablecolumn#excel-javascript/api/excel/-tablecolumn-getdatabodyrange-member(1))|Gets the range object associated with the data body of the column.|
||[getHeaderRowRange()](/.tablecolumn#excel-javascript/api/excel/-tablecolumn-getheaderrowrange-member(1))|Gets the range object associated with the header row of the column.|
||[getRange()](/.tablecolumn#excel-javascript/api/excel/-tablecolumn-getrange-member(1))|Gets the range object associated with the entire column.|
||[getTotalRowRange()](/.tablecolumn#excel-javascript/api/excel/-tablecolumn-gettotalrowrange-member(1))|Gets the range object associated with the totals row of the column.|
||[id](/.tablecolumn#excel-javascript/api/excel/-tablecolumn-id-member)|Returns a unique key that identifies the column within the table.|
||[index](/.tablecolumn#excel-javascript/api/excel/-tablecolumn-index-member)|Returns the index number of the column within the columns collection of the table.|
||[name](/.tablecolumn#excel-javascript/api/excel/-tablecolumn-name-member)|Specifies the name of the table column.|
||[values](/.tablecolumn#excel-javascript/api/excel/-tablecolumn-values-member)|Represents the raw values of the specified range.|
|[TableColumnCollection](/.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, name?: string)](/.tablecolumncollection#excel-javascript/api/excel/-tablecolumncollection-add-member(1))|Adds a new column to the table.|
||[count](/.tablecolumncollection#excel-javascript/api/excel/-tablecolumncollection-count-member)|Returns the number of columns in the table.|
||[getItem(key: number \| string)](/.tablecolumncollection#excel-javascript/api/excel/-tablecolumncollection-getitem-member(1))|Gets a column object by name or ID.|
||[getItemAt(index: number)](/.tablecolumncollection#excel-javascript/api/excel/-tablecolumncollection-getitemat-member(1))|Gets a column based on its position in the collection.|
||[items](/.tablecolumncollection#excel-javascript/api/excel/-tablecolumncollection-items-member)|Gets the loaded child items in this collection.|
|[TableRow](/.tablerow)|[delete()](/.tablerow#excel-javascript/api/excel/-tablerow-delete-member(1))|Deletes the row from the table.|
||[getRange()](/.tablerow#excel-javascript/api/excel/-tablerow-getrange-member(1))|Returns the range object associated with the entire row.|
||[index](/.tablerow#excel-javascript/api/excel/-tablerow-index-member)|Returns the index number of the row within the rows collection of the table.|
||[values](/.tablerow#excel-javascript/api/excel/-tablerow-values-member)|Represents the raw values of the specified range.|
|[TableRowCollection](/.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, alwaysInsert?: boolean)](/.tablerowcollection#excel-javascript/api/excel/-tablerowcollection-add-member(1))|Adds one or more rows to the table.|
||[count](/.tablerowcollection#excel-javascript/api/excel/-tablerowcollection-count-member)|Returns the number of rows in the table.|
||[getItemAt(index: number)](/.tablerowcollection#excel-javascript/api/excel/-tablerowcollection-getitemat-member(1))|Gets a row based on its position in the collection.|
||[items](/.tablerowcollection#excel-javascript/api/excel/-tablerowcollection-items-member)|Gets the loaded child items in this collection.|
|[Workbook](/.workbook)|[application](/.workbook#excel-javascript/api/excel/-workbook-application-member)|Represents the Excel application instance that contains this workbook.|
||[bindings](/.workbook#excel-javascript/api/excel/-workbook-bindings-member)|Represents a collection of bindings that are part of the workbook.|
||[getSelectedRange()](/.workbook#excel-javascript/api/excel/-workbook-getselectedrange-member(1))|Gets the currently selected single range from the workbook.|
||[names](/.workbook#excel-javascript/api/excel/-workbook-names-member)|Represents a collection of workbook-scoped named items (named ranges and constants).|
||[tables](/.workbook#excel-javascript/api/excel/-workbook-tables-member)|Represents a collection of tables associated with the workbook.|
||[worksheets](/.workbook#excel-javascript/api/excel/-workbook-worksheets-member)|Represents a collection of worksheets associated with the workbook.|
|[Worksheet](/.worksheet)|[activate()](/.worksheet#excel-javascript/api/excel/-worksheet-activate-member(1))|Activate the worksheet in the Excel UI.|
||[charts](/.worksheet#excel-javascript/api/excel/-worksheet-charts-member)|Returns a collection of charts that are part of the worksheet.|
||[delete()](/.worksheet#excel-javascript/api/excel/-worksheet-delete-member(1))|Deletes the worksheet from the workbook.|
||[getCell(row: number, column: number)](/.worksheet#excel-javascript/api/excel/-worksheet-getcell-member(1))|Gets the `Range` object containing the single cell based on row and column numbers.|
||[getRange(address?: string)](/.worksheet#excel-javascript/api/excel/-worksheet-getrange-member(1))|Gets the `Range` object, representing a single rectangular block of cells, specified by the address or name.|
||[id](/.worksheet#excel-javascript/api/excel/-worksheet-id-member)|Returns a value that uniquely identifies the worksheet in a given workbook.|
||[name](/.worksheet#excel-javascript/api/excel/-worksheet-name-member)|The display name of the worksheet.|
||[position](/.worksheet#excel-javascript/api/excel/-worksheet-position-member)|The zero-based position of the worksheet within the workbook.|
||[tables](/.worksheet#excel-javascript/api/excel/-worksheet-tables-member)|Collection of tables that are part of the worksheet.|
||[visibility](/.worksheet#excel-javascript/api/excel/-worksheet-visibility-member)|The visibility of the worksheet.|
|[WorksheetCollection](/.worksheetcollection)|[add(name?: string)](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-add-member(1))|Adds a new worksheet to the workbook.|
||[getActiveWorksheet()](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-getactiveworksheet-member(1))|Gets the currently active worksheet in the workbook.|
||[getItem(key: string)](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-getitem-member(1))|Gets a worksheet object using its name or ID.|
||[items](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-items-member)|Gets the loaded child items in this collection.|
