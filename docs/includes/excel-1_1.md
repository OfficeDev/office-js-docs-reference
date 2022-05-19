| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|||
|[Binding](/javascript/api/excel/excel.binding)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.binding#excel-excel-binding-add-member(1))|Creates a new table.|
||[count](/javascript/api/excel/excel.binding#excel-excel-binding-count-member)|Returns the number of bindings in the collection.|
||[count](/javascript/api/excel/excel.binding#excel-excel-binding-count-member)|Returns the number of tables in the workbook.|
||[id](/javascript/api/excel/excel.binding#excel-excel-binding-id-member)|Represents the binding identifier.|
||[items](/javascript/api/excel/excel.binding#excel-excel-binding-items-member)|Gets the loaded child items in this collection.|
||[items](/javascript/api/excel/excel.binding#excel-excel-binding-items-member)|Gets the loaded child items in this collection.|
||[items](/javascript/api/excel/excel.binding#excel-excel-binding-items-member)|Gets the loaded child items in this collection.|
||[type](/javascript/api/excel/excel.binding#excel-excel-binding-type-member)|Returns the type of the binding.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-categoryaxis-member)|Represents the category axis in a chart.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-seriesaxis-member)|Represents the series axis of a 3-D chart.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-valueaxis-member)|Represents the value axis in an axis.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[format](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-format-member)|Represents the formatting of a chart object, which includes line and font formatting.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorgridlines-member)|Returns an object that represents the major gridlines for the specified axis.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorgridlines-member)|Returns an object that represents the minor gridlines for the specified axis.|
||[title](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-title-member)|Represents the axis title.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-format-member)|Specifies the formatting of the chart axis title.|
||[text](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-text-member)|Specifies the axis title.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|||
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel.ChartType, sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-add-member(1))|Creates a new chart.|
||[axes](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-axes-member)|Represents chart axes.|
||[count](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-count-member)|Returns the number of charts in the worksheet.|
||[count](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-count-member)|Returns the number of series in the collection.|
||[dataLabels](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-datalabels-member)|Represents the data labels on the chart.|
||[fill](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-fill-member)|Represents the fill format of a chart series, which includes background formatting information.|
||[format](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-format-member)|Encapsulates the format properties for the chart area.|
||[items](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-items-member)|Gets the loaded child items in this collection.|
||[items](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-items-member)|Gets the loaded child items in this collection.|
||[legend](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-legend-member)|Represents the legend for the chart.|
||[line](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-line-member)|Represents line formatting.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[format](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-format-member)|Specifies the format of chart data labels, which includes fill and font formatting.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-clear-member(1))|Clears the fill color of a chart element.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-setsolidcolor-member(1))|Sets the fill formatting of a chart element to a uniform color.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-bold-member)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-color-member)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[italic](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-italic-member)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-name-member)|Font name (e.g., "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-size-member)|Size of the font (e.g., 11)|
||[underline](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-underline-member)|Type of underline applied to the font.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-format-member)|Represents the formatting of chart gridlines.|
||[visible](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-visible-member)|Specifies if the axis gridlines are visible.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#excel-excel-chartgridlinesformat-line-member)|Represents chart line formatting.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[format](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-format-member)|Represents the formatting of a chart legend, which includes fill and font formatting.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|||
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[color](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-color-member)|HTML color code representing the color of lines in the chart.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|||
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[count](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-count-member)|Returns the number of chart points in the series.|
||[items](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-items-member)|Gets the loaded child items in this collection.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[format](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-format-member)|Represents the formatting of a chart title, which includes fill and font formatting.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|||
|[NamedItem](/javascript/api/excel/excel.nameditem)|||
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[items](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-items-member)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/excel/excel.range)|||
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-color-member)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-sideindex-member)|Constant value that indicates the specific side of the border.|
||[style](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-style-member)|One of the constants of line style specifying the line style for the border.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[bold](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-bold-member)|Represents the bold status of the font.|
||[color](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-color-member)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[count](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-count-member)|Number of border objects in the collection.|
||[italic](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-italic-member)|Specifies the italic status of the font.|
||[items](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-items-member)|Gets the loaded child items in this collection.|
||[name](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-name-member)|Font name (e.g., "Calibri").|
||[size](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-size-member)|Font size.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[color](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-color-member)|HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[borders](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-borders-member)|Collection of border objects that apply to the overall range.|
||[fill](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-fill-member)|Returns the fill object defined on the overall range.|
||[font](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-font-member)|Returns the font object defined on the overall range.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, alwaysInsert?: boolean)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))|Adds one or more rows to the table.|
||[count](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-count-member)|Returns the number of rows in the table.|
||[delete()](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-delete-member(1))|Deletes the row from the table.|
||[getRange()](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getrange-member(1))|Returns the range object associated with the entire row.|
||[index](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-index-member)|Returns the index number of the row within the rows collection of the table.|
||[items](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-items-member)|Gets the loaded child items in this collection.|
||[values](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-values-member)|Represents the raw values of the specified range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[application](/javascript/api/excel/excel.workbook#excel-excel-workbook-application-member)|Represents the Excel application instance that contains this workbook.|
||[bindings](/javascript/api/excel/excel.workbook#excel-excel-workbook-bindings-member)|Represents a collection of bindings that are part of the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|||
