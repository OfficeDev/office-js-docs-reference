# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Excel add-ins run across multiple versions of Office, including Office 2016 for Windows, Office for iPad, Office for Mac, and Office Online. The following table lists the Excel requirement sets, the Office host applications that support each requirement set, and the build versions or number for those applications.

> [!NOTE]
> Any API that is marked as **Beta** is not ready for end-user production. We make them available for developers to try them out in test and development environments. They are not meant to be used against production/business critical documents.
> 
> For the requirement sets that are marked as **Beta**, use the specified (or later) version of the Office software and use the Beta library on the CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Entries not marked as **Beta** are generally available and you can use Production library on the CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.

|  Requirement set  |  Office 365 for Windows\*  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Beta  | Please [visit our Excel JavaScript API open specification page](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)! |
| ExcelApi1.7  | Version 1801 (Build 9001.2171) or later| 2.9 or later | 16.9 or later | April 2018 | Coming soon |
| ExcelApi1.6  | Version 1704 (Build 8201.2001) or later| 2.2 or later |15.36 or later| April 2017 | Coming soon|
| ExcelApi1.5  | Version 1703 (Build 8067.2070) or later| 2.2 or later |15.36 or later| March 2017 | Coming soon|
| ExcelApi1.4 | Version 1701 (Build 7870.2024) or later| 2.2 or later |15.36 or later| January 2017 | Coming soon|
| ExcelApi1.3  | Version 1608 (Build 7369.2055) or later| 1.27 or later |  15.27 or later| September 2016 | Version 1608 (Build 7601.6800) or later|
| ExcelApi1.2  | Version 1601 (Build 6741.2088) or later | 1.21 or later | 15.22 or later| January 2016 ||
| ExcelApi1.1  | Version 1509 (Build 4266.1001) or later | 1.19 or later | 15.20 or later| January 2016 ||

> [!NOTE]
> The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1 requirement set.

For more information about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server overview](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## What's new in Excel JavaScript API 1.7

The Excel JavaScript API requirement set 1.7 features include APIs for charts, events, data validation, worksheets, ranges, document properties, named items, protection options and styles.

### Customize charts

With the new chart APIs, you can create additional chart types, add a data series to a chart, set the chart title, add an axis title, add display unit, add a trendline with moving average, change a trendline to linear, and more. The following are some examples:

* Chart axis - get, set, format and remove axis unit, label and title in a chart.
* Chart series - add, set, and delete a series in a chart.  Change series markers, plot orders and sizing.
* Chart trendlines - add, get, and format trendlines in a chart.
* Chart legend - format the legend font in a chart.
* Chart point - set chart point color.
* Chart title substring -  get and set title substring for a chart.
* Chart type - option to create more chart types.

### Events

Excel events APIs provide a variety of event handlers that allow your add-in to automatically run a designated function when a specific event occurs. You can design that function to perform whatever actions your scenario requires. For a list of events that are currently available, see [Work with Events using the Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events).

### Customize the appearance of worksheets and ranges

Using the new APIs, you can customize the appearance of worksheets in multiple ways:

* Freeze panes to keep specific rows or columns visible when you scroll in the worksheet. For example, if the first row in your worksheet contains headers, you might freeze that row so that the column headers will remain visible as you scroll down the worksheet.
* Modify the worksheet tab color.
* Add worksheet headings.


You can customize the appearance of ranges in multiple ways:

* Set the cell style for a range to ensure sure that all cells in the range have consistent formatting. A cell style is a defined set of formatting characteristics, such as fonts and font sizes, number formats, cell borders, and cell shading. Use any of Excel's built-in cell styles or create your own custom cell style.
* Set the text orientation for a range.
* Add or modify a hyperlink on a range that links to another location in the workbook or to an external location.

### Manage document properties

Using the document properties APIs, you can access built-in document properties and also create and manage custom document properties to store state of the workbook and drive workflow and business logic.

### Copy worksheets

Using the worksheet copy APIs, you can copy the data and format from one worksheet to a new worksheet within the same workbook and reduce the amount of data transfer needed.

### Handle ranges with ease

Using the various range APIs, you can do things such as get the surrounding region, get a resized range, and more. These APIs should make tasks like range manipulation and addressing much more efficient.

In addition:

* Workbook and worksheet protection options - use these APIs to protect data in a worksheet and the workbook structure.
* Update a named item - use this API to update a named item.
* Get active cell  - use this API to get the active cell of a workbook.

|Object| What is new| Description|Requirement set|
|:----|:----|:----|:----|
|[chart](../../api/excel/excel.chart)|_Property_ > chartType|Represents the type of the chart. Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|1.7|
|[chart](../../api/excel/excel.chart)|_Property_ > id|The unique id of chart. Read-only.|1.7|
|[chart](../../api/excel/excel.chart)|_Property_ > showAllFieldButtons|Represents whether to display all field buttons on a PivotChart.|1.7|
|[chartAreaFormat](../../api/excel/excel.chartareaformat)|_Relationship_ > border|Represents the border format of chart area, which includes color, linestyle and weight. Read-only.|1.7|
|[chartAxes](../../api/excel/excel.chartaxes)|_Method_ > getItem(type: string, group: string)|Returns the specific axis identified by type and group.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > axisBetweenCategories|Represents whether value axis crosses the category axis between categories.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > axisGroup|Represents the group for the specified axis. Read-only. Possible values are: Primary, Secondary.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > categoryType|Returns or sets the category axis type. Possible values are: Automatic, TextAxis, DateAxis.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > crosses|Represents the specified axis where the other axis crosses. Possible values are: Automatic, Maximum, Minimum, Custom.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > crossesAt|Represents the specified axis where the other axis crosses at. Read Only. Set to this property should use SetCrossesAt(double) method. Read-only.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > customDisplayUnit|Represents the custom axis display unit value. Read Only. To set this property, please use the SetCustomDisplayUnit(double) method. Read-only.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > displayUnit|Represents the axis display unit. Possible values are: None, Hundreds, Thousands, TenThousands, HundredThousands, Millions, TenMillions, HundredMillions, Billions, Trillions, Custom.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > height|Represents the height, in points, of the chart axis. Null if the axis's not visible. Read-only.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > left|Represents the distance, in points, from the left edge of the axis to the left of chart area. Null if the axis's not visible. Read-only.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > logBase|Represents the base of the logarithm when using logarithmic scales.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > reversePlotOrder|Represents whether Microsoft Excel plots data points from last to first.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > scaleType|Represents the value axis scale type. Possible values are: Linear, Logarithmic.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > showDisplayUnitLabel|Represents whether the axis display unit label is visible.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > tickLabelSpacing|Represents the number of categories or series between tick-mark labels. Can be a value from 1 through 31999 or an empty string for automatic setting. The returned value is always a number.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > tickMarkSpacing|Represents the number of categories or series between tick marks.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > top|Represents the distance, in points, from the top edge of the axis to the top of chart area. Null if the axis's not visible. Read-only.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > type|Represents the axis type. Read-only. Possible values are: Invalid, Category, Value, Series.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > visible|A boolean value represents the visibility of the axis.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Property_ > width|Represents the width, in points, of the chart axis. Null if the axis's not visible. Read-only.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Relationship_ > baseTimeUnit|Returns or sets the base unit for the specified category axis.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Relationship_ > majorTickMark|Represents the type of major tick mark for the specified axis.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Relationship_ > majorTimeUnitScale|Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Relationship_ > minorTickMark|Represents the type of minor tick mark for the specified axis.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Relationship_ > minorTimeUnitScale|Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Relationship_ > tickLabelPosition|Represents the position of tick-mark labels on the specified axis.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Method_ > setCategoryNames(sourceData: Range)|Sets all the category names for the specified axis.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Method_ > setCrossesAt(value: double)|Set the specified axis where the other axis crosses at.|1.7|
|[chartAxis](../../api/excel/excel.chartaxis)|_Method_ > setCustomDisplayUnit(value: double)|Sets the axis display unit to a custom value.|1.7|
|[chartBorder](../../api/excel/excel.chartborder)|_Property_ > color|HTML color code representing the color of borders in the chart.|1.7|
|[chartBorder](../../api/excel/excel.chartborder)|_Property_ > weight|Represents weight of the border, in points.|1.7|
|[chartBorder](../../api/excel/excel.chartborder)|_Relationship_ > lineStyle|Represents the line style of the border.|1.7|
|[chartDataLabel](../../api/excel/excel.chartdatalabel)|_Property_ > position|DataLabelPosition value that represents the position of the data label. Possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](../../api/excel/excel.chartdatalabel)|_Property_ > separator|String representing the separator used for the data label on a chart.|1.7|
|[chartDataLabel](../../api/excel/excel.chartdatalabel)|_Property_ > showBubbleSize|Boolean value representing if the data label bubble size is visible or not.|1.7|
|[chartDataLabel](../../api/excel/excel.chartdatalabel)|_Property_ > showCategoryName|Boolean value representing if the data label category name is visible or not.|1.7|
|[chartDataLabel](../../api/excel/excel.chartdatalabel)|_Property_ > showLegendKey|Boolean value representing if the data label legend key is visible or not.|1.7|
|[chartDataLabel](../../api/excel/excel.chartdatalabel)|_Property_ > showPercentage|Boolean value representing if the data label percentage is visible or not.|1.7|
|[chartDataLabel](../../api/excel/excel.chartdatalabel)|_Property_ > showSeriesName|Boolean value representing if the data label series name is visible or not.|1.7|
|[chartDataLabel](../../api/excel/excel.chartdatalabel)|_Property_ > showValue|Boolean value representing if the data label value is visible or not.|1.7|
|[chartLegend](../../api/excel/excel.chartlegend)|_Property_ > height|Represents the height of the legend on the chart.|1.7|
|[chartLegend](../../api/excel/excel.chartlegend)|_Property_ > left|Represents the left of a chart legend.|1.7|
|[chartLegend](../../api/excel/excel.chartlegend)|_Property_ > showShadow|Represents if the legend has shadow on the chart.|1.7|
|[chartLegend](../../api/excel/excel.chartlegend)|_Property_ > top|Represents the top of a chart legend.|1.7|
|[chartLegend](../../api/excel/excel.chartlegend)|_Property_ > width|Represents the width of the legend on the chart.|1.7|
|[chartLegend](../../api/excel/excel.chartlegend)|_Relationship_ > legendEntries|Represents a collection of legendEntries in the legend. Read-only.|1.7|
|[chartLegendEntry](../../api/excel/excel.chartlegendentry)|_Property_ > visible|Represents the visible of a chart legend entry.|1.7|
|[chartLegendEntryCollection](../../api/excel/excel.chartlegendentrycollection)|_Property_ > items|A collection of chartLegendEntry objects. Read-only.|1.7|
|[chartLegendEntryCollection](../../api/excel/excel.chartlegendentrycollection)|_Method_ > getCount()|Returns the number of legendEntry in the collection.|1.7|
|[chartLegendEntryCollection](../../api/excel/excel.chartlegendentrycollection)|_Method_ > getItemAt(index: number)|Returns a legendEntry at the given index.|1.7|
|[chartPoint](../../api/excel/excel.chartpoint)|_Property_ > hasDataLabel|Represents whether a data point has datalabel. Not applicable for surface charts.|1.7|
|[chartPoint](../../api/excel/excel.chartpoint)|_Property_ > markerBackgroundColor|HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.|1.7|
|[chartPoint](../../api/excel/excel.chartpoint)|_Property_ > markerForegroundColor|HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.|1.7|
|[chartPoint](../../api/excel/excel.chartpoint)|_Property_ > markerSize|Represents marker size of data point.|1.7|
|[chartPoint](../../api/excel/excel.chartpoint)|_Property_ > markerStyle|Represents marker style of a chart data point. Possible values are: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[chartPoint](../../api/excel/excel.chartpoint)|_Relationship_ > dataLabel|Returns the data label of a chart point. Read-only.|1.7|
|[chartPointFormat](../../api/excel/excel.chartpointformat)|_Relationship_ > border|Represents the border format of a chart data point, which includes color, style and weight information. Read-only.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > chartType|Represents the chart type of a series. Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > doughnutHoleSize|Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > filtered|Boolean value representing if the series is filtered or not. Not applicable for surface charts.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > gapWidth|Represents the gap width of a chart series.  Only valid on bar and column charts, as well as|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > hasDataLabels|Boolean value representing if the series has data labels or not.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > markerBackgroundColor|Represents markers background color of a chart series.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > markerForegroundColor|Represents markers foreground color of a chart series.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > markerSize|Represents marker size of a chart series.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > markerStyle|Represents marker style of a chart series. Possible values are: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > plotOrder|Represents the plot order of a chart series within the chart group.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > showShadow|Boolean value representing if the series has shadow or not.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Property_ > smooth|Boolean value representing if the series is smooth or not. Only for line and scatter charts.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Relationship_ > dataLabels|Represents a collection of all dataLabels in the series. Read-only.|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](../../api/excel/excel.chartseries)|_Relationship_ > trendlines|Represents a collection of trendlines in the series. Read-only.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Method_ > delete()|Deletes the chart series.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Method_ > setBubbleSizes(sourceData: Range)|Set bubble sizes for a chart series. Only works for bubble charts.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Method_ > setValues(sourceData: Range)|Set values for a chart series. For scatter chart, it means Y axis values.|1.7|
|[chartSeries](../../api/excel/excel.chartseries)|_Method_ > setXAxisValues(sourceData: Range)|Set values of X axis for a chart series. Only works for scatter charts.|1.7|
|[chartSeriesCollection](../../api/excel/excel.chartseriescollection)|_Method_ > add(name: string, index: number)|Add a new series to the collection.|1.7|
|[chartTitle](../../api/excel/excel.charttitle)|_Property_ > height|Returns the height, in points, of the chart title. Read-only. Null if chart title's not visible. Read-only.|1.7|
|[chartTitle](../../api/excel/excel.charttitle)|_Property_ > horizontalAlignment|Represents the horizontal alignment for chart title. Possible values are: Center, Left, Justify, Distributed, Right.|1.7|
|[chartTitle](../../api/excel/excel.charttitle)|_Property_ > left|Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title's not visible.|1.7|
|[chartTitle](../../api/excel/excel.charttitle)|_Property_ > position|Represents the position of chart title. Possible values are: Top, Automatic, Bottom, Right, Left.|1.7|
|[chartTitle](../../api/excel/excel.charttitle)|_Property_ > showShadow|Represents a boolean value that determines if the chart title has a shadow.|1.7|
|[chartTitle](../../api/excel/excel.charttitle)|_Property_ > textOrientation|Represents the text orientation of chart title. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|1.7|
|[chartTitle](../../api/excel/excel.charttitle)|_Property_ > top|Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title's not visible.|1.7|
|[chartTitle](../../api/excel/excel.charttitle)|_Property_ > verticalAlignment|Represents the vertical alignment of chart title. Possible values are: Center, Bottom, Top, Justify, Distributed.|1.7|
|[chartTitle](../../api/excel/excel.charttitle)|_Property_ > width|Returns the width, in points, of the chart title. Read-only. Null if chart title's not visible. Read-only.|1.7|
|[chartTitle](../../api/excel/excel.charttitle)|_Method_ > setFormula(formula: string)|Sets a string value that represents the formula of chart title using A1-style notation.|1.7|
|[chartTitleFormat](../../api/excel/excel.charttitleformat)|_Relationship_ > border|Represents the border format of chart title, which includes color, linestyle and weight. Read-only.|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Property_ > backward|Represents the number of periods that the trendline extends backward.|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Property_ > displayEquation|True if the equation for the trendline is displayed on the chart.|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Property_ > displayRSquared|True if the R-squared for the trendline is displayed on the chart.|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Property_ > forward|Represents the number of periods that the trendline extends forward.|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Property_ > intercept|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Property_ > movingAveragePeriod|Represents the period of a chart trendline, only for trendline with MovingAverage type.|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Property_ > name|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Property_ > polynomialOrder|Represents the order of a chart trendline, only for trendline with Polynomial type.|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Property_ > type|Represents the type of a chart trendline. Possible values are: Linear, Exponential, Logarithmic, MovingAverage, Polynomial, Power.|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Relationship_ > format|Represents the formatting of a chart trendline. Read-only.|1.7|
|[chartTrendline](../../api/excel/excel.charttrendline)|_Method_ > delete()|Delete the trendline object.|1.7|
|[chartTrendlineCollection](../../api/excel/excel.charttrendlinecollection)|_Property_ > items|A collection of chartTrendline objects. Read-only.|1.7|
|[chartTrendlineCollection](../../api/excel/excel.charttrendlinecollection)|_Method_ > add(type: string)|Adds a new trendline to trendline collection.|1.7|
|[chartTrendlineCollection](../../api/excel/excel.charttrendlinecollection)|_Method_ > getCount()|Returns the number of trendlines in the collection.|1.7|
|[chartTrendlineCollection](../../api/excel/excel.charttrendlinecollection)|_Method_ > getItem(index: number)|Get trendline object by index, which is the insertion order in items array.|1.7|
|[chartTrendlineFormat](../../api/excel/excel.charttrendlineformat)|_Relationship_ > line|Represents chart line formatting. Read-only.|1.7|
|[customFunctionPostProcessAction](../../api/excel/excel.customfunctionpostprocessaction)|_Property_ > actionType|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent. Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|1.7|
|[customFunctionPostProcessAction](../../api/excel/excel.customfunctionpostprocessaction)|_Property_ > operationType|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent. Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|1.7|
|[customProperty](../../api/excel/excel.customproperty)|_Property_ > key|Gets the key of the custom property. Read only. Read-only.|1.7|
|[customProperty](../../api/excel/excel.customproperty)|_Property_ > type|Gets the value type of the custom property. Read only. Read-only. Possible values are: Number, Boolean, Date, String, Float.|1.7|
|[customProperty](../../api/excel/excel.customproperty)|_Property_ > value|Gets or sets the value of the custom property.|1.7|
|[customProperty](../../api/excel/excel.customproperty)|_Method_ > delete()|Deletes the custom property.|1.7|
|[customPropertyCollection](../../api/excel/excel.custompropertycollection)|_Property_ > items|A collection of customProperty objects. Read-only.|1.7|
|[customPropertyCollection](../../api/excel/excel.custompropertycollection)|_Method_ > add(key: string, value: object)|Creates a new or sets an existing custom property.|1.7|
|[customPropertyCollection](../../api/excel/excel.custompropertycollection)|_Method_ > deleteAll()|Deletes all custom properties in this collection.|1.7|
|[customPropertyCollection](../../api/excel/excel.custompropertycollection)|_Method_ > getCount()|Gets the count of custom properties.|1.7|
|[customPropertyCollection](../../api/excel/excel.custompropertycollection)|_Method_ > getItem(key: string)|Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.|1.7|
|[customPropertyCollection](../../api/excel/excel.custompropertycollection)|_Method_ > getItemOrNullObject(key: string)|Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.|1.7|
|[dataConnectionCollection](../../api/excel/excel.dataconnectioncollection)|_Property_ > items|A collection of dataConnection objects. Read-only.|1.7|
|[dataConnectionCollection](../../api/excel/excel.dataconnectioncollection)|_Method_ > refreshAll()|Refreshes all the Data Connections in the collection.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Property_ > author|Gets or sets the author of the workbook.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Property_ > category|Gets or sets the category of the workbook.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Property_ > comments|Gets or sets the comments of the workbook.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Property_ > company|Gets or sets the company of the workbook.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Property_ > keywords|Gets or sets the keywords of the workbook.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Property_ > lastAuthor|Gets the last author of the workbook. Read only. Read-only.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Property_ > manager|Gets or sets the manager of the workbook.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Property_ > revisionNumber|Gets the revision number of the workbook. Read only.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Property_ > subject|Gets or sets the subject of the workbook.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Property_ > title|Gets or sets the title of the workbook.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Relationship_ > creationDate|Gets the creation date of the workbook. Read only. Read-only.|1.7|
|[documentProperties](../../api/excel/excel.documentproperties)|_Relationship_ > custom|Gets the collection of custom properties of the workbook. Read only. Read-only.|1.7|
|[listDataValidation](../../api/excel/excel.listdatavalidation)|_Property_ > source|Source of the list for data validation|1.8|
|[namedItem](../../api/excel/excel.nameditem)|_Property_ > formula|Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|1.7|
|[namedItem](../../api/excel/excel.nameditem)|_Relationship_ > arrayValues|Returns an object containing values and types of the named item. Read-only.|1.7|
|[namedItemArrayValues](../../api/excel/excel.nameditemarrayvalues)|_Property_ > types|Represents the types for each item in the named item array Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.7|
|[namedItemArrayValues](../../api/excel/excel.nameditemarrayvalues)|_Property_ > values|Represents the values of each item in the named item array. Read-only.|1.7|
|[range](../../api/excel/excel.range)|_Property_ > isEntireColumn|Represents if the current range is an entire column. Read-only.|1.7|
|[range](../../api/excel/excel.range)|_Property_ > isEntireRow|Represents if the current range is an entire row. Read-only.|1.7|
|[range](../../api/excel/excel.range)|_Property_ > numberFormatLocal|Represents Excel's number format code for the given range as a string in the language of the user.|1.7|
|[range](../../api/excel/excel.range)|_Property_ > style|Represents the style of the current range. This return either null or a string.|1.7|
|[range](../../api/excel/excel.range)|_Method_ > getAbsoluteResizedRange(numRows: number, numColumns: number)|Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.|1.7|
|[range](../../api/excel/excel.range)|_Method_ > getImage()|Renders the range as a base64-encoded image.|1.7|
|[range](../../api/excel/excel.range)|_Method_ > getSurroundingRegion()|Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|1.7|
|[range](../../api/excel/excel.range)|_Method_ > showCard()|Displays the card for an active cell if it has rich value content.|1.7|
|[rangeFormat](../../api/excel/excel.rangeformat)|_Property_ > textOrientation|Gets or sets the text orientation of all the cells within the range.|1.7|
|[rangeFormat](../../api/excel/excel.rangeformat)|_Property_ > useStandardHeight|Determines if the row height of the Range object equals the standard height of the sheet.|1.7|
|[rangeFormat](../../api/excel/excel.rangeformat)|_Property_ > useStandardWidth|Determines if the columnwidth of the Range object equals the standard width of the sheet.|1.7|
|[rangeHyperlink](../../api/excel/excel.rangehyperlink)|_Property_ > address|Represents the url target for the hyperlink.|1.7|
|[rangeHyperlink](../../api/excel/excel.rangehyperlink)|_Property_ > document..|Represents the document .. target for the hyperlink.|1.7|
|[rangeHyperlink](../../api/excel/excel.rangehyperlink)|_Property_ > screenTip|Represents the string displayed when hovering over the hyperlink.|1.7|
|[rangeHyperlink](../../api/excel/excel.rangehyperlink)|_Property_ > textToDisplay|Represents the string that is displayed in the top left most cell in the range.|1.7|
|[registerEventPostProcessAction](../../api/excel/excel.registereventpostprocessaction)|_Property_ > actionType|Gets the top border Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|1.7|
|[registerEventPostProcessAction](../../api/excel/excel.registereventpostprocessaction)|_Property_ > message|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.7|
|[registerEventPostProcessAction](../../api/excel/excel.registereventpostprocessaction)|_Property_ > messageType|Gets the top border Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|1.7|
|[registerEventPostProcessAction](../../api/excel/excel.registereventpostprocessaction)|_Property_ > targetId|Gets the top border Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|1.7|
|[registerEventPostProcessAction](../../api/excel/excel.registereventpostprocessaction)|_Relationship_ > controlId|Gets the top border Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|1.7|
|[showCardPostProcessAction](../../api/excel/excel.showcardpostprocessaction)|_Property_ > actionType|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent. Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|1.7|
|[showCardPostProcessAction](../../api/excel/excel.showcardpostprocessaction)|_Property_ > column|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent. Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|1.7|
|[showCardPostProcessAction](../../api/excel/excel.showcardpostprocessaction)|_Property_ > row|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent. Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > addIndent|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > autoIndent|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > builtIn|Indicates if the style is a built-in style. Read-only.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > formulaHidden|Indicates if the formula will be hidden when the worksheet is protected.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > horizontalAlignment|Represents the horizontal alignment for the style. Possible values are: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > includeAlignment|Indicates if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > includeBorder|Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > includeFont|Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > includeNumber|Indicates if the style includes the NumberFormat property.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > includePatterns|Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > includeProtection|Indicates if the style includes the FormulaHidden and Locked protection properties.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > indentLevel|An integer from 0 to 250 that indicates the indent level for the style.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > locked|Indicates if the object is locked when the worksheet is protected.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > name|The name of the style. Read-only.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > numberFormat|The format code of the number format for the style.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > numberFormatLocal|The localized format code of the number format for the style.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > orientation|The text orientation for the style.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > readingOrder|The reading order for the style. Possible values are: Context, LeftToRight, RightToLeft.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > shrinkToFit|Indicates if text automatically shrinks to fit in the available column width.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > textOrientation|The text orientation for the style.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > verticalAlignment|Represents the vertical alignment for the style. Possible values are: Top, Center, Bottom, Justify, Distributed.|1.7|
|[style](../../api/excel/excel.style)|_Property_ > wrapText|Indicates if Microsoft Excel wraps the text in the object.|1.7|
|[style](../../api/excel/excel.style)|_Relationship_ > borders|A Border collection of four Border objects that represent the style of the four borders. Read-only.|1.7|
|[style](../../api/excel/excel.style)|_Relationship_ > fill|The Fill of the style. Read-only.|1.7|
|[style](../../api/excel/excel.style)|_Relationship_ > font|A Font object that represents the font of the style. Read-only.|1.7|
|[style](../../api/excel/excel.style)|_Method_ > delete()|Deletes this style.|1.7|
|[styleCollection](../../api/excel/excel.stylecollection)|_Property_ > items|A collection of style objects. Read-only.|1.7|
|[styleCollection](../../api/excel/excel.stylecollection)|_Method_ > add(name: string)]|Adds a new style to the collection.|1.7|
|[styleCollection](../../api/excel/excel.stylecollection)|_Method_ > getItem(name: string)|Gets a style by name.|1.7|
|[tableChangedEventArgs](../../api/excel/excel.tablechangedeventargs)|_Property_ > address|Gets the address that represents the changed area of a table on a specific worksheet.|1.7|
|[tableChangedEventArgs](../../api/excel/excel.tablechangedeventargs)|_Property_ > changeType|Gets the change type that represents how the Changed event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[tableChangedEventArgs](../../api/excel/excel.tablechangedeventargs)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|1.7|
|[tableChangedEventArgs](../../api/excel/excel.tablechangedeventargs)|_Property_ > tableId|Gets the id of the table in which the data changed.|1.7|
|[tableChangedEventArgs](../../api/excel/excel.tablechangedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableChangedEventArgs](../../api/excel/excel.tablechangedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|1.7|
|[tableSelectionChangedEventArgs](../../api/excel/excel.tableselectionchangedeventargs)|_Property_ > address|Gets the range address that represents the selected area of the table on a specific worksheet.|1.7|
|[tableSelectionChangedEventArgs](../../api/excel/excel.tableselectionchangedeventargs)|_Property_ > isInsideTable|Indicates if the selection is inside a table, address will be useless if IsInsideTable is false.|1.7|
|[tableSelectionChangedEventArgs](../../api/excel/excel.tableselectionchangedeventargs)|_Property_ > tableId|Gets the id of the table in which the selection changed.|1.7|
|[tableSelectionChangedEventArgs](../../api/excel/excel.tableselectionchangedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableSelectionChangedEventArgs](../../api/excel/excel.tableselectionchangedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the selection changed.|1.7|
|[workbook](../../api/excel/excel.workbook)|_Property_ > name|Gets the workbook name. Read-only.|1.7|
|[workbook](../../api/excel/excel.workbook)|_Relationship_ > dataConnections|Refreshes all data connections in the workbook. Read-only.|1.7|
|[workbook](../../api/excel/excel.workbook)|_Relationship_ > properties|Gets the workbook properties. Read-only.|1.7|
|[workbook](../../api/excel/excel.workbook)|_Relationship_ > protection|Returns workbook protection object for a workbook. Read-only.|1.7|
|[workbook](../../api/excel/excel.workbook)|_Relationship_ > styles|Represents a collection of styles associated with the workbook. Read-only.|1.7|
|[workbook](../../api/excel/excel.workbook)|_Method_ > getActiveCell()|Gets the currently active cell from the workbook.|1.7|
|[workbookProtection](../../api/excel/excel.workbookprotection)|_Property_ > protected|Indicates if the workbook is protected. Read-Only. Read-only.|1.7|
|[workbookProtection](../../api/excel/excel.workbookprotection)|_Method_ > protect(password: string)|Protects a workbook. Fails if the workbook has been protected.|1.7|
|[workbookProtection](../../api/excel/excel.workbookprotection)|_Method_ > unprotect(password: string)|Unprotects a workbook.|1.7|
|[worksheet](../../api/excel/excel.worksheet)|_Property_ > gridlines|Gets or sets the worksheet's gridlines flag.|1.7|
|[worksheet](../../api/excel/excel.worksheet)|_Property_ > headings|Gets or sets the worksheet's headings flag.|1.7|
|[worksheet](../../api/excel/excel.worksheet)|_Property_ > showHeadings|Gets or sets the worksheet's headings flag.|1.7|
|[worksheet](../../api/excel/excel.worksheet)|_Property_ > standardHeight|Returns the standard (default) height of all the rows in the worksheet, in points. Read-only.|1.7|
|[worksheet](../../api/excel/excel.worksheet)|_Property_ > standardWidth|Returns or sets the standard (default) width of all the columns in the worksheet.|1.7|
|[worksheet](../../api/excel/excel.worksheet)|_Property_ > tabColor|Gets or sets the worksheet tab color.|1.7|
|[worksheet](../../api/excel/excel.worksheet)|_Relationship_ > freezePanes|Gets an object that can be used to manipulate frozen panes on the worksheet Read-only.|1.7|
|[worksheet](../../api/excel/excel.worksheet)|_Method_ > copy(positionType: WorksheetPositionType, relativeTo: Worksheet)|Copy a worksheet and place it at the specified position. Return the copied worksheet.|1.7|
|[worksheet](../../api/excel/excel.worksheet)|_Method_ > getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)|Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|1.7|
|[worksheetActivatedEventArgs](../../api/excel/excel.worksheetactivatedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetActivatedEventArgs](../../api/excel/excel.worksheetactivatedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet that is activated.|1.7|
|[worksheetAddedEventArgs](../../api/excel/excel.worksheetaddedeventargs)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|1.7|
|[worksheetAddedEventArgs](../../api/excel/excel.worksheetaddedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetAddedEventArgs](../../api/excel/excel.worksheetaddedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet that is added to the workbook.|1.7|
|[worksheetChangedEventArgs](../../api/excel/excel.worksheetchangedeventargs)|_Property_ > address|Gets the range address that represents the changed area of a specific worksheet.|1.7|
|[worksheetChangedEventArgs](../../api/excel/excel.worksheetchangedeventargs)|_Property_ > changeType|Gets the change type that represents how the Changed event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[worksheetChangedEventArgs](../../api/excel/excel.worksheetchangedeventargs)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|1.7|
|[worksheetChangedEventArgs](../../api/excel/excel.worksheetchangedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetChangedEventArgs](../../api/excel/excel.worksheetchangedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|1.7|
|[worksheetDeactivatedEventArgs](../../api/excel/excel.worksheetdeactivatedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeactivatedEventArgs](../../api/excel/excel.worksheetdeactivatedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet that is deactivated.|1.7|
|[worksheetDeletedEventArgs](../../api/excel/excel.worksheetdeletedeventargs)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|1.7|
|[worksheetDeletedEventArgs](../../api/excel/excel.worksheetdeletedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeletedEventArgs](../../api/excel/excel.worksheetdeletedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet that is deleted from the workbook.|1.7|
|[worksheetFreezePanes](../../api/excel/excel.worksheetfreezepanes)|_Method_ > freezeAt(frozenRange: Range or string)|Sets the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](../../api/excel/excel.worksheetfreezepanes)|_Method_ > freezeColumns(count: number)|Freeze the first column(s) of the worksheet in place.|1.7|
|[worksheetFreezePanes](../../api/excel/excel.worksheetfreezepanes)|_Method_ > freezeRows(count: number)|Freeze the top row(s) of the worksheet in place.|1.7|
|[worksheetFreezePanes](../../api/excel/excel.worksheetfreezepanes)|_Method_ > getLocation()|Gets a range that describes the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](../../api/excel/excel.worksheetfreezepanes)|_Method_ > getLocationOrNullObject()|Gets a range that describes the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](../../api/excel/excel.worksheetfreezepanes)|_Method_ > unfreeze()|Removes all frozen panes in the worksheet.|1.7|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowEditObjects|Represents the worksheet protection option of allowing editing objects.|1.7|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowEditScenarios|Represents the worksheet protection option of allowing editing scenarios.|1.7|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Relationship_ > selectionMode|Represents the worksheet protection option of selection mode.|1.7|
|[worksheetSelectionChangedEventArgs](../../api/excel/excel.worksheetselectionchangedeventargs)|_Property_ > address|Gets the range address that represents the selected area of a specific worksheet.|1.7|
|[worksheetSelectionChangedEventArgs](../../api/excel/excel.worksheetselectionchangedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetSelectionChangedEventArgs](../../api/excel/excel.worksheetselectionchangedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the selection changed.|1.7|


## What's new in Excel JavaScript API 1.6 

### Conditional formatting

Introduces conditional formating of a range. Allows the following types of conditional formatting:

* Color scale
* Data bar
* Icon set
* Custom

In addition:

* Returns the range the conditional format is applied to. 
* Removal of conditional formatting. 
* Provides priority and stopifTrue capability. 
* Get collection of all conditional formatting on a given range. 
* Clears all conditional formats active on the current specified range. 

|Object| What is new| Description|Requirement set|
|:----|:----|:----|:----|
|[application](../../api/excel/excel.application)|_Method_ > suspendApiCalculationUntilNextSync()|Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.|1.6|
|[cellValueConditionalFormat](../../api/excel/excel.cellvalueconditionalformat)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[cellValueConditionalFormat](../../api/excel/excel.cellvalueconditionalformat)|_Relationship_ > rule|Represents the Rule object on this conditional format.|1.6|
|[colorScaleConditionalFormat](../../api/excel/excel.colorscaleconditionalformat)|_Property_ > threeColorScale|If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum). Read-only.|1.6|
|[colorScaleConditionalFormat](../../api/excel/excel.colorscaleconditionalformat)|_Relationship_ > criteria|The criteria of the color scale. Midpoint is optional when using a two point color scale.|1.6|
|[conditionalCellValueRule](../../api/excel/excel.conditionalcellvaluerule)|_Property_ > formula1|The formula, if required, to evaluate the conditional format rule on.|1.6|
|[conditionalCellValueRule](../../api/excel/excel.conditionalcellvaluerule)|_Property_ > formula2|The formula, if required, to evaluate the conditional format rule on.|1.6|
|[conditionalCellValueRule](../../api/excel/excel.conditionalcellvaluerule)|_Property_ > operator|The operator of the text conditional format. Possible values are: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](../../api/excel/excel.conditionalcolorscalecriteria)|_Relationship_ > maximum|The maximum point Color Scale Criterion.|1.6|
|[conditionalColorScaleCriteria](../../api/excel/excel.conditionalcolorscalecriteria)|_Relationship_ > midpoint|The midpoint Color Scale Criterion if the color scale is a 3-color scale.|1.6|
|[conditionalColorScaleCriteria](../../api/excel/excel.conditionalcolorscalecriteria)|_Relationship_ > minimum|The minimum point Color Scale Criterion.|1.6|
|[conditionalColorScaleCriterion](../../api/excel/excel.conditionalcolorscalecriterion)|_Property_ > color|HTML color code representation of the color scale color. E.g. #FF0000 represents Red.|1.6|
|[conditionalColorScaleCriterion](../../api/excel/excel.conditionalcolorscalecriterion)|_Property_ > formula|A number, a formula, or null (if Type is LowestValue).|1.6|
|[conditionalColorScaleCriterion](../../api/excel/excel.conditionalcolorscalecriterion)|_Property_ > type|What the icon conditional formula should be based on. Possible values are: Invalid, LowestValue, HighestValue, Number, Percent, Formula, Percentile.|1.6|
|[conditionalDataBarNegativeFormat](../../api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > borderColor|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarNegativeFormat](../../api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > fillColor|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarNegativeFormat](../../api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > matchPositiveBorderColor|Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.|1.6|
|[conditionalDataBarNegativeFormat](../../api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > matchPositiveFillColor|Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.|1.6|
|[conditionalDataBarPositiveFormat](../../api/excel/excel.conditionaldatabarpositiveformat)|_Property_ > borderColor|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarPositiveFormat](../../api/excel/excel.conditionaldatabarpositiveformat)|_Property_ > fillColor|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarPositiveFormat](../../api/excel/excel.conditionaldatabarpositiveformat)|_Property_ > gradientFill|Boolean representation of whether or not the DataBar has a gradient.|1.6|
|[conditionalDataBarRule](../../api/excel/excel.conditionaldatabarrule)|_Property_ > formula|The formula, if required, to evaluate the databar rule on.|1.6|
|[conditionalDataBarRule](../../api/excel/excel.conditionaldatabarrule)|_Property_ > type|The type of rule for the databar. Possible values are: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Property_ > id|The Priority of the Conditional Format within the current ConditionalFormatCollection. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Property_ > priority|The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Property_ > stopIfTrue|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Property_ > type|A type of conditional format. Only one can be set at a time. Read-Only. Read-only. Possible values are: Custom, DataBar, ColorScale, IconSet.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > cellValue|Returns the cell value conditional format properties if the current conditional format is a CellValue type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > cellValueOrNullObject|Returns the cell value conditional format properties if the current conditional format is a CellValue type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > colorScale|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > colorScaleOrNullObject|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > custom|Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > customOrNullObject|Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > dataBar|Returns the data bar properties if the current conditional format is a data bar. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > dataBarOrNullObject|Returns the data bar properties if the current conditional format is a data bar. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > iconSet|Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > iconSetOrNullObject|Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > preset|Returns the preset criteria conditional format such as above averagebelow averageunique valuescontains blanknonblankerrornoerror properties. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > presetOrNullObject|Returns the preset criteria conditional format such as above averagebelow averageunique valuescontains blanknonblankerrornoerror properties. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > textComparison|Returns the specific text conditional format properties if the current conditional format is a text type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > textComparisonOrNullObject|Returns the specific text conditional format properties if the current conditional format is a text type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > topBottom|Returns the TopBottom conditional format properties if the current conditional format is an TopBottom type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Relationship_ > topBottomOrNullObject|Returns the TopBottom conditional format properties if the current conditional format is an TopBottom type. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Method_ > delete()|Deletes this conditional format.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Method_ > getRange()|Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.|1.6|
|[conditionalFormat](../../api/excel/excel.conditionalformat)|_Method_ > getRangeOrNullObject()|Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.|1.6|
|[conditionalFormatCollection](../../api/excel/excel.conditionalformatcollection)|_Property_ > items|A collection of conditionalFormat objects. Read-only.|1.6|
|[conditionalFormatCollection](../../api/excel/excel.conditionalformatcollection)|_Method_ > add(type: string)|Adds a new conditional format to the collection at the firsttop priority.|1.6|
|[conditionalFormatCollection](../../api/excel/excel.conditionalformatcollection)|_Method_ > clearAll()|Clears all conditional formats active on the current specified range.|1.6|
|[conditionalFormatCollection](../../api/excel/excel.conditionalformatcollection)|_Method_ > getCount()|Returns the number of conditional formats in the workbook. Read-only.|1.6|
|[conditionalFormatCollection](../../api/excel/excel.conditionalformatcollection)|_Method_ > getItem(id: string)|Returns a conditional format for the given ID.|1.6|
|[conditionalFormatCollection](../../api/excel/excel.conditionalformatcollection)|_Method_ > getItemAt(index: number)|Returns a conditional format at the given index.|1.6|
|[conditionalFormatRule](../../api/excel/excel.conditionalformatrule)|_Property_ > formula|The formula, if required, to evaluate the conditional format rule on.|1.6|
|[conditionalFormatRule](../../api/excel/excel.conditionalformatrule)|_Property_ > formulaLocal|The formula, if required, to evaluate the conditional format rule on in the user's language.|1.6|
|[conditionalFormatRule](../../api/excel/excel.conditionalformatrule)|_Property_ > formulaR1C1|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|1.6|
|[conditionalIconCriterion](../../api/excel/excel.conditionaliconcriterion)|_Property_ > formula|A number or a formula depending on the type.|1.6|
|[conditionalIconCriterion](../../api/excel/excel.conditionaliconcriterion)|_Property_ > operator|GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format. Possible values are: Invalid, GreaterThan, GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](../../api/excel/excel.conditionaliconcriterion)|_Relationship_ > customIcon|The custom icon for the current criterion if different from the default IconSet, else null will be returned.|1.6|
|[conditionalIconCriterion](../../api/excel/excel.conditionaliconcriterion)|_Relationship_ > type|What the icon conditional formula should be based on.|1.6|
|[conditionalPresetCriteriaRule](../../api/excel/excel.conditionalpresetcriteriarule)|_Property_ > criterion|The criterion of the conditional format. Possible values are: Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues, DuplicateValues.|1.6|
|[conditionalRangeBorder](../../api/excel/excel.conditionalrangeborder)|_Property_ > color|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalRangeBorder](../../api/excel/excel.conditionalrangeborder)|_Property_ > id|Represents border identifier. Read-only. Possible values are: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](../../api/excel/excel.conditionalrangeborder)|_Property_ > sideIndex|Constant value that indicates the specific side of the border. Read-only. Possible values are: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](../../api/excel/excel.conditionalrangeborder)|_Property_ > style|One of the constants of line style specifying the line style for the border. Possible values are: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.6|
|[conditionalRangeBorderCollection](../../api/excel/excel.conditionalrangebordercollection)|_Property_ > count|Number of border objects in the collection. Read-only.|1.6|
|[conditionalRangeBorderCollection](../../api/excel/excel.conditionalrangebordercollection)|_Property_ > items|A collection of conditionalRangeBorder objects. Read-only.|1.6|
|[conditionalRangeBorderCollection](../../api/excel/excel.conditionalrangebordercollection)|_Relationship_ > bottom|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](../../api/excel/excel.conditionalrangebordercollection)|_Relationship_ > left|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](../../api/excel/excel.conditionalrangebordercollection)|_Relationship_ > right|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](../../api/excel/excel.conditionalrangebordercollection)|_Relationship_ > top|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](../../api/excel/excel.conditionalrangebordercollection)|_Method_ > getItem(index: string)|Gets a border object using its name|1.6|
|[conditionalRangeBorderCollection](../../api/excel/excel.conditionalrangebordercollection)|_Method_ > getItemAt(index: number)|Gets a border object using its index|1.6|
|[conditionalRangeFill](../../api/excel/excel.conditionalrangefill)|_Property_ > color|HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalRangeFill](../../api/excel/excel.conditionalrangefill)|_Method_ > clear()|Resets the fill.|1.6|
|[conditionalRangeFont](../../api/excel/excel.conditionalrangefont)|_Property_ > bold|Represents the bold status of font.|1.6|
|[conditionalRangeFont](../../api/excel/excel.conditionalrangefont)|_Property_ > color|HTML color code representation of the text color. E.g. #FF0000 represents Red.|1.6|
|[conditionalRangeFont](../../api/excel/excel.conditionalrangefont)|_Property_ > italic|Represents the italic status of the font.|1.6|
|[conditionalRangeFont](../../api/excel/excel.conditionalrangefont)|_Property_ > strikethrough|Represents the strikethrough status of the font.|1.6|
|[conditionalRangeFont](../../api/excel/excel.conditionalrangefont)|_Property_ > underline|Type of underline applied to the font. Possible values are: None, Single, Double.|1.6|
|[conditionalRangeFont](../../api/excel/excel.conditionalrangefont)|_Method_ > clear()|Resets the font formats.|1.6|
|[conditionalRangeFormat](../../api/excel/excel.conditionalrangeformat)|_Property_ > numberFormat|Represents Excel's number format code for the given range. Cleared if null is passed in.|1.6|
|[conditionalRangeFormat](../../api/excel/excel.conditionalrangeformat)|_Relationship_ > borders|Collection of border objects that apply to the overall conditional format range. Read-only.|1.6|
|[conditionalRangeFormat](../../api/excel/excel.conditionalrangeformat)|_Relationship_ > fill|Returns the fill object defined on the overall conditional format range. Read-only.|1.6|
|[conditionalRangeFormat](../../api/excel/excel.conditionalrangeformat)|_Relationship_ > font|Returns the font object defined on the overall conditional format range. Read-only.|1.6|
|[conditionalTextComparisonRule](../../api/excel/excel.conditionaltextcomparisonrule)|_Property_ > operator|The operator of the text conditional format. Possible values are: Invalid, Contains, NotContains, BeginsWith, EndsWith.|1.6|
|[conditionalTextComparisonRule](../../api/excel/excel.conditionaltextcomparisonrule)|_Property_ > text|The Text value of conditional format.|1.6|
|[conditionalTopBottomRule](../../api/excel/excel.conditionaltopbottomrule)|_Property_ > rank|The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.|1.6|
|[conditionalTopBottomRule](../../api/excel/excel.conditionaltopbottomrule)|_Property_ > type|Format values based on the top or bottom rank. Possible values are: Invalid, TopItems, TopPercent, BottomItems, BottomPercent.|1.6|
|[customConditionalFormat](../../api/excel/excel.customconditionalformat)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[customConditionalFormat](../../api/excel/excel.customconditionalformat)|_Relationship_ > rule|Represents the Rule object on this conditional format. Read-only.|1.6|
|[dataBarConditionalFormat](../../api/excel/excel.databarconditionalformat)|_Property_ > axisColor|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[dataBarConditionalFormat](../../api/excel/excel.databarconditionalformat)|_Property_ > axisFormat|Representation of how the axis is determined for an Excel data bar. Possible values are: Automatic, None, CellMidPoint.|1.6|
|[dataBarConditionalFormat](../../api/excel/excel.databarconditionalformat)|_Property_ > barDirection|Represents the direction that the data bar graphic should be based on. Possible values are: Context, LeftToRight, RightToLeft.|1.6|
|[dataBarConditionalFormat](../../api/excel/excel.databarconditionalformat)|_Property_ > showDataBarOnly|If true, hides the values from the cells where the data bar is applied.|1.6|
|[dataBarConditionalFormat](../../api/excel/excel.databarconditionalformat)|_Relationship_ > lowerBoundRule|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|1.6|
|[dataBarConditionalFormat](../../api/excel/excel.databarconditionalformat)|_Relationship_ > negativeFormat|Representation of all values to the left of the axis in an Excel data bar. Read-only.|1.6|
|[dataBarConditionalFormat](../../api/excel/excel.databarconditionalformat)|_Relationship_ > positiveFormat|Representation of all values to the right of the axis in an Excel data bar. Read-only.|1.6|
|[dataBarConditionalFormat](../../api/excel/excel.databarconditionalformat)|_Relationship_ > upperBoundRule|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|1.6|
|[iconSetConditionalFormat](../../api/excel/excel.iconsetconditionalformat)|_Property_ > reverseIconOrder|If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.|1.6|
|[iconSetConditionalFormat](../../api/excel/excel.iconsetconditionalformat)|_Property_ > showIconOnly|If true, hides the values and only shows icons.|1.6|
|[iconSetConditionalFormat](../../api/excel/excel.iconsetconditionalformat)|_Property_ > style|If set, displays the IconSet option for the conditional format. Possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](../../api/excel/excel.iconsetconditionalformat)|_Relationship_ > criteria|An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula and operator will be ignored when set.|1.6|
|[presetCriteriaConditionalFormat](../../api/excel/excel.presetcriteriaconditionalformat)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[presetCriteriaConditionalFormat](../../api/excel/excel.presetcriteriaconditionalformat)|_Relationship_ > rule|The rule of the conditional format.|1.6|
|[range](../../api/excel/excel.range)|_Relationship_ > conditionalFormats|Collection of ConditionalFormats that intersect the range. Read-only.|1.6|
|[range](../../api/excel/excel.range)|_Method_ > calculate()|Calculates a range of cells on a worksheet.|1.6|
|[textConditionalFormat](../../api/excel/excel.textconditionalformat)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[textConditionalFormat](../../api/excel/excel.textconditionalformat)|_Relationship_ > rule|The rule of the conditional format.|1.6|
|[topBottomConditionalFormat](../../api/excel/excel.topbottomconditionalformat)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[topBottomConditionalFormat](../../api/excel/excel.topbottomconditionalformat)|_Relationship_ > rule|The criteria of the TopBottom conditional format.|1.6|
|[workbook](../../api/excel/excel.workbook)|_Relationship_ > internalTest|For internal use only. Read-only.|1.6|
|[worksheet](../../api/excel/excel.worksheet)|_Method_ > calculate(markAllDirty: bool)|Calculates all cells on a worksheet.|1.6|

##  What's new in Excel JavaScript API 1.5

### Custom XML part

* Addition of custom XML parts collection to workbook object.
* Get custom XML part using ID
* Get a new scoped collection of custom XML parts whose namespaces match the given namespace.
* Get XML string associated with a part.
* Provide id and namespace of a part.
* Adds a new custom XML part to the workbook.
* Set entire XML part.
* Delete a custom XML part.
* Delete an attribute with the given name from the element identified by xpath.
* Query the XML content by xpath.
* Insert, update and delete attribute.

**Reference implementation:** Please refer [here](https://github.com/mandren/Excel-CustomXMLPart-Demo) for a reference implementation that shows how custom XML parts can be used in an add-in.

### Others
* `range.getSurroundingRegion()` Returns a Range object that represents the surrounding region for this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.
* `getNextColumn()` and `getPreviousColumn()`, `getLast() on table column.
* `getActiveWorksheet()` on the workbook.
* `getRange(address: string)` off of workbook.
* `getBoundingRange(ranges: )` Gets the smallest range object that encompasses the provided ranges. For example, the bounding range between "B2:C5" and "D10:E15" is "B2:E15".
* `getCount()` on various collections such as named item, worksheet, table, etc. to get number of items in a collection. `workbook.worksheets.getCount()`
* `getFirst()` and `getLast()` and get last on various collection such as tworksheet, able column, chart points, range view collection.
* `getNext()` and `getPrevious()` on worksheet, table column collection.
* `getRangeR1C1()` Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.

|Object| What is new| Description|Requirement set|
|:----|:----|:----|:----|
|[customXmlPart](../../api/excel/excel.customxmlpart)|_Property_ > id|The custom XML part's ID. Read-only.|1.5|
|[customXmlPart](../../api/excel/excel.customxmlpart)|_Property_ > namespaceUri|The custom XML part's namespace URI. Read-only.|1.5|
|[customXmlPart](../../api/excel/excel.customxmlpart)|_Method_ > delete()|Deletes the custom XML part.|1.5|
|[customXmlPart](../../api/excel/excel.customxmlpart)|_Method_ > getXml()|Gets the custom XML part's full XML content.|1.5|
|[customXmlPart](../../api/excel/excel.customxmlpart)|_Method_ > setXml(xml: string)|Sets the custom XML part's full XML content.|1.5|
|[customXmlPartCollection](../../api/excel/excel.customxmlpartcollection)|_Property_ > items|A collection of customXmlPart objects. Read-only.|1.5|
|[customXmlPartCollection](../../api/excel/excel.customxmlpartcollection)|_Method_ > add(xml: string)|Adds a new custom XML part to the workbook.|1.5|
|[customXmlPartCollection](../../api/excel/excel.customxmlpartcollection)|_Method_ > getByNamespace(namespaceUri: string)|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|1.5|
|[customXmlPartCollection](../../api/excel/excel.customxmlpartcollection)|_Method_ > getCount()|Gets the number of CustomXml parts in the collection.|1.5|
|[customXmlPartCollection](../../api/excel/excel.customxmlpartcollection)|_Method_ > getItem(id: string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartCollection](../../api/excel/excel.customxmlpartcollection)|_Method_ > getItemOrNullObject(id: string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartScopedCollection](../../api/excel/excel.customxmlpartscopedcollection)|_Property_ > items|A collection of customXmlPartScoped objects. Read-only.|1.5|
|[customXmlPartScopedCollection](../../api/excel/excel.customxmlpartscopedcollection)|_Method_ > getCount()|Gets the number of CustomXML parts in this collection.|1.5|
|[customXmlPartScopedCollection](../../api/excel/excel.customxmlpartscopedcollection)|_Method_ > getItem(id: string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartScopedCollection](../../api/excel/excel.customxmlpartscopedcollection)|_Method_ > getItemOrNullObject(id: string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartScopedCollection](../../api/excel/excel.customxmlpartscopedcollection)|_Method_ > getOnlyItem()|If the collection contains exactly one item, this method returns it.|1.5|
|[customXmlPartScopedCollection](../../api/excel/excel.customxmlpartscopedcollection)|_Method_ > getOnlyItemOrNullObject()|If the collection contains exactly one item, this method returns it.|1.5|
|[workbook](../../api/excel/excel.workbook)|_Relationship_ > customXmlParts|Represents the collection of custom XML parts contained by this workbook. Read-only.|1.5|
|[worksheet](../../api/excel/excel.worksheet)|_Method_ > getNext(visibleOnly: bool)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.|1.5|
|[worksheet](../../api/excel/excel.worksheet)|_Method_ > getNextOrNullObject(visibleOnly: bool)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.|1.5|
|[worksheet](../../api/excel/excel.worksheet)|_Method_ > getPrevious(visibleOnly: bool)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.|1.5|
|[worksheet](../../api/excel/excel.worksheet)|_Method_ > getPreviousOrNullObject(visibleOnly: bool)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.|1.5|
|[worksheetCollection](../../api/excel/excel.worksheetcollection)|_Method_ > getFirst(visibleOnly: bool)|Gets the first worksheet in the collection.|1.5|
|[worksheetCollection](../../api/excel/excel.worksheetcollection)|_Method_ > getLast(visibleOnly: bool)|Gets the last worksheet in the collection.|1.5|

## What's new in Excel JavaScript API 1.4
The following are the new additions to the Excel JavaScript APIs in requirement set 1.4.

### Named item add and new properties

New properties:

* `comment`
* `scope` worksheet or workbook scoped items
* `worksheet` returns the worksheet on which the named item is scoped to.

New methods:

* `add(name: string, reference: Range or string, comment: string)`Adds a new name to the collection of the given scope.
* `addFormulaLocal(name: string, formula: string, comment: string)` Adds a new name to the collection of the given scope using the user's locale for the formula.

### Settings API in in Excel namespace

[Setting](../../api/excel/excel.setting) object represents a key-value pair of a setting persisted to the document. Now, we've added settings related APIs under Excel namespace. This doesn't offer net new functionality - however this make easy to remain in the promise based batched API syntax reduce the dependency on common API for Excel related tasks.

APIs include `getItem()` to get setting entry via the key, `add()` to add the specified key:value setting pair to the workbook.

### Others

* Set table column name (prior version only allows reading).
* Add table column to the end of the table (prior version only allows anywhere but last).
* Add multiple rows to a table at a time (prior version only allows 1 row at a time).
* `range.getColumnsAfter(count: number)` and `range.getColumnsBefore(count: number)` to get a certain number of columns to the right/left of the current Range object.
* Get item or null object function: This functionality allows getting object using a key. If the object does not exist, the returned object's isNullObject property will be true. This alows developers to check if an object exists or not without having to handle it thorugh exception handling. Available on worksheet, named-item, binding, chart series, etc.

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|Object| What is new| Description|Requirement set|
|:----|:----|:----|:----|
|[bindingCollection](../../api/excel/excel.bindingcollection)|_Method_ > getCount()|Gets the number of bindings in the collection.|1.4|
|[bindingCollection](../../api/excel/excel.bindingcollection)|_Method_ > getItemOrNullObject(id: string)|Gets a binding object by ID. If the binding object does not exist, will return a null object.|1.4|
|[chartCollection](../../api/excel/excel.chartcollection)|_Method_ > getCount()|Returns the number of charts in the worksheet.|1.4|
|[chartCollection](../../api/excel/excel.chartcollection)|_Method_ > getItemOrNullObject(name: string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.4|
|[chartPointsCollection](../../api/excel/excel.chartpointscollection)|_Method_ > getCount()|Returns the number of chart points in the series.|1.4|
|[chartSeriesCollection](../../api/excel/excel.chartseriescollection)|_Method_ > getCount()|Returns the number of series in the collection.|1.4|
|[namedItem](../../api/excel/excel.nameditem)|_Property_ > comment|Represents the comment associated with this name.|1.4|
|[namedItem](../../api/excel/excel.nameditem)|_Property_ > scope|Indicates whether the name is scoped to the workbook or to a specific worksheet. Read-only. Possible values are: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](../../api/excel/excel.nameditem)|_Relationship_ > worksheet|Returns the worksheet on which the named item is scoped to. Throws an error if the items is scoped to the workbook instead. Read-only.|1.4|
|[namedItem](../../api/excel/excel.nameditem)|_Relationship_ > worksheetOrNullObject|Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead. Read-only.|1.4|
|[namedItem](../../api/excel/excel.nameditem)|_Method_ > delete()|Deletes the given name.|1.4|
|[namedItem](../../api/excel/excel.nameditem)|_Method_ > getRangeOrNullObject()|Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.|1.4|
|[namedItemCollection](../../api/excel/excel.nameditemcollection)|_Method_ > add(name: string, reference: Range or string, comment: string)|Adds a new name to the collection of the given scope.|1.4|
|[namedItemCollection](../../api/excel/excel.nameditemcollection)|_Method_ > addFormulaLocal(name: string, formula: string, comment: string)|Adds a new name to the collection of the given scope using the user's locale for the formula.|1.4|
|[namedItemCollection](../../api/excel/excel.nameditemcollection)|_Method_ > getCount()|Gets the number of named items in the collection.|1.4|
|[namedItemCollection](../../api/excel/excel.nameditemcollection)|_Method_ > getItemOrNullObject(name: string)|Gets a nameditem object using its name. If the nameditem object does not exist, will return a null object.|1.4|
|[pivotTableCollection](../../api/excel/excel.pivottablecollection)|_Method_ > getCount()|Gets the number of pivot tables in the collection.|1.4|
|[pivotTableCollection](../../api/excel/excel.pivottablecollection)|_Method_ > getItemOrNullObject(name: string)|Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.|1.4|
|[range](../../api/excel/excel.range)|_Method_ > getIntersectionOrNullObject(anotherRange: Range or string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|1.4|
|[range](../../api/excel/excel.range)|_Method_ > getUsedRangeOrNullObject(valuesOnly: bool)|Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.|1.4|
|[rangeViewCollection](../../api/excel/excel.rangeviewcollection)|_Method_ > getCount()|Gets the number of RangeView objects in the collection.|1.4|
|[setting](../../api/excel/excel.setting)|_Property_ > key|Returns the key that represents the id of the Setting. Read-only.|1.4|
|[setting](../../api/excel/excel.setting)|_Property_ > value|Represents the value stored for this setting.|1.4|
|[setting](../../api/excel/excel.setting)|_Method_ > delete()|Deletes the setting.|1.4|
|[settingCollection](../../api/excel/excel.settingcollection)|_Property_ > items|A collection of setting objects. Read-only.|1.4|
|[settingCollection](../../api/excel/excel.settingcollection)|_Method_ > add(key: string, value: (any))|Sets or adds the specified setting to the workbook.|1.4|
|[settingCollection](../../api/excel/excel.settingcollection)|_Method_ > getCount()|Gets the number of Settings in the collection.|1.4|
|[settingCollection](../../api/excel/excel.settingcollection)|_Method_ > getItem(key: string)|Gets a Setting entry via the key.|1.4|
|[settingCollection](../../api/excel/excel.settingcollection)|_Method_ > getItemOrNullObject(key: string)|Gets a Setting entry via the key. If the Setting does not exist, will return a null object.|1.4|
|[settingsChangedEventArgs](../../api/excel/excel.settingschangedeventargs)|_Relationship_ > settings|Gets the Setting object that represents the binding that raised the SettingsChanged event|1.4|
|[tableCollection](../../api/excel/excel.tablecollection)|_Method_ > getCount()]|Gets the number of tables in the collection.|1.4|
|[tableCollection](../../api/excel/excel.tablecollection)|_Method_ > getItemOrNullObject(key: number or string)|Gets a table by Name or ID. If the table does not exist, will return a null object.|1.4|
|[tableColumnCollection](../../api/excel/excel.tablecolumncollection)|_Method_ > getCount()|Gets the number of columns in the table.|1.4|
|[tableColumnCollection](../../api/excel/excel.tablecolumncollection)|_Method_ > getItemOrNullObject(key: number or string)|Gets a column object by Name or ID. If the column does not exist, will return a null object.|1.4|
|[tableRowCollection](../../api/excel/excel.tablerowcollection)|_Method_ > getCount()|Gets the number of rows in the table.|1.4|
|[workbook](../../api/excel/excel.workbook)|_Relationship_ > settings|Represents a collection of Settings associated with the workbook. Read-only.|1.4|
|[worksheet](../../api/excel/excel.worksheet)|_Relationship_ > names|Collection of names scoped to the current worksheet. Read-only.|1.4|
|[worksheet](../../api/excel/excel.worksheet)|_Method_ > getUsedRangeOrNullObject(valuesOnly: bool)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.|1.4|
|[worksheetCollection](../../api/excel/excel.worksheetcollection)|_Method_ > getCount(visibleOnly: bool)|Gets the number of worksheets in the collection.|1.4|
|[worksheetCollection](../../api/excel/excel.worksheetcollection)|_Method_ > getItemOrNullObject(key: string)|Gets a worksheet object using its Name or ID. If the worksheet does not exist, will return a null object.|1.4|

## What's new in Excel JavaScript API 1.3

The following are the new additions to the Excel JavaScript APIs in requirement set 1.3.

|Object| What's new| Description|Requirement set|
|:----|:----|:----|:----|
|[binding](../../api/excel/excel.binding)|_Method_ > delete()|Deletes the binding.|1.3|
|[bindingCollection](../../api/excel/excel.bindingcollection)|_Method_ > add(range: Range or string, bindingType: string, id: string)|Add a new binding to a particular Range.|1.3|
|[bindingCollection](../../api/excel/excel.bindingcollection)|_Method_ > addFromNamedItem(name: string, bindingType: string, id: string)|Add a new binding based on a named item in the workbook.|1.3|
|[bindingCollection](../../api/excel/excel.bindingcollection)|_Method_ > addFromSelection(bindingType: string, id: string)|Add a new binding based on the current selection.|1.3|
|[bindingCollection](../../api/excel/excel.bindingcollection)|_Method_ > getItemOrNull(id: string)|Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.|1.3|
|[chartCollection](../../api/excel/excel.chartcollection)|_Method_ > getItemOrNull(name: string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.3|
|[namedItemCollection](../../api/excel/excel.nameditemcollection)|_Method_ > getItemOrNull(name: string)|Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.|1.3|
|[pivotTable](../../api/excel/excel.pivottable)|_Property_ > name|Name of the PivotTable.|1.3|
|[pivotTable](../../api/excel/excel.pivottable)|_Relationship_ > worksheet|The worksheet containing the current PivotTable. Read-only.|1.3|
|[pivotTable](../../api/excel/excel.pivottable)|_Method_ > refresh()|Refreshes the PivotTable.|1.3|
|[pivotTableCollection](../../api/excel/excel.pivottablecollection)|_Property_ > items|A collection of pivotTable objects. Read-only.|1.3|
|[pivotTableCollection](../../api/excel/excel.pivottablecollection)|_Method_ > getItem(name: string)|Gets a PivotTable by name.|1.3|
|[pivotTableCollection](../../api/excel/excel.pivottablecollection)|_Method_ > getItemOrNull(name: string)|Gets a PivotTable by name. If the PivotTable does not exist, the return object's isNull property will be true.|1.3|
|[range](../../api/excel/excel.range)|_Method_ > getIntersectionOrNull(anotherRange: Range or string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|1.3|
|[range](../../api/excel/excel.range)|_Method_ > getVisibleView()|Represents the visible rows of the current range.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > cellAddresses|Represents the cell addresses of the RangeView. Read-only.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > columnCount|Returns the number of visible columns. Read-only.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > formulas|Represents the formula in A1-style notation.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > formulasLocal|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, introduced in 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > index|Returns a value that represents the index of the RangeView. Read-only.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > numberFormat|Represents Excel's number format code for the given cell.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > rowCount|Returns the number of visible rows. Read-only.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > text|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > valueTypes|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Property_ > values|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Relationship_ > rows|Represents a collection of range views associated with the range. Read-only.|1.3|
|[rangeView](../../api/excel/excel.rangeview)|_Method_ > getRange()|Gets the parent range associated with the current RangeView.|1.3|
|[rangeViewCollection](../../api/excel/excel.rangeviewcollection)|_Property_ > items|A collection of rangeView objects. Read-only.|1.3|
|[rangeViewCollection](../../api/excel/excel.rangeviewcollection)|_Method_ > getItemAt(index: number)|Gets a RangeView Row via it's index. Zero-Indexed.|1.3|
|[setting](../../api/excel/excel.setting)|_Property_ > key|Returns the key that represents the id of the Setting. Read-only.|1.3|
|[setting](../../api/excel/excel.setting)|_Method_ > delete()|Deletes the setting.|1.3|
|[settingCollection](../../api/excel/excel.settingcollection)|_Property_ > items|A collection of setting objects. Read-only.|1.3|
|[settingCollection](../../api/excel/excel.settingcollection)|_Method_ > getItem(key: string)|Gets a Setting entry via the key.|1.3|
|[settingCollection](../../api/excel/excel.settingcollection)|_Method_ > getItemOrNull(key: string)|Gets a Setting entry via the key. If the Setting does not exist, the returned object's isNull property will be true.|1.3|
|[settingCollection](../../api/excel/excel.settingcollection)|_Method_ > set(key: string, value: string)|Sets or adds the specified setting to the workbook.|1.3|
|[settingsChangedEventArgs](../../api/excel/excel.settingschangedeventargs)|_Relationship_ > settingCollection|Gets the Setting object that represents the binding that raised the SettingsChanged event|1.3|
|[table](../../api/excel/excel.table)|_Property_ > highlightFirstColumn|Indicates whether the first column contains special formatting.|1.3|
|[table](../../api/excel/excel.table)|_Property_ > highlightLastColumn|Indicates whether the last column contains special formatting.|1.3|
|[table](../../api/excel/excel.table)|_Property_ > showBandedColumns|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|1.3|
|[table](../../api/excel/excel.table)|_Property_ > showBandedRows|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|1.3|
|[table](../../api/excel/excel.table)|_Property_ > showFilterButton|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|1.3|
|[tableCollection](../../api/excel/excel.tablecollection)|_Method_ > getItemOrNull(key: number or string)|Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.|1.3|
|[tableColumnCollection](../../api/excel/excel.tablecolumncollection)|_Method_ > getItemOrNull(key: number or string)|Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.|1.3|
|[workbook](../../api/excel/excel.workbook)|_Relationship_ > pivotTables|Represents a collection of PivotTables associated with the workbook. Read-only.|1.3|
|[workbook](../../api/excel/excel.workbook)|_Relationship_ > settings|Represents a collection of Settings associated with the workbook. Read-only.|1.3|
|[worksheet](../../api/excel/excel.worksheet)|_Relationship_ > pivotTables|Collection of PivotTables that are part of the worksheet. Read-only.|1.3|

## What's new in Excel JavaScript API 1.2

The following are the new additions to the Excel JavaScript APIs in requirement set 1.2.

|Object| What's new| Description|Requirement set|
|:----|:----|:----|:----|
|[chart](../../api/excel/excel.chart)|_Property_ > id|Gets a chart based on its position in the collection. Read-only.|1.2|
|[chart](../../api/excel/excel.chart)|_Relationship_ > worksheet|The worksheet containing the current chart. Read-only.|1.2|
|[chart](../../api/excel/excel.chart)|_Method_ > getImage(height: number, width: number, fittingMode: string)|Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.|1.2|
|[filter](../../api/excel/excel.filter)|_Relationship_ > criteria|The currently applied filter on the given column. Read-only.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > apply(criteria: FilterCriteria)|Apply the given filter criteria on the given column.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > applyBottomItemsFilter(count: number)|Apply a "Bottom Item" filter to the column for the given number of elements.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > applyBottomPercentFilter(percent: number)]|Apply a "Bottom Percent" filter to the column for the given percentage of elements.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > applyCellColorFilter(color: string)|Apply a "Cell Color" filter to the column for the given color.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > applyCustomFilter(criteria1: string, criteria2: string, oper: string)|Apply a "Icon" filter to the column for the given criteria strings.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > applyDynamicFilter(criteria: string)|Apply a "Dynamic" filter to the column.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > applyFontColorFilter(color: string)|Apply a "Font Color" filter to the column for the given color.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > applyIconFilter(icon: Icon)|Apply a "Icon" filter to the column for the given icon.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > applyTopItemsFilter(count: number)|Apply a "Top Item" filter to the column for the given number of elements.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > applyTopPercentFilter(percent: number)|Apply a "Top Percent" filter to the column for the given percentage of elements.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > applyValuesFilter(values: ())|Apply a "Values" filter to the column for the given values.|1.2|
|[filter](../../api/excel/excel.filter)|_Method_ > clear()|Clear the filter on the given column.|1.2|
|[filterCriteria](../../api/excel/excel.filtercriteria)|_Property_ > color|The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering.|1.2|
|[filterCriteria](../../api/excel/excel.filtercriteria)|_Property_ > criterion1|The first criterion used to filter data. Used as an operator in the case of "custom" filtering.|1.2|
|[filterCriteria](../../api/excel/excel.filtercriteria)|_Property_ > criterion2|The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.|1.2|
|[filterCriteria](../../api/excel/excel.filtercriteria)|_Property_ > dynamicCriteria|The dynamic criteria from the Excel.DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering. Possible values are: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](../../api/excel/excel.filtercriteria)|_Property_ > filterOn|The property used by the filter to determine whether the values should stay visible. Possible values are: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](../../api/excel/excel.filtercriteria)|_Property_ > operator|The operator used to combine criterion 1 and 2 when using "custom" filtering. Possible values are: And, Or.|1.2|
|[filterCriteria](../../api/excel/excel.filtercriteria)|_Property_ > values|The set of values to be used as part of "values" filtering.|1.2|
|[filterCriteria](../../api/excel/excel.filtercriteria)|_Relationship_ > icon|The icon used to filter cells. Used with "icon" filtering.|1.2|
|[filterDatetime](../../api/excel/excel.filterdatetime)|_Property_ > date|The date in ISO8601 format used to filter data.|1.2|
|[filterDatetime](../../api/excel/excel.filterdatetime)|_Property_ > specificity|How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009. Possible values are: Year, Monday, Day, Hour, Minute, Second.|1.2|
|[formatProtection](../../api/excel/excel.formatprotection)|_Property_ > formulaHidden|Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.|1.2|
|[formatProtection](../../api/excel/excel.formatprotection)|_Property_ > locked|Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.|1.2|
|[icon](../../api/excel/excel.icon)|_Property_ > index|Represents the index of the icon in the given set.|1.2|
|[icon](../../api/excel/excel.icon)|_Property_ > set|Represents the set that the icon is part of. Possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](../../api/excel/excel.range)|_Property_ > columnHidden|Represents if all columns of the current range are hidden.|1.2|
|[range](../../api/excel/excel.range)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|1.2|
|[range](../../api/excel/excel.range)|_Property_ > hidden|Represents if all cells of the current range are hidden. Read-only.|1.2|
|[range](../../api/excel/excel.range)|_Property_ > rowHidden|Represents if all rows of the current range are hidden.|1.2|
|[range](../../api/excel/excel.range)|_Relationship_ > sort|Represents the range sort of the current range. Read-only.|1.2|
|[range](../../api/excel/excel.range)|_Method_ > merge(across: bool)|Merge the range cells into one region in the worksheet.|1.2|
|[range](../../api/excel/excel.range)|_Method_ > unmerge()|Unmerge the range cells into separate cells.|1.2|
|[rangeFormat](../../api/excel/excel.rangeformat)|_Property_ > columnWidth|Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.|1.2|
|[rangeFormat](../../api/excel/excel.rangeformat)|_Property_ > rowHeight|Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.|1.2|
|[rangeFormat](../../api/excel/excel.rangeformat)|_Relationship_ > protection|Returns the format protection object for a range. Read-only.|1.2|
|[rangeFormat](../../api/excel/excel.rangeformat)|_Method_ > autofitColumns()|Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[rangeFormat](../../api/excel/excel.rangeformat)|_Method_ > autofitRows()|Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[rangeReference](../../api/excel/excel.rangereference)|_Property_ > address|Represents the visible rows of the current range.|1.2|
|[rangeSort](../../api/excel/excel.rangesort)|_Method_ > apply(fields: SortField, matchCase: bool, hasHeaders: bool, orientation: string, method: string)|Perform a sort operation.|1.2|
|[sortField](../../api/excel/excel.sortfield)|_Property_ > ascending|Represents whether the sorting is done in an ascending fashion.|1.2|
|[sortField](../../api/excel/excel.sortfield)|_Property_ > color|Represents the color that is the target of the condition if the sorting is on font or cell color.|1.2|
|[sortField](../../api/excel/excel.sortfield)|_Property_ > dataOption|Represents additional sorting options for this field. Possible values are: Normal, TextAsNumber.|1.2|
|[sortField](../../api/excel/excel.sortfield)|_Property_ > key|Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).|1.2|
|[sortField](../../api/excel/excel.sortfield)|_Property_ > sortOn|Represents the type of sorting of this condition. Possible values are: Value, CellColor, FontColor, Icon.|1.2|
|[sortField](../../api/excel/excel.sortfield)|_Relationship_ > icon|Represents the icon that is the target of the condition if the sorting is on the cell's icon.|1.2|
|[table](../../api/excel/excel.table)|_Relationship_ > sort|Represents the sorting for the table. Read-only.|1.2|
|[table](../../api/excel/excel.table)|_Relationship_ > worksheet|The worksheet containing the current table. Read-only.|1.2|
|[table](../../api/excel/excel.table)|_Method_ > clearFilters()|Clears all the filters currently applied on the table.|1.2|
|[table](../../api/excel/excel.table)|_Method_ > convertToRange()|Converts the table into a normal range of cells. All data is preserved.|1.2|
|[table](../../api/excel/excel.table)|_Method_ > reapplyFilters()|Reapplies all the filters currently on the table.|1.2|
|[tableColumn](../../api/excel/excel.tablecolumn)|_Relationship_ > filter|Retrieve the filter applied to the column. Read-only.|1.2|
|[tableSort](../../api/excel/excel.tablesort)|_Property_ > matchCase|Represents whether the casing impacted the last sort of the table. Read-only.|1.2|
|[tableSort](../../api/excel/excel.tablesort)|_Property_ > method|Represents Chinese character ordering method last used to sort the table. Read-only. Possible values are: PinYin, StrokeCount.|1.2|
|[tableSort](../../api/excel/excel.tablesort)|_Relationship_ > fields|Represents the current conditions used to last sort the table. Read-only.|1.2|
|[tableSort](../../api/excel/excel.tablesort)|_Method_ > apply(fields: SortField, matchCase: bool, method: string)|Perform a sort operation.|1.2|
|[tableSort](../../api/excel/excel.tablesort)|_Method_ > clear()|Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.|1.2|
|[tableSort](../../api/excel/excel.tablesort)|_Method_ > reapply()|Reapplies the current sorting parameters to the table.|1.2|
|[workbook](../../api/excel/excel.workbook)|_Relationship_ > functions|Represents Excel application instance that contains this workbook. Read-only.|1.2|
|[worksheet](../../api/excel/excel.worksheet)|_Relationship_ > protection|Returns sheet protection object for a worksheet. Read-only.|1.2|
|[worksheetProtection](../../api/excel/excel.worksheetprotection)|_Property_ > protected|Indicates if the worksheet is protected. Read-Only. Read-only.|1.2|
|[worksheetProtection](../../api/excel/excel.worksheetprotection)|_Relationship_ > options|Sheet protection options. Read-only.|1.2|
|[worksheetProtection](../../api/excel/excel.worksheetprotection)|_Method_ > protect(options: WorksheetProtectionOptions)|Protects a worksheet. Fails if the worksheet has been protected.|1.2|
|[worksheetProtection](../../api/excel/excel.worksheetprotection)|_Method_ > unprotect()|Unprotects a worksheet.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowAutoFilter|Represents the worksheet protection option of allowing using auto filter feature.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowDeleteColumns|Represents the worksheet protection option of allowing deleting columns.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowDeleteRows|Represents the worksheet protection option of allowing deleting rows.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowFormatCells|Represents the worksheet protection option of allowing formatting cells.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowFormatColumns|Represents the worksheet protection option of allowing formatting columns.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowFormatRows|Represents the worksheet protection option of allowing formatting rows.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowInsertColumns|Represents the worksheet protection option of allowing inserting columns.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowInsertHyperlinks|Represents the worksheet protection option of allowing inserting hyperlinks.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowInsertRows|Represents the worksheet protection option of allowing inserting rows.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowPivotTables|Represents the worksheet protection option of allowing using PivotTable feature.|1.2|
|[worksheetProtectionOptions](../../api/excel/excel.worksheetprotectionoptions)|_Property_ > allowSort|Represents the worksheet protection option of allowing using sort feature.|1.2|

## Excel JavaScript API 1.1

Excel JavaScript API 1.1 is the first version of the API. For details about the API,  see the [Excel JavaScript API](../../api/excel) reference topics.

## See also

- [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
