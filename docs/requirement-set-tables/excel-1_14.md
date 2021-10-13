| Class | Fields | Description |
|:---|:---|:---|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|Represents the direction (such as up or to the left) that the remaining cells will shift when a cell or cells are deleted.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|Represents the direction (such as down or to the right) that the existing cells will shift when a new cell or cells are inserted.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|Gets the data table on the chart.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|Gets the data table on the chart.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Represents the format of a chart data table, which includes fill, font, and border format.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|Specifies whether to display the horizontal border of the data table.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|Specifies whether to show the legend key of the data table.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|Specifies whether to display the outline border of the data table.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|Specifies whether to display the vertical border of the data table.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Specifies whether to show the data table of the chart.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[border](/javascript/api/excel/excel.chartdatatableformat#border)|Represents the border format of chart data table, which includes color, line style, and weight.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.chartdatatableformat#font)|Represents the font attributes (such as font name, font size, and color) for the current object.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|Gets a comment from the collection based on its ID.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|Returns a comment reply identified by its ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|Returns a conditional format identified by its ID.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getItemOrNullObject_key_)|Gets a shape using its name or ID.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Gets the query error message from when the query was last refreshed.|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|Gets the query loaded to object type.|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|Specifies if the query loaded to the data model.|
||[name](/javascript/api/excel/excel.query#name)|Gets the name of the query.|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|Gets the date and time when the query was last refreshed.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|Gets the number of rows that were loaded when the query was last refreshed.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|Gets the number of queries in the workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|Gets a query from the collection based on its name.|
||[items](/javascript/api/excel/excel.querycollection#items)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|Returns a `WorkbookRangeAreas` object that represents the range containing all the precedents of a cell in the same worksheet or in multiple worksheets.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|Gets a shape using its name or ID.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|Gets a style by name.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItemOrNullObject_key_)|Gets a table by name or ID.|
|[Workbook](/javascript/api/excel/excel.workbook)|[queries](/javascript/api/excel/excel.workbook#queries)|Returns a collection of Power Query queries that are part of the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|Occurs when the worksheet protection state is changed.|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Returns a value representing this worksheet that can be read by Open Office XML.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|Represents a change to the direction that the cells in a worksheet will shift when a cell or cells are deleted or inserted.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|Represents the trigger source of the event.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|Occurs when the worksheet protection state is changed.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|Gets the current protection status of the worksheet.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|Gets the ID of the worksheet in which the protection status is changed.|
