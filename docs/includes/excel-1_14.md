| Class | Fields | Description |
|:---|:---|:---|
|[AutoFilter](/.autofilter)|[clearColumnCriteria(columnIndex: number)](/.autofilter#excel-javascript/api/excel/-autofilter-clearcolumncriteria-member(1))|Clears the column filter criteria of the AutoFilter.|
|[ChangeDirectionState](/.changedirectionstate)|[deleteShiftDirection](/.changedirectionstate#excel-javascript/api/excel/-changedirectionstate-deleteshiftdirection-member)|Represents the direction (such as up or to the left) that the remaining cells will shift when a cell or cells are deleted.|
||[insertShiftDirection](/.changedirectionstate#excel-javascript/api/excel/-changedirectionstate-insertshiftdirection-member)|Represents the direction (such as down or to the right) that the existing cells will shift when a new cell or cells are inserted.|
|[Chart](/.chart)|[getDataTable()](/.chart#excel-javascript/api/excel/-chart-getdatatable-member(1))|Gets the data table on the chart.|
||[getDataTableOrNullObject()](/.chart#excel-javascript/api/excel/-chart-getdatatableornullobject-member(1))|Gets the data table on the chart.|
|[ChartDataTable](/.chartdatatable)|[format](/.chartdatatable#excel-javascript/api/excel/-chartdatatable-format-member)|Represents the format of a chart data table, which includes fill, font, and border format.|
||[showHorizontalBorder](/.chartdatatable#excel-javascript/api/excel/-chartdatatable-showhorizontalborder-member)|Specifies whether to display the horizontal border of the data table.|
||[showLegendKey](/.chartdatatable#excel-javascript/api/excel/-chartdatatable-showlegendkey-member)|Specifies whether to show the legend key of the data table.|
||[showOutlineBorder](/.chartdatatable#excel-javascript/api/excel/-chartdatatable-showoutlineborder-member)|Specifies whether to display the outline border of the data table.|
||[showVerticalBorder](/.chartdatatable#excel-javascript/api/excel/-chartdatatable-showverticalborder-member)|Specifies whether to display the vertical border of the data table.|
||[visible](/.chartdatatable#excel-javascript/api/excel/-chartdatatable-visible-member)|Specifies whether to show the data table of the chart.|
|[ChartDataTableFormat](/.chartdatatableformat)|[border](/.chartdatatableformat#excel-javascript/api/excel/-chartdatatableformat-border-member)|Represents the border format of chart data table, which includes color, line style, and weight.|
||[fill](/.chartdatatableformat#excel-javascript/api/excel/-chartdatatableformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/.chartdatatableformat#excel-javascript/api/excel/-chartdatatableformat-font-member)|Represents the font attributes (such as font name, font size, and color) for the current object.|
|[CommentCollection](/.commentcollection)|[getItemOrNullObject(commentId: string)](/.commentcollection#excel-javascript/api/excel/-commentcollection-getitemornullobject-member(1))|Gets a comment from the collection based on its ID.|
|[CommentReplyCollection](/.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/.commentreplycollection#excel-javascript/api/excel/-commentreplycollection-getitemornullobject-member(1))|Returns a comment reply identified by its ID.|
|[ConditionalFormatCollection](/.conditionalformatcollection)|[getItemOrNullObject(id: string)](/.conditionalformatcollection#excel-javascript/api/excel/-conditionalformatcollection-getitemornullobject-member(1))|Returns a conditional format identified by its ID.|
|[GroupShapeCollection](/.groupshapecollection)|[getItemOrNullObject(key: string)](/.groupshapecollection#excel-javascript/api/excel/-groupshapecollection-getitemornullobject-member(1))|Gets a shape using its name or ID.|
|[Query](/.query)|[error](/.query#excel-javascript/api/excel/-query-error-member)|Gets the query error message from when the query was last refreshed.|
||[loadedTo](/.query#excel-javascript/api/excel/-query-loadedto-member)|Gets the query loaded to object type.|
||[loadedToDataModel](/.query#excel-javascript/api/excel/-query-loadedtodatamodel-member)|Specifies if the query loaded to the data model.|
||[name](/.query#excel-javascript/api/excel/-query-name-member)|Gets the name of the query.|
||[refreshDate](/.query#excel-javascript/api/excel/-query-refreshdate-member)|Gets the date and time when the query was last refreshed.|
||[rowsLoadedCount](/.query#excel-javascript/api/excel/-query-rowsloadedcount-member)|Gets the number of rows that were loaded when the query was last refreshed.|
|[QueryCollection](/.querycollection)|[getCount()](/.querycollection#excel-javascript/api/excel/-querycollection-getcount-member(1))|Gets the number of queries in the workbook.|
||[getItem(key: string)](/.querycollection#excel-javascript/api/excel/-querycollection-getitem-member(1))|Gets a query from the collection based on its name.|
||[items](/.querycollection#excel-javascript/api/excel/-querycollection-items-member)|Gets the loaded child items in this collection.|
|[Range](/.range)|[getPrecedents()](/.range#excel-javascript/api/excel/-range-getprecedents-member(1))|Returns a `WorkbookRangeAreas` object that represents the range containing all the precedent cells of a specified range in the same worksheet or across multiple worksheets.|
|[ShapeCollection](/.shapecollection)|[getItemOrNullObject(key: string)](/.shapecollection#excel-javascript/api/excel/-shapecollection-getitemornullobject-member(1))|Gets a shape using its name or ID.|
|[StyleCollection](/.stylecollection)|[getItemOrNullObject(name: string)](/.stylecollection#excel-javascript/api/excel/-stylecollection-getitemornullobject-member(1))|Gets a style by name.|
|[TableScopedCollection](/.tablescopedcollection)|[getItemOrNullObject(key: string)](/.tablescopedcollection#excel-javascript/api/excel/-tablescopedcollection-getitemornullobject-member(1))|Gets a table by name or ID.|
|[Workbook](/.workbook)|[queries](/.workbook#excel-javascript/api/excel/-workbook-queries-member)|Returns a collection of Power Query queries that are part of the workbook.|
|[Worksheet](/.worksheet)|[onProtectionChanged](/.worksheet#excel-javascript/api/excel/-worksheet-onprotectionchanged-member)|Occurs when the worksheet protection state is changed.|
||[tabId](/.worksheet#excel-javascript/api/excel/-worksheet-tabid-member)|Returns a value representing this worksheet that can be read by Open Office XML.|
|[WorksheetChangedEventArgs](/.worksheetchangedeventargs)|[changeDirectionState](/.worksheetchangedeventargs#excel-javascript/api/excel/-worksheetchangedeventargs-changedirectionstate-member)|Represents a change to the direction that the cells in a worksheet will shift when a cell or cells are deleted or inserted.|
||[triggerSource](/.worksheetchangedeventargs#excel-javascript/api/excel/-worksheetchangedeventargs-triggersource-member)|Represents the trigger source of the event.|
|[WorksheetCollection](/.worksheetcollection)|[onProtectionChanged](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onprotectionchanged-member)|Occurs when the worksheet protection state is changed.|
|[WorksheetProtectionChangedEventArgs](/.worksheetprotectionchangedeventargs)|[isProtected](/.worksheetprotectionchangedeventargs#excel-javascript/api/excel/-worksheetprotectionchangedeventargs-isprotected-member)|Gets the current protection status of the worksheet.|
||[source](/.worksheetprotectionchangedeventargs#excel-javascript/api/excel/-worksheetprotectionchangedeventargs-source-member)|The source of the event.|
||[type](/.worksheetprotectionchangedeventargs#excel-javascript/api/excel/-worksheetprotectionchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetprotectionchangedeventargs#excel-javascript/api/excel/-worksheetprotectionchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the protection status is changed.|
