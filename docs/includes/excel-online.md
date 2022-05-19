| Class | Fields | Description |
|:---|:---|:---|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-breaklinks-member(1))|Makes a request to break the links pointing to the linked workbook.|
||[id](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-id-member)|The original URL pointing to the linked workbook.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-refresh-member(1))|Makes a request to refresh the data retrieved from the linked workbook.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-breakalllinks-member(1))|Breaks all the links to the linked workbooks.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitem-member(1))|Gets information about a linked workbook by its URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitemornullobject-member(1))|Gets information about a linked workbook by its URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-items-member)|Gets the loaded child items in this collection.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-refreshall-member(1))|Makes a request to refresh all the workbook links.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-workbooklinksrefreshmode-member)|Represents the update mode of the workbook links.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-activate-member(1))|Activates this sheet view.|
||[delete()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-delete-member(1))|Removes the sheet view from the worksheet.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-duplicate-member(1))|Creates a copy of this sheet view.|
||[name](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-name-member)|Gets or sets the name of the sheet view.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-add-member(1))|Creates a new sheet view with the given name.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-entertemporary-member(1))|Creates and activates a new temporary sheet view.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-exit-member(1))|Exits the currently active sheet view.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getactive-member(1))|Gets the worksheet's currently active sheet view.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getcount-member(1))|Gets the number of sheet views in this worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitem-member(1))|Gets a sheet view using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemat-member(1))|Gets a sheet view by its index in the collection.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-items-member)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-delete-member(1))|Deletes the row from the table.|
||[getRange()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-getrange-member(1))|Returns the range object associated with the entire row.|
||[index](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-index-member)|Returns the index number of the row within the rows collection of the table.|
||[values](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-values-member)|Represents the raw values of the specified range.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterows-member(1))|Delete multiple rows from a table.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterowsat-member(1))|Delete a specified number of rows from a table, starting at a given index.|
||[getCount()](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getcount-member(1))|Gets the number of rows in the table.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getitemat-member(1))|Gets a row based on its position in the collection.|
|[Workbook](/javascript/api/excel/excel.workbook)|[autoSave](/javascript/api/excel/excel.workbook#excel-excel-workbook-autosave-member)|Specifies if the workbook is in AutoSave mode.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#excel-excel-workbook-calculationengineversion-member)|Returns a number about the version of Excel Calculation Engine.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbook#excel-excel-workbook-chartdatapointtrack-member)|True if all charts in the workbook are tracking the actual data points to which they are attached.|
||[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-close-member(1))|Close current workbook.|
||[getActiveCell()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivecell-member(1))|Gets the currently active cell from the workbook.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechart-member(1))|Gets the currently active chart in the workbook.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechartornullobject-member(1))|Gets the currently active chart in the workbook.|
||[isDirty](/javascript/api/excel/excel.workbook#excel-excel-workbook-isdirty-member)|Specifies if changes have been made since the workbook was last saved.|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkedworkbooks-member)|Returns a collection of linked workbooks.|
||[name](/javascript/api/excel/excel.workbook#excel-excel-workbook-name-member)|Gets the workbook name.|
||[names](/javascript/api/excel/excel.workbook#excel-excel-workbook-names-member)|Represents a collection of workbook-scoped named items (named ranges and constants).|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottablestyles-member)|Represents a collection of PivotTableStyles associated with the workbook.|
||[pivotTables](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottables-member)|Represents a collection of PivotTables associated with the workbook.|
||[previouslySaved](/javascript/api/excel/excel.workbook#excel-excel-workbook-previouslysaved-member)|Specifies if the workbook has ever been saved locally or online.|
||[properties](/javascript/api/excel/excel.workbook#excel-excel-workbook-properties-member)|Gets the workbook properties.|
||[protection](/javascript/api/excel/excel.workbook#excel-excel-workbook-protection-member)|Returns the protection object for a workbook.|
||[queries](/javascript/api/excel/excel.workbook#excel-excel-workbook-queries-member)|Returns a collection of Power Query queries that are part of the workbook.|
||[readOnly](/javascript/api/excel/excel.workbook#excel-excel-workbook-readonly-member)|Returns `true` if the workbook is open in read-only mode.|
||[settings](/javascript/api/excel/excel.workbook#excel-excel-workbook-settings-member)|Represents a collection of settings associated with the workbook.|
||[slicerStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicerstyles-member)|Represents a collection of SlicerStyles associated with the workbook.|
||[slicers](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicers-member)|Represents a collection of slicers associated with the workbook.|
||[styles](/javascript/api/excel/excel.workbook#excel-excel-workbook-styles-member)|Represents a collection of styles associated with the workbook.|
||[tableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-tablestyles-member)|Represents a collection of TableStyles associated with the workbook.|
||[tables](/javascript/api/excel/excel.workbook#excel-excel-workbook-tables-member)|Represents a collection of tables associated with the workbook.|
||[timelineStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-timelinestyles-member)|Represents a collection of TimelineStyles associated with the workbook.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#excel-excel-workbook-useprecisionasdisplayed-member)|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|
||[worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member)|Represents a collection of worksheets associated with the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-activate-member(1))|Activate the worksheet in the Excel UI.|
||[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-calculate-member(1))|Calculates all cells on a worksheet.|
||[copy(positionType?: Excel.WorksheetPositionType, relativeTo?: Excel.Worksheet)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-copy-member(1))|Copies a worksheet and places it at the specified position.|
||[delete()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-delete-member(1))|Deletes the worksheet from the workbook.|
||[enableCalculation](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-enablecalculation-member)|Determines if Excel should recalculate the worksheet when necessary.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findall-member(1))|Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findallornullobject-member(1))|Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.|
||[id](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-id-member)|Returns a value that uniquely identifies the worksheet in a given workbook.|
||[name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member)|The display name of the worksheet.|
||[namedSheetViews](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-namedsheetviews-member)|Returns a collection of sheet views that are present in the worksheet.|
||[names](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-names-member)|Collection of names scoped to the current worksheet.|
||[onNameChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onnamechanged-member)|Occurs when the worksheet name is changed.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member)|Occurs when the worksheet protection state is changed.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member)|Occurs when one or more rows have been sorted.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member)|Occurs when the selection changes on a specific worksheet.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member)|Occurs when a left-clicked/tapped action happens in the worksheet.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onvisibilitychanged-member)|Occurs when the worksheet visibility is changed.|
||[pageLayout](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pagelayout-member)|Gets the `PageLayout` object of the worksheet.|
||[pivotTables](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pivottables-member)|Collection of PivotTables that are part of the worksheet.|
||[position](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-position-member)|The zero-based position of the worksheet within the workbook.|
||[protection](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-protection-member)|Returns the sheet protection object for a worksheet.|
||[shapes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-shapes-member)|Returns the collection of all the Shape objects on the worksheet.|
||[showGridlines](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showgridlines-member)|Specifies if gridlines are visible to the user.|
||[showHeadings](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showheadings-member)|Specifies if headings are visible to the user.|
||[slicers](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-slicers-member)|Returns a collection of slicers that are part of the worksheet.|
||[standardHeight](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardheight-member)|Returns the standard (default) height of all the rows in the worksheet, in points.|
||[standardWidth](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardwidth-member)|Specifies the standard (default) width of all the columns in the worksheet.|
||[tabColor](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabcolor-member)|The tab color of the worksheet.|
||[tabId](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabid-member)|Returns a value representing this worksheet that can be read by Open Office XML.|
||[tables](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tables-member)|Collection of tables that are part of the worksheet.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-verticalpagebreaks-member)|Gets the vertical page break collection for the worksheet.|
||[visibility](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-visibility-member)|The visibility of the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-add-member(1))|Adds a new worksheet to the workbook.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getactiveworksheet-member(1))|Gets the currently active worksheet in the workbook.|
||[getCount(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getcount-member(1))|Gets the number of worksheets in the collection.|
||[getFirst(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getfirst-member(1))|Gets the first worksheet in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitem-member(1))|Gets a worksheet object using its name or ID.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitemornullobject-member(1))|Gets a worksheet object using its name or ID.|
||[getLast(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getlast-member(1))|Gets the last worksheet in the collection.|
||[items](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-items-member)|Gets the loaded child items in this collection.|
||[onActivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member)|Occurs when any worksheet in the workbook is activated.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member)|Occurs when a new worksheet is added to the workbook.|
||[onCalculated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member)|Occurs when any worksheet in the workbook is calculated.|
||[onChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member)|Occurs when any worksheet in the workbook is changed.|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member)|Occurs when one or more columns have been sorted.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member)|Occurs when any worksheet in the workbook is deactivated.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member)|Occurs when a worksheet is deleted from the workbook.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member)|Occurs when any worksheet in the workbook has a format changed.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member)|Occurs when one or more formulas are changed in any worksheet of this collection.|
||[onMoved](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onmoved-member)|Occurs when a worksheet is moved within a workbook.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onnamechanged-member)|Occurs when the worksheet name is changed in the worksheet collection.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member)|Occurs when the worksheet protection state is changed.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member)|Occurs when one or more rows have been sorted.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member)|Occurs when the selection changes on any worksheet.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member)|Occurs when left-clicked/tapped operation happens in the worksheet collection.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onvisibilitychanged-member)|Occurs when the worksheet visibility is changed in the worksheet collection.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionafter-member)|Gets the new position of the worksheet, after the move.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionbefore-member)|Gets the previous position of the worksheet, prior to the move.|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-worksheetid-member)|Gets the ID of the worksheet that was moved.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-nameafter-member)|Gets the new name of the worksheet, after the name change.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-namebefore-member)|Gets the previous name of the worksheet, before the name changed.|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-worksheetid-member)|Gets the ID of the worksheet with the new name.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[options](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-options-member)|Specifies the protection options for the worksheet.|
||[protect(options?: Excel.WorksheetProtectionOptions, password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-protect-member(1))|Protects a worksheet.|
||[protected](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-protected-member)|Specifies if the worksheet is protected.|
||[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-unprotect-member(1))|Unprotects a worksheet.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-type-member)|Gets the type of the event.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilityafter-member)|Gets the new visibility setting of the worksheet, after the visibility change.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilitybefore-member)|Gets the previous visibility setting of the worksheet, before the visibility change.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-worksheetid-member)|Gets the ID of the worksheet whose visibility has changed.|
