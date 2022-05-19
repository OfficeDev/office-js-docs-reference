| Class | Fields | Description |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#excel-excel-formulachangedeventdetail-celladdress-member)|The address of the cell that contains the changed formula.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#excel-excel-formulachangedeventdetail-previousformula-member)|Represents the previous formula, before it was changed.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-positiontype-member)|The insert position, in the current workbook, of the new worksheets.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-relativeto-member)|The worksheet in the current workbook that is referenced for the `WorksheetPositionType` parameter.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-sheetnamestoinsert-member)|The names of individual worksheets to insert.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-alttextdescription-member)|The alt text description of the PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-alttexttitle-member)|The alt text title of the PivotTable.|
||[autoFormat](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-autoformat-member)|Specifies if formatting will be automatically formatted when it’s refreshed or when fields are moved.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-displayblanklineaftereachitem-member(1))|Sets whether or not to display a blank line after each item.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-emptycelltext-member)|The text that is automatically filled into any empty cell in the PivotTable if `fillEmptyCells == true`.|
||[enableFieldList](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-enablefieldlist-member)|Specifies if the field list can be shown in the UI.|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-fillemptycells-member)|Specifies whether empty cells in the PivotTable should be populated with the `emptyCellText`.|
||[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcolumnlabelrange-member(1))|Returns the range where the PivotTable's column labels reside.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatabodyrange-member(1))|Returns the range where the PivotTable's data values reside.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatahierarchy-member(1))|Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getfilteraxisrange-member(1))|Returns the range of the PivotTable's filter area.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getpivotitems-member(1))|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrange-member(1))|Returns the range the PivotTable exists on, excluding the filter area.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrowlabelrange-member(1))|Returns the range where the PivotTable's row labels reside.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-layouttype-member)|This property indicates the PivotLayoutType of all fields on the PivotTable.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-preserveformatting-member)|Specifies if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-repeatallitemlabels-member(1))|Sets the "repeat all item labels" setting across all fields in the PivotTable.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setautosortoncell-member(1))|Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showcolumngrandtotals-member)|Specifies if the PivotTable report shows grand totals for columns.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showfieldheaders-member)|Specifies whether the PivotTable displays field headers (field captions and filter drop-downs).|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showrowgrandtotals-member)|Specifies if the PivotTable report shows grand totals for rows.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-subtotallocation-member)|This property indicates the `SubtotalLocationType` of all fields on the PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-delete-member(1))|Deletes the PivotTable.|
||[refresh()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-refresh-member(1))|Refreshes the PivotTable.|
||[refreshOnOpen](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-refreshonopen-member)|Specifies whether the PivotTable refreshes when the workbook opens.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-usecustomsortlists-member)|Specifies if the PivotTable uses custom lists when sorting.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1))|Returns a `WorkbookRangeAreas` object that represents the range containing all the direct dependents of a cell in the same worksheet or in multiple worksheets.|
||[getDirectPrecedents()](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1))|Returns a `WorkbookRangeAreas` object that represents the range containing all the direct precedents of a cell in the same worksheet or in multiple worksheets.|
||[getEntireColumn()](/javascript/api/excel/excel.range#excel-excel-range-getentirecolumn-member(1))|Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#excel-excel-range-getentirerow-member(1))|Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").|
||[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getextendedrange-member(1))|Returns a range object that includes the current range and up to the edge of the range, based on the provided direction.|
||[getImage()](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1))|Renders the range as a base64-encoded png image.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getmergedareasornullobject-member(1))|Returns a `RangeAreas` object that represents the merged areas in this range.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#excel-excel-range-getoffsetrange-member(1))|Gets an object which represents a range that's offset from the specified range.|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-getpivottables-member(1))|Gets a scoped collection of PivotTables that overlap with the range.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-resize-member(1))|Resize the table to the new range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1))|Inserts the specified worksheets from a source workbook into the current workbook.|
||[onActivated](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member)|Occurs when the workbook is activated.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member)|Occurs when the AutoSave setting is changed on the workbook.|
||[onSelectionChanged](/javascript/api/excel/excel.workbook#excel-excel-workbook-onselectionchanged-member)|Occurs when the selection in the document is changed.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-save-member(1))|Save current workbook.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#excel-excel-workbookactivatedeventargs-type-member)|Gets the type of the event.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protect-member(1))|Protects a workbook.|
||[protected](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protected-member)|Specifies if the workbook is protected.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-unprotect-member(1))|Unprotects a workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[add(name?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-add-member(1))|Adds a new worksheet to the workbook.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getactiveworksheet-member(1))|Gets the currently active worksheet in the workbook.|
||[getCount(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getcount-member(1))|Gets the number of worksheets in the collection.|
||[getFirst(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getfirst-member(1))|Gets the first worksheet in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getitem-member(1))|Gets a worksheet object using its name or ID.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getitemornullobject-member(1))|Gets a worksheet object using its name or ID.|
||[items](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-items-member)|Gets the loaded child items in this collection.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|Occurs when one or more formulas are changed in this worksheet.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|Occurs when one or more formulas are changed in any worksheet of this collection.|
||[options](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-options-member)|Specifies the protection options for the worksheet.|
||[protect(options?: Excel.WorksheetProtectionOptions, password?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-protect-member(1))|Protects a worksheet.|
||[protected](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-protected-member)|Specifies if the worksheet is protected.|
||[unprotect(password?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-unprotect-member(1))|Unprotects a worksheet.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-formuladetails-member)|Gets an array of `FormulaChangedEventDetail` objects, which contain the details about the all of the changed formulas.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the formula changed.|
