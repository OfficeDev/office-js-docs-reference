| Class | Fields | Description |
|:---|:---|:---|
|[FormulaChangedEventDetail](/.formulachangedeventdetail)|[cellAddress](/.formulachangedeventdetail#excel-javascript/api/excel/-formulachangedeventdetail-celladdress-member)|The address of the cell that contains the changed formula.|
||[previousFormula](/.formulachangedeventdetail#excel-javascript/api/excel/-formulachangedeventdetail-previousformula-member)|Represents the previous formula, before it was changed.|
|[InsertWorksheetOptions](/.insertworksheetoptions)|[positionType](/.insertworksheetoptions#excel-javascript/api/excel/-insertworksheetoptions-positiontype-member)|The insert position, in the current workbook, of the new worksheets.|
||[relativeTo](/.insertworksheetoptions#excel-javascript/api/excel/-insertworksheetoptions-relativeto-member)|The worksheet in the current workbook that is referenced for the `WorksheetPositionType` parameter.|
||[sheetNamesToInsert](/.insertworksheetoptions#excel-javascript/api/excel/-insertworksheetoptions-sheetnamestoinsert-member)|The names of individual worksheets to insert.|
|[PivotLayout](/.pivotlayout)|[altTextDescription](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-alttextdescription-member)|The alt text description of the PivotTable.|
||[altTextTitle](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-alttexttitle-member)|The alt text title of the PivotTable.|
||[displayBlankLineAfterEachItem(display: boolean)](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-displayblanklineaftereachitem-member(1))|Sets whether or not to display a blank line after each item.|
||[emptyCellText](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-emptycelltext-member)|The text that is automatically filled into any empty cell in the PivotTable if `fillEmptyCells == true`.|
||[fillEmptyCells](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-fillemptycells-member)|Specifies whether empty cells in the PivotTable should be populated with the `emptyCellText`.|
||[repeatAllItemLabels(repeatLabels: boolean)](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-repeatallitemlabels-member(1))|Sets the "repeat all item labels" setting across all fields in the PivotTable.|
||[showFieldHeaders](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-showfieldheaders-member)|Specifies whether the PivotTable displays field headers (field captions and filter drop-downs).|
|[PivotTable](/.pivottable)|[refreshOnOpen](/.pivottable#excel-javascript/api/excel/-pivottable-refreshonopen-member)|Specifies whether the PivotTable refreshes when the workbook opens.|
|[Range](/.range)|[getDirectDependents()](/.range#excel-javascript/api/excel/-range-getdirectdependents-member(1))|Returns a `WorkbookRangeAreas` object that represents the range containing all the direct dependent cells of a specified range in the same worksheet or across multiple worksheets.|
||[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/.range#excel-javascript/api/excel/-range-getextendedrange-member(1))|Returns a range object that includes the current range and up to the edge of the range, based on the provided direction.|
||[getMergedAreasOrNullObject()](/.range#excel-javascript/api/excel/-range-getmergedareasornullobject-member(1))|Returns a `RangeAreas` object that represents the merged areas in this range.|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/.range#excel-javascript/api/excel/-range-getrangeedge-member(1))|Returns a range object that is the edge cell of the data region that corresponds to the provided direction.|
|[Table](/.table)|[resize(newRange: Range \| string)](/.table#excel-javascript/api/excel/-table-resize-member(1))|Resize the table to the new range.|
|[Workbook](/.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions)](/.workbook#excel-javascript/api/excel/-workbook-insertworksheetsfrombase64-member(1))|Inserts the specified worksheets from a source workbook into the current workbook.|
||[onActivated](/.workbook#excel-javascript/api/excel/-workbook-onactivated-member)|Occurs when the workbook is activated.|
|[WorkbookActivatedEventArgs](/.workbookactivatedeventargs)|[type](/.workbookactivatedeventargs#excel-javascript/api/excel/-workbookactivatedeventargs-type-member)|Gets the type of the event.|
|[Worksheet](/.worksheet)|[onFormulaChanged](/.worksheet#excel-javascript/api/excel/-worksheet-onformulachanged-member)|Occurs when one or more formulas are changed in this worksheet.|
|[WorksheetCollection](/.worksheetcollection)|[onFormulaChanged](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onformulachanged-member)|Occurs when one or more formulas are changed in any worksheet of this collection.|
|[WorksheetFormulaChangedEventArgs](/.worksheetformulachangedeventargs)|[formulaDetails](/.worksheetformulachangedeventargs#excel-javascript/api/excel/-worksheetformulachangedeventargs-formuladetails-member)|Gets an array of `FormulaChangedEventDetail` objects, which contain the details about the all of the changed formulas.|
||[source](/.worksheetformulachangedeventargs#excel-javascript/api/excel/-worksheetformulachangedeventargs-source-member)|The source of the event.|
||[type](/.worksheetformulachangedeventargs#excel-javascript/api/excel/-worksheetformulachangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetformulachangedeventargs#excel-javascript/api/excel/-worksheetformulachangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the formula changed.|
