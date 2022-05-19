| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel.CalculationType)](/javascript/api/excel/excel.application#excel-excel-application-calculate-member(1))|Recalculate all currently opened workbooks in Excel.|
||[calculationEngineVersion](/javascript/api/excel/excel.application#excel-excel-application-calculationengineversion-member)|Returns the Excel calculation engine version used for the last full recalculation.|
||[calculationMode](/javascript/api/excel/excel.application#excel-excel-application-calculationmode-member)|Returns the calculation mode used in the workbook, as defined by the constants in `Excel.CalculationMode`.|
||[calculationState](/javascript/api/excel/excel.application#excel-excel-application-calculationstate-member)|Returns the calculation state of the application.|
||[cultureInfo](/javascript/api/excel/excel.application#excel-excel-application-cultureinfo-member)|Provides information based on current system culture settings.|
||[decimalSeparator](/javascript/api/excel/excel.application#excel-excel-application-decimalseparator-member)|Gets the string used as the decimal separator for numeric values.|
||[iterativeCalculation](/javascript/api/excel/excel.application#excel-excel-application-iterativecalculation-member)|Returns the iterative calculation settings.|
||[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendapicalculationuntilnextsync-member(1))|Suspends calculation until the next `context.sync()` is called.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendscreenupdatinguntilnextsync-member(1))|Suspends screen updating until the next `context.sync()` is called.|
||[thousandsSeparator](/javascript/api/excel/excel.application#excel-excel-application-thousandsseparator-member)|Gets the string used to separate groups of digits to the left of the decimal for numeric values.|
||[useSystemSeparators](/javascript/api/excel/excel.application#excel-excel-application-usesystemseparators-member)|Specifies if the system separators of Excel are enabled.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|Creates a new comment with the given content on the given cell.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-email-member)|The email address of the entity that is mentioned in a comment.|
||[id](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-id-member)|The ID of the entity.|
||[name](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-name-member)|The name of the entity that is mentioned in a comment.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|Creates a comment reply for a comment.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-mentions-member)|An array containing all the entities (e.g., people) mentioned within the comment.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-richcontent-member)|Specifies the rich content of the comment (e.g., comment content with mentions, the first mentioned entity has an ID attribute of 0, and the second mentioned entity has an ID attribute of 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|||
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numberdecimalseparator-member)|Gets the string used as the decimal separator for numeric values.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numbergroupseparator-member)|Gets the string used to separate groups of digits to the left of the decimal for numeric values.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-moveto-member(1))|Moves cell values, formatting, and formulas from current range to the destination range, replacing the old information in those cells.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1))|Removes duplicate values from the range specified by the columns.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#excel-excel-range-replaceall-member(1))|Finds and replaces the given string based on the criteria specified within the current range.|
||[select()](/javascript/api/excel/excel.range#excel-excel-range-select-member(1))|Selects the specified range in the Excel UI.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#excel-excel-range-setcellproperties-member(1))|Updates the range based on a 2D array of cell properties, encapsulating things like font, fill, borders, and alignment.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setcolumnproperties-member(1))|Updates the range based on a single-dimensional array of column properties, encapsulating things like font, fill, borders, and alignment.|
||[setDirty()](/javascript/api/excel/excel.range#excel-excel-range-setdirty-member(1))|Set a range to be recalculated when the next recalculation occurs.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setrowproperties-member(1))|Updates the range based on a single-dimensional array of row properties, encapsulating things like font, fill, borders, and alignment.|
||[showCard()](/javascript/api/excel/excel.range#excel-excel-range-showcard-member(1))|Displays the card for an active cell if it has rich value content.|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-showgroupdetails-member(1))|Shows the details of the row or column group.|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-ungroup-member(1))|Ungroups columns and rows for an outline.|
||[unmerge()](/javascript/api/excel/excel.range#excel-excel-range-unmerge-member(1))|Unmerge the range cells into separate cells.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-adjustindent-member(1))|Adjusts the indentation of the range formatting.|
||[autofitColumns()](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-autofitcolumns-member(1))|Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.|
||[autofitRows()](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-autofitrows-member(1))|Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-source-member)|Gets the source of the event.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-tableid-member)|Gets the ID of the table that is added.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the table is added.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-address-member)|The address of the range that completed calculation.|
||[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-type-member)|Gets the type of the event.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-address-member)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-changetype-member)|Gets the type of change that represents how the event was triggered.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the data changed.|
