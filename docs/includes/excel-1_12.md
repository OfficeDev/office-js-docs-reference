| Class | Fields | Description |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-setformula-member(1))|A string value that represents the formula of chart axis title using A1-style notation.|
||[textOrientation](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-textorientation-member)|Specifies the angle to which the text is oriented for the chart axis title.|
||[visible](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-visible-member)|Specifies if the axis title is visibile.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-getdimensionvalues-member(1))|Gets the values from a single dimension of the chart series.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setbubblesizes-member(1))|Sets the bubble sizes for a chart series.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setvalues-member(1))|Sets the values for a chart series.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setxaxisvalues-member(1))|Sets the values of the x-axis for a chart series.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-commentdetails-member)|Gets the `CommentDetail` array that contains the comment ID and IDs of its related replies.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-source-member)|Specifies the source of the event.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the event happened.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-changetype-member)|Gets the change type that represents how the changed event is triggered.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-commentdetails-member)|Get the `CommentDetail` array which contains the comment ID and IDs of its related replies.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-source-member)|Specifies the source of the event.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the event happened.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[contentType](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-contenttype-member)|Gets the content type of the comment.|
||[creationDate](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-creationdate-member)|Gets the creation time of the comment.|
||[delete()](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-delete-member(1))|Deletes the comment and all the connected replies.|
||[getLocation()](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getlocation-member(1))|Gets the cell where this comment is located.|
||[id](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-id-member)|Specifies the comment identifier.|
||[mentions](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-mentions-member)|Gets the entities (e.g., people) that are mentioned in comments.|
||[resolved](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-resolved-member)|The comment thread status.|
||[richContent](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-richcontent-member)|Gets the rich comment content (e.g., mentions in comments).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-updatementions-member(1))|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-commentdetails-member)|Gets the `CommentDetail` array that contains the comment ID and IDs of its related replies.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-source-member)|Specifies the source of the event.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the event happened.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#excel-excel-commentdetail-commentid-member)|Represents the ID of the comment.|
||[replyIds](/javascript/api/excel/excel.commentdetail#excel-excel-commentdetail-replyids-member)|Represents the IDs of the related replies that belong to the comment.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[contentType](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-contenttype-member)|The content type of the reply.|
||[creationDate](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-creationdate-member)|Gets the creation time of the comment reply.|
||[delete()](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-delete-member(1))|Deletes the comment reply.|
||[getLocation()](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getlocation-member(1))|Gets the cell where this comment reply is located.|
||[getParentComment()](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getparentcomment-member(1))|Gets the parent comment of this reply.|
||[id](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-id-member)|Specifies the comment reply identifier.|
||[mentions](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-mentions-member)|The entities (e.g., people) that are mentioned in comments.|
||[resolved](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-resolved-member)|The comment reply status.|
||[richContent](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-richcontent-member)|The rich comment content (e.g., mentions in comments).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-updatementions-member(1))|Updates the comment content with a specially formatted string and a list of mentions.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-datetimeformat-member)|Defines the culturally appropriate format of displaying date and time.|
||[name](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-name-member)|Gets the culture name in the format languagecode2-country/regioncode2 (e.g., "zh-cn" or "en-us").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-numberformat-member)|Defines the culturally appropriate format of displaying numbers.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-dateseparator-member)|Gets the string used as the date separator.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-longdatepattern-member)|Gets the format string for a long date value.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-longtimepattern-member)|Gets the format string for a long time value.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-shortdatepattern-member)|Gets the format string for a short date value.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-timeseparator-member)|Gets the string used as the time separator.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparator](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-comparator-member)|The comparator is the static value to which other values are compared.|
||[condition](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-condition-member)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-exclusive-member)|If `true`, filter *excludes* items that meet criteria.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-lowerbound-member)|The lower-bound of the range for the `between` filter condition.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-upperbound-member)|The upper-bound of the range for the `between` filter condition.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-wholedays-member)|For `equals`, `before`, `after`, and `between` filter conditions, indicates if comparisons should be made as whole days.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-applyfilter-member(1))|Sets one or more of the field's current PivotFilters and applies them to the field.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-clearallfilters-member(1))|Clears all criteria from all of the field's filters.|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-clearfilter-member(1))|Clears all existing criteria from the field's filter of the given type (if one is currently applied).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-getfilters-member(1))|Gets all filters currently applied on the field.|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-isfiltered-member(1))|Checks if there are any applied filters on the field.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbylabels-member(1))|Sorts the PivotField.|
||[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbyvalues-member(1))|Sorts the PivotField by specified values in a given scope.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-datefilter-member)|The PivotField's currently applied date filter.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-labelfilter-member)|The PivotField's currently applied label filter.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-manualfilter-member)|The PivotField's currently applied manual filter.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-valuefilter-member)|The PivotField's currently applied value filter.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparator](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-comparator-member)|The comparator is the static value to which other values are compared.|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-condition-member)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-exclusive-member)|If `true`, filter *excludes* items that meet criteria.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-lowerbound-member)|The lower-bound of the range for the `between` filter condition.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-substring-member)|The substring used for `beginsWith`, `endsWith`, and `contains` filter conditions.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-upperbound-member)|The upper-bound of the range for the `between` filter condition.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|||
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#excel-excel-pivotmanualfilter-selecteditems-member)|A list of selected items to manually filter.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-allowmultiplefiltersperfield-member)|Specifies if the PivotTable allows the application of multiple PivotFilters on a given PivotField in the table.|
||[enableDataValueEditing](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-enabledatavalueediting-member)|Specifies if the PivotTable allows values in the data body to be edited by the user.|
||[id](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-id-member)|ID of the PivotTable.|
||[name](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-name-member)|Name of the PivotTable.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getcount-member(1))|Gets the number of PivotTables in the collection.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getfirst-member(1))|Gets the first PivotTable in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getitem-member(1))|Gets a PivotTable by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getitemornullobject-member(1))|Gets a PivotTable by name.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-items-member)|Gets the loaded child items in this collection.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparator](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-comparator-member)|The comparator is the static value to which other values are compared.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-condition-member)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-exclusive-member)|If `true`, filter *excludes* items that meet criteria.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-lowerbound-member)|The lower-bound of the range for the `between` filter condition.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-selectiontype-member)|Specifies if the filter is for the top/bottom N items, top/bottom N percent, or top/bottom N sum.|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-threshold-member)|The "N" threshold number of items, percent, or sum to be filtered for a top/bottom filter condition.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-upperbound-member)|The upper-bound of the range for the `between` filter condition.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-value-member)|Name of the chosen "value" in the field by which to filter.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#excel-excel-range-autofill-member(1))|Fills a range from the current range to the destination range using the specified AutoFill logic.|
||[getSpillParent()](/javascript/api/excel/excel.range#excel-excel-range-getspillparent-member(1))|Gets the range object containing the anchor cell for a cell getting spilled into.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getspillparentornullobject-member(1))|Gets the range object containing the anchor cell for the cell getting spilled into.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorange-member(1))|Gets the range object containing the spill range when called on an anchor cell.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorangeornullobject-member(1))|Gets the range object containing the spill range when called on an anchor cell.|
||[hasSpill](/javascript/api/excel/excel.range#excel-excel-range-hasspill-member)|Represents if all cells have a spill border.|
||[height](/javascript/api/excel/excel.range#excel-excel-range-height-member)|Returns the distance in points, for 100% zoom, from the top edge of the range to the bottom edge of the range.|
||[hidden](/javascript/api/excel/excel.range#excel-excel-range-hidden-member)|Represents if all cells in the current range are hidden.|
||[hyperlink](/javascript/api/excel/excel.range#excel-excel-range-hyperlink-member)|Represents the hyperlink for the current range.|
||[isEntireColumn](/javascript/api/excel/excel.range#excel-excel-range-isentirecolumn-member)|Represents if the current range is an entire column.|
||[isEntireRow](/javascript/api/excel/excel.range#excel-excel-range-isentirerow-member)|Represents if the current range is an entire row.|
||[left](/javascript/api/excel/excel.range#excel-excel-range-left-member)|Returns the distance in points, for 100% zoom, from the left edge of the worksheet to the left edge of the range.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#excel-excel-range-linkeddatatypestate-member)|Represents the data type state of each cell.|
||[numberFormat](/javascript/api/excel/excel.range#excel-excel-range-numberformat-member)|Represents Excel's number format code for the given range.|
||[numberFormatCategories](/javascript/api/excel/excel.range#excel-excel-range-numberformatcategories-member)|Represents the category of number format of each cell.|
||[numberFormatLocal](/javascript/api/excel/excel.range#excel-excel-range-numberformatlocal-member)|Represents Excel's number format code for the given range, based on the language settings of the user.|
||[rowCount](/javascript/api/excel/excel.range#excel-excel-range-rowcount-member)|Returns the total number of rows in the range.|
||[rowHidden](/javascript/api/excel/excel.range#excel-excel-range-rowhidden-member)|Represents if all rows in the current range are hidden.|
||[rowIndex](/javascript/api/excel/excel.range#excel-excel-range-rowindex-member)|Returns the row number of the first cell in the range.|
||[savedAsArray](/javascript/api/excel/excel.range#excel-excel-range-savedasarray-member)|Represents if all the cells would be saved as an array formula.|
||[style](/javascript/api/excel/excel.range#excel-excel-range-style-member)|Represents the style of the current range.|
||[text](/javascript/api/excel/excel.range#excel-excel-range-text-member)|Text values of the specified range.|
||[top](/javascript/api/excel/excel.range#excel-excel-range-top-member)|Returns the distance in points, for 100% zoom, from the top edge of the worksheet to the top edge of the range.|
||[valueTypes](/javascript/api/excel/excel.range#excel-excel-range-valuetypes-member)|Specifies the type of data in each cell.|
||[values](/javascript/api/excel/excel.range#excel-excel-range-values-member)|Represents the raw values of the specified range.|
||[width](/javascript/api/excel/excel.range#excel-excel-range-width-member)|Returns the distance in points, for 100% zoom, from the left edge of the range to the right edge of the range.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-getcount-member(1))|Gets the number of `RangeAreas` objects in this collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-getitemat-member(1))|Returns the `RangeAreas` object based on position in the collection.|
||[items](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-items-member)|Gets the loaded child items in this collection.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[addresses](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-addresses-member)|Returns an array of addresses in A1-style.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-areas-member)|Returns the `RangeAreasCollection` object.|
||[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-getrangeareasbysheet-member(1))|Returns the `RangeAreas` object based on worksheet ID or name in the collection.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-getrangeareasornullobjectbysheet-member(1))|Returns the `RangeAreas` object based on worksheet name or ID in the collection.|
||[ranges](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-ranges-member)|Returns ranges that comprise this object in a `RangeCollection` object.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-customproperties-member)|Gets a collection of worksheet-level custom properties.|
||[freezePanes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-freezepanes-member)|Gets an object that can be used to manipulate frozen panes on the worksheet.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getcell-member(1))|Gets the `Range` object containing the single cell based on row and column numbers.|
||[getNext(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getnext-member(1))|Gets the worksheet that follows this one.|
||[getNextOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getnextornullobject-member(1))|Gets the worksheet that follows this one.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-horizontalpagebreaks-member)|Gets the horizontal page break collection for the worksheet.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-delete-member(1))|Deletes the custom property.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-key-member)|Gets the key of the custom property.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-value-member)|Gets or sets the value of the custom property.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-add-member(1))|Adds a new custom property that maps to the provided key.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getcount-member(1))|Gets the number of custom properties on this worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getitem-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getitemornullobject-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-items-member)|Gets the loaded child items in this collection.|
