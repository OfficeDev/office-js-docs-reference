| Class | Fields | Description |
|:---|:---|:---|
|[ChartAxisTitle](/.chartaxistitle)|[textOrientation](/.chartaxistitle#excel-javascript/api/excel/-chartaxistitle-textorientation-member)|Specifies the angle to which the text is oriented for the chart axis title.|
|[ChartSeries](/.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/.chartseries#excel-javascript/api/excel/-chartseries-getdimensionvalues-member(1))|Gets the values from a single dimension of the chart series.|
|[Comment](/.comment)|[contentType](/.comment#excel-javascript/api/excel/-comment-contenttype-member)|Gets the content type of the comment.|
|[CommentAddedEventArgs](/.commentaddedeventargs)|[commentDetails](/.commentaddedeventargs#excel-javascript/api/excel/-commentaddedeventargs-commentdetails-member)|Gets the `CommentDetail` array that contains the comment ID and IDs of its related replies.|
||[source](/.commentaddedeventargs#excel-javascript/api/excel/-commentaddedeventargs-source-member)|Specifies the source of the event.|
||[type](/.commentaddedeventargs#excel-javascript/api/excel/-commentaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.commentaddedeventargs#excel-javascript/api/excel/-commentaddedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the event happened.|
|[CommentChangedEventArgs](/.commentchangedeventargs)|[changeType](/.commentchangedeventargs#excel-javascript/api/excel/-commentchangedeventargs-changetype-member)|Gets the change type that represents how the changed event is triggered.|
||[commentDetails](/.commentchangedeventargs#excel-javascript/api/excel/-commentchangedeventargs-commentdetails-member)|Get the `CommentDetail` array which contains the comment ID and IDs of its related replies.|
||[source](/.commentchangedeventargs#excel-javascript/api/excel/-commentchangedeventargs-source-member)|Specifies the source of the event.|
||[type](/.commentchangedeventargs#excel-javascript/api/excel/-commentchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.commentchangedeventargs#excel-javascript/api/excel/-commentchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the event happened.|
|[CommentCollection](/.commentcollection)|[onAdded](/.commentcollection#excel-javascript/api/excel/-commentcollection-onadded-member)|Occurs when the comments are added.|
||[onChanged](/.commentcollection#excel-javascript/api/excel/-commentcollection-onchanged-member)|Occurs when comments or replies in a comment collection are changed, including when replies are deleted.|
||[onDeleted](/.commentcollection#excel-javascript/api/excel/-commentcollection-ondeleted-member)|Occurs when comments are deleted in the comment collection.|
|[CommentDeletedEventArgs](/.commentdeletedeventargs)|[commentDetails](/.commentdeletedeventargs#excel-javascript/api/excel/-commentdeletedeventargs-commentdetails-member)|Gets the `CommentDetail` array that contains the comment ID and IDs of its related replies.|
||[source](/.commentdeletedeventargs#excel-javascript/api/excel/-commentdeletedeventargs-source-member)|Specifies the source of the event.|
||[type](/.commentdeletedeventargs#excel-javascript/api/excel/-commentdeletedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.commentdeletedeventargs#excel-javascript/api/excel/-commentdeletedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the event happened.|
|[CommentDetail](/.commentdetail)|[commentId](/.commentdetail#excel-javascript/api/excel/-commentdetail-commentid-member)|Represents the ID of the comment.|
||[replyIds](/.commentdetail#excel-javascript/api/excel/-commentdetail-replyids-member)|Represents the IDs of the related replies that belong to the comment.|
|[CommentReply](/.commentreply)|[contentType](/.commentreply#excel-javascript/api/excel/-commentreply-contenttype-member)|The content type of the reply.|
|[CultureInfo](/.cultureinfo)|[datetimeFormat](/.cultureinfo#excel-javascript/api/excel/-cultureinfo-datetimeformat-member)|Defines the culturally appropriate format of displaying date and time.|
|[DatetimeFormatInfo](/.datetimeformatinfo)|[dateSeparator](/.datetimeformatinfo#excel-javascript/api/excel/-datetimeformatinfo-dateseparator-member)|Gets the string used as the date separator.|
||[longDatePattern](/.datetimeformatinfo#excel-javascript/api/excel/-datetimeformatinfo-longdatepattern-member)|Gets the format string for a long date value.|
||[longTimePattern](/.datetimeformatinfo#excel-javascript/api/excel/-datetimeformatinfo-longtimepattern-member)|Gets the format string for a long time value.|
||[shortDatePattern](/.datetimeformatinfo#excel-javascript/api/excel/-datetimeformatinfo-shortdatepattern-member)|Gets the format string for a short date value.|
||[timeSeparator](/.datetimeformatinfo#excel-javascript/api/excel/-datetimeformatinfo-timeseparator-member)|Gets the string used as the time separator.|
|[PivotDateFilter](/.pivotdatefilter)|[comparator](/.pivotdatefilter#excel-javascript/api/excel/-pivotdatefilter-comparator-member)|The comparator is the static value to which other values are compared.|
||[condition](/.pivotdatefilter#excel-javascript/api/excel/-pivotdatefilter-condition-member)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/.pivotdatefilter#excel-javascript/api/excel/-pivotdatefilter-exclusive-member)|If `true`, filter *excludes* items that meet criteria.|
||[lowerBound](/.pivotdatefilter#excel-javascript/api/excel/-pivotdatefilter-lowerbound-member)|The lower-bound of the range for the `between` filter condition.|
||[upperBound](/.pivotdatefilter#excel-javascript/api/excel/-pivotdatefilter-upperbound-member)|The upper-bound of the range for the `between` filter condition.|
||[wholeDays](/.pivotdatefilter#excel-javascript/api/excel/-pivotdatefilter-wholedays-member)|For `equals`, `before`, `after`, and `between` filter conditions, indicates if comparisons should be made as whole days.|
|[PivotField](/.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/.pivotfield#excel-javascript/api/excel/-pivotfield-applyfilter-member(1))|Sets one or more of the field's current PivotFilters and applies them to the field.|
||[clearAllFilters()](/.pivotfield#excel-javascript/api/excel/-pivotfield-clearallfilters-member(1))|Clears all criteria from all of the field's filters.|
||[clearFilter(filterType: Excel.PivotFilterType)](/.pivotfield#excel-javascript/api/excel/-pivotfield-clearfilter-member(1))|Clears all existing criteria from the field's filter of the given type (if one is currently applied).|
||[getFilters()](/.pivotfield#excel-javascript/api/excel/-pivotfield-getfilters-member(1))|Gets all filters currently applied on the field.|
||[isFiltered(filterType?: Excel.PivotFilterType)](/.pivotfield#excel-javascript/api/excel/-pivotfield-isfiltered-member(1))|Checks if there are any applied filters on the field.|
|[PivotFilters](/.pivotfilters)|[dateFilter](/.pivotfilters#excel-javascript/api/excel/-pivotfilters-datefilter-member)|The PivotField's currently applied date filter.|
||[labelFilter](/.pivotfilters#excel-javascript/api/excel/-pivotfilters-labelfilter-member)|The PivotField's currently applied label filter.|
||[manualFilter](/.pivotfilters#excel-javascript/api/excel/-pivotfilters-manualfilter-member)|The PivotField's currently applied manual filter.|
||[valueFilter](/.pivotfilters#excel-javascript/api/excel/-pivotfilters-valuefilter-member)|The PivotField's currently applied value filter.|
|[PivotLabelFilter](/.pivotlabelfilter)|[comparator](/.pivotlabelfilter#excel-javascript/api/excel/-pivotlabelfilter-comparator-member)|The comparator is the static value to which other values are compared.|
||[condition](/.pivotlabelfilter#excel-javascript/api/excel/-pivotlabelfilter-condition-member)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/.pivotlabelfilter#excel-javascript/api/excel/-pivotlabelfilter-exclusive-member)|If `true`, filter *excludes* items that meet criteria.|
||[lowerBound](/.pivotlabelfilter#excel-javascript/api/excel/-pivotlabelfilter-lowerbound-member)|The lower-bound of the range for the `between` filter condition.|
||[substring](/.pivotlabelfilter#excel-javascript/api/excel/-pivotlabelfilter-substring-member)|The substring used for the `beginsWith`, `endsWith`, and `contains` filter conditions.|
||[upperBound](/.pivotlabelfilter#excel-javascript/api/excel/-pivotlabelfilter-upperbound-member)|The upper-bound of the range for the `between` filter condition.|
|[PivotManualFilter](/.pivotmanualfilter)|[selectedItems](/.pivotmanualfilter#excel-javascript/api/excel/-pivotmanualfilter-selecteditems-member)|A list of selected items to manually filter.|
|[PivotTable](/.pivottable)|[allowMultipleFiltersPerField](/.pivottable#excel-javascript/api/excel/-pivottable-allowmultiplefiltersperfield-member)|Specifies if the PivotTable allows the application of multiple PivotFilters on a given PivotField in the table.|
|[PivotTableScopedCollection](/.pivottablescopedcollection)|[getCount()](/.pivottablescopedcollection#excel-javascript/api/excel/-pivottablescopedcollection-getcount-member(1))|Gets the number of PivotTables in the collection.|
||[getFirst()](/.pivottablescopedcollection#excel-javascript/api/excel/-pivottablescopedcollection-getfirst-member(1))|Gets the first PivotTable in the collection.|
||[getItem(key: string)](/.pivottablescopedcollection#excel-javascript/api/excel/-pivottablescopedcollection-getitem-member(1))|Gets a PivotTable by name.|
||[getItemOrNullObject(name: string)](/.pivottablescopedcollection#excel-javascript/api/excel/-pivottablescopedcollection-getitemornullobject-member(1))|Gets a PivotTable by name.|
||[items](/.pivottablescopedcollection#excel-javascript/api/excel/-pivottablescopedcollection-items-member)|Gets the loaded child items in this collection.|
|[PivotValueFilter](/.pivotvaluefilter)|[comparator](/.pivotvaluefilter#excel-javascript/api/excel/-pivotvaluefilter-comparator-member)|The comparator is the static value to which other values are compared.|
||[condition](/.pivotvaluefilter#excel-javascript/api/excel/-pivotvaluefilter-condition-member)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/.pivotvaluefilter#excel-javascript/api/excel/-pivotvaluefilter-exclusive-member)|If `true`, filter *excludes* items that meet criteria.|
||[lowerBound](/.pivotvaluefilter#excel-javascript/api/excel/-pivotvaluefilter-lowerbound-member)|The lower-bound of the range for the `between` filter condition.|
||[selectionType](/.pivotvaluefilter#excel-javascript/api/excel/-pivotvaluefilter-selectiontype-member)|Specifies if the filter is for the top/bottom N items, top/bottom N percent, or top/bottom N sum.|
||[threshold](/.pivotvaluefilter#excel-javascript/api/excel/-pivotvaluefilter-threshold-member)|The "N" threshold number of items, percent, or sum to be filtered for a top/bottom filter condition.|
||[upperBound](/.pivotvaluefilter#excel-javascript/api/excel/-pivotvaluefilter-upperbound-member)|The upper-bound of the range for the `between` filter condition.|
||[value](/.pivotvaluefilter#excel-javascript/api/excel/-pivotvaluefilter-value-member)|Name of the chosen "value" in the field by which to filter.|
|[Range](/.range)|[getDirectPrecedents()](/.range#excel-javascript/api/excel/-range-getdirectprecedents-member(1))|Returns a `WorkbookRangeAreas` object that represents the range containing all the direct precedent cells of a specified range in the same worksheet or across multiple worksheets.|
||[getPivotTables(fullyContained?: boolean)](/.range#excel-javascript/api/excel/-range-getpivottables-member(1))|Gets a scoped collection of PivotTables that overlap with the range.|
||[getSpillParent()](/.range#excel-javascript/api/excel/-range-getspillparent-member(1))|Gets the range object containing the anchor cell for a cell getting spilled into.|
||[getSpillParentOrNullObject()](/.range#excel-javascript/api/excel/-range-getspillparentornullobject-member(1))|Gets the range object containing the anchor cell for the cell getting spilled into.|
||[getSpillingToRange()](/.range#excel-javascript/api/excel/-range-getspillingtorange-member(1))|Gets the range object containing the spill range when called on an anchor cell.|
||[getSpillingToRangeOrNullObject()](/.range#excel-javascript/api/excel/-range-getspillingtorangeornullobject-member(1))|Gets the range object containing the spill range when called on an anchor cell.|
||[hasSpill](/.range#excel-javascript/api/excel/-range-hasspill-member)|Represents if all cells have a spill border.|
||[numberFormatCategories](/.range#excel-javascript/api/excel/-range-numberformatcategories-member)|Represents the category of number format of each cell.|
||[savedAsArray](/.range#excel-javascript/api/excel/-range-savedasarray-member)|Represents if all the cells would be saved as an array formula.|
|[RangeAreasCollection](/.rangeareascollection)|[getCount()](/.rangeareascollection#excel-javascript/api/excel/-rangeareascollection-getcount-member(1))|Gets the number of `RangeAreas` objects in this collection.|
||[getItemAt(index: number)](/.rangeareascollection#excel-javascript/api/excel/-rangeareascollection-getitemat-member(1))|Returns the `RangeAreas` object based on position in the collection.|
||[items](/.rangeareascollection#excel-javascript/api/excel/-rangeareascollection-items-member)|Gets the loaded child items in this collection.|
|[WorkbookRangeAreas](/.workbookrangeareas)|[addresses](/.workbookrangeareas#excel-javascript/api/excel/-workbookrangeareas-addresses-member)|Returns an array of addresses in A1-style.|
||[areas](/.workbookrangeareas#excel-javascript/api/excel/-workbookrangeareas-areas-member)|Returns the `RangeAreasCollection` object.|
||[getRangeAreasBySheet(key: string)](/.workbookrangeareas#excel-javascript/api/excel/-workbookrangeareas-getrangeareasbysheet-member(1))|Returns the `RangeAreas` object based on worksheet ID or name in the collection.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/.workbookrangeareas#excel-javascript/api/excel/-workbookrangeareas-getrangeareasornullobjectbysheet-member(1))|Returns the `RangeAreas` object based on worksheet name or ID in the collection.|
||[ranges](/.workbookrangeareas#excel-javascript/api/excel/-workbookrangeareas-ranges-member)|Returns ranges that comprise this object in a `RangeCollection` object.|
|[Worksheet](/.worksheet)|[customProperties](/.worksheet#excel-javascript/api/excel/-worksheet-customproperties-member)|Gets a collection of worksheet-level custom properties.|
|[WorksheetCustomProperty](/.worksheetcustomproperty)|[delete()](/.worksheetcustomproperty#excel-javascript/api/excel/-worksheetcustomproperty-delete-member(1))|Deletes the custom property.|
||[key](/.worksheetcustomproperty#excel-javascript/api/excel/-worksheetcustomproperty-key-member)|Gets the key of the custom property.|
||[value](/.worksheetcustomproperty#excel-javascript/api/excel/-worksheetcustomproperty-value-member)|Gets or sets the value of the custom property.|
|[WorksheetCustomPropertyCollection](/.worksheetcustompropertycollection)|[add(key: string, value: string)](/.worksheetcustompropertycollection#excel-javascript/api/excel/-worksheetcustompropertycollection-add-member(1))|Adds a new custom property that maps to the provided key.|
||[getCount()](/.worksheetcustompropertycollection#excel-javascript/api/excel/-worksheetcustompropertycollection-getcount-member(1))|Gets the number of custom properties on this worksheet.|
||[getItem(key: string)](/.worksheetcustompropertycollection#excel-javascript/api/excel/-worksheetcustompropertycollection-getitem-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/.worksheetcustompropertycollection#excel-javascript/api/excel/-worksheetcustompropertycollection-getitemornullobject-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[items](/.worksheetcustompropertycollection#excel-javascript/api/excel/-worksheetcustompropertycollection-items-member)|Gets the loaded child items in this collection.|
