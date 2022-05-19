| Class | Fields | Description |
|:---|:---|:---|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|Creates a new comment with the given content on the given cell.|
||[authorEmail](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-authoremail-member)|Gets the email of the comment's author.|
||[authorName](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-authorname-member)|Gets the name of the comment's author.|
||[content](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-content-member)|The comment's content.|
||[getCount()](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getcount-member(1))|Gets the number of comments in the collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitem-member(1))|Gets a comment from the collection based on its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemat-member(1))|Gets a comment from the collection based on its position.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembycell-member(1))|Gets the comment from the specified cell.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembyreplyid-member(1))|Gets the comment to which the given reply is connected.|
||[items](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-items-member)|Gets the loaded child items in this collection.|
||[replies](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-replies-member)|Represents a collection of reply objects associated with the comment.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|Creates a comment reply for a comment.|
||[authorEmail](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-authoremail-member)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-authorname-member)|Gets the name of the comment reply's author.|
||[content](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-content-member)|The comment reply's content.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getcount-member(1))|Gets the number of comment replies in the collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitem-member(1))|Returns a comment reply identified by its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemat-member(1))|Gets a comment reply based on its position in the collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-items-member)|Gets the loaded child items in this collection.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-delete-member(1))|Deletes the PivotTable style.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-duplicate-member(1))|Creates a duplicate of this PivotTable style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-name-member)|Gets the name of the PivotTable style.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-readonly-member)|Specifies if this `PivotTableStyle` object is read-only.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-add-member(1))|Creates a blank `PivotTableStyle` with the specified name.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getcount-member(1))|Gets the number of PivotTable styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getdefault-member(1))|Gets the default PivotTable style for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitem-member(1))|Gets a `PivotTableStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitemornullobject-member(1))|Gets a `PivotTableStyle` by name.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-setdefault-member(1))|Sets the default PivotTable style for use in the parent object's scope.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-group-member(1))|Groups columns and rows for an outline.|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-hidegroupdetails-member(1))|Hides the details of the row or column group.|
||[insert(shift: Excel.InsertShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-insert-member(1))|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space.|
||[merge(across?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-merge-member(1))|Merge the range cells into one region in the worksheet.|
|[RangeReference](/javascript/api/excel/excel.rangereference)|[address](/javascript/api/excel/excel.rangereference#excel-excel-rangereference-address-member)|The address of the range, for example "SheetName!A1:B5".|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-copyto-member(1))|Copies and pastes a `Shape` object.|
||[delete()](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-delete-member(1))|Removes the shape from the worksheet.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getasimage-member(1))|Converts the shape to an image and returns the image as a base64-encoded string.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-incrementleft-member(1))|Moves the shape horizontally by the specified number of points.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-incrementrotation-member(1))|Rotates the shape clockwise around the z-axis by the specified number of degrees.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-incrementtop-member(1))|Moves the shape vertically by the specified number of points.|
||[placement](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-placement-member)|Represents how the object is attached to the cells below it.|
||[rotation](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-rotation-member)|Specifies the rotation, in degrees, of the shape.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-scaleheight-member(1))|Scales the height of the shape by a specified factor.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-scalewidth-member(1))|Scales the width of the shape by a specified factor.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-setzorder-member(1))|Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.|
||[top](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-top-member)|The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[type](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-type-member)|Returns the type of this shape.|
||[visible](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-visible-member)|Specifies if the shape is visible.|
||[width](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-width-member)|Specifies the width, in points, of the shape.|
||[zOrderPosition](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-zorderposition-member)|Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack.|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#excel-excel-slicer-caption-member)|Represents the caption of the slicer.|
||[clearFilters()](/javascript/api/excel/excel.slicer#excel-excel-slicer-clearfilters-member(1))|Clears all the filters currently applied on the slicer.|
||[delete()](/javascript/api/excel/excel.slicer#excel-excel-slicer-delete-member(1))|Deletes the slicer.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#excel-excel-slicer-getselecteditems-member(1))|Returns an array of selected items' keys.|
||[height](/javascript/api/excel/excel.slicer#excel-excel-slicer-height-member)|Represents the height, in points, of the slicer.|
||[id](/javascript/api/excel/excel.slicer#excel-excel-slicer-id-member)|Represents the unique ID of the slicer.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#excel-excel-slicer-isfiltercleared-member)|Value is `true` if all filters currently applied on the slicer are cleared.|
||[left](/javascript/api/excel/excel.slicer#excel-excel-slicer-left-member)|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|
||[name](/javascript/api/excel/excel.slicer#excel-excel-slicer-name-member)|Represents the name of the slicer.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#excel-excel-slicer-selectitems-member(1))|Selects slicer items based on their keys.|
||[slicerItems](/javascript/api/excel/excel.slicer#excel-excel-slicer-sliceritems-member)|Represents the collection of slicer items that are part of the slicer.|
||[sortBy](/javascript/api/excel/excel.slicer#excel-excel-slicer-sortby-member)|Represents the sort order of the items in the slicer.|
||[style](/javascript/api/excel/excel.slicer#excel-excel-slicer-style-member)|Constant value that represents the slicer style.|
||[top](/javascript/api/excel/excel.slicer#excel-excel-slicer-top-member)|Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.|
||[width](/javascript/api/excel/excel.slicer#excel-excel-slicer-width-member)|Represents the width, in points, of the slicer.|
||[worksheet](/javascript/api/excel/excel.slicer#excel-excel-slicer-worksheet-member)|Represents the worksheet containing the slicer.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-add-member(1))|Adds a new slicer to the workbook.|
||[getCount()](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getcount-member(1))|Returns the number of slicers in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitem-member(1))|Gets a slicer object using its name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemat-member(1))|Gets a slicer based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemornullobject-member(1))|Gets a slicer using its name or ID.|
||[items](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-items-member)|Gets the loaded child items in this collection.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[hasData](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-hasdata-member)|Value is `true` if the slicer item has data.|
||[isSelected](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-isselected-member)|Value is `true` if the slicer item is selected.|
||[key](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-key-member)|Represents the unique value representing the slicer item.|
||[name](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-name-member)|Represents the title displayed in the Excel UI.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getcount-member(1))|Returns the number of slicer items in the slicer.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitem-member(1))|Gets a slicer item object using its key or name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemat-member(1))|Gets a slicer item based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemornullobject-member(1))|Gets a slicer item using its key or name.|
||[items](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-items-member)|Gets the loaded child items in this collection.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-delete-member(1))|Deletes the slicer style.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-duplicate-member(1))|Creates a duplicate of this slicer style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-name-member)|Gets the name of the slicer style.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-readonly-member)|Specifies if this `SlicerStyle` object is read-only.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-add-member(1))|Creates a blank slicer style with the specified name.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getcount-member(1))|Gets the number of slicer styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getdefault-member(1))|Gets the default `SlicerStyle` for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitem-member(1))|Gets a `SlicerStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitemornullobject-member(1))|Gets a `SlicerStyle` by name.|
||[items](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-setdefault-member(1))|Sets the default slicer style for use in the parent object's scope.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-add-member(1))|Creates a blank `TableStyle` with the specified name.|
||[getDefault()](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getdefault-member(1))|Gets the default table style for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitem-member(1))|Gets a `TableStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemornullobject-member(1))|Gets a `TableStyle` by name.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-setdefault-member(1))|Sets the default table style for use in the parent object's scope.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-delete-member(1))|Deletes the table style.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-duplicate-member(1))|Creates a duplicate of this table style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-name-member)|Gets the name of the table style.|
||[readOnly](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-readonly-member)|Specifies if this `TableStyle` object is read-only.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-delete-member(1))|Deletes the table style.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-duplicate-member(1))|Creates a duplicate of this timeline style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-name-member)|Gets the name of the timeline style.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-readonly-member)|Specifies if this `TimelineStyle` object is read-only.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-add-member(1))|Creates a blank `TimelineStyle` with the specified name.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getcount-member(1))|Gets the number of timeline styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getdefault-member(1))|Gets the default timeline style for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitem-member(1))|Gets a `TimelineStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitemornullobject-member(1))|Gets a `TimelineStyle` by name.|
||[items](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-setdefault-member(1))|Sets the default timeline style for use in the parent object's scope.|
|[Workbook](/javascript/api/excel/excel.workbook)|[comments](/javascript/api/excel/excel.workbook#excel-excel-workbook-comments-member)|Represents a collection of comments associated with the workbook.|
||[customXmlParts](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member)|Represents the collection of custom XML parts contained by this workbook.|
||[dataConnections](/javascript/api/excel/excel.workbook#excel-excel-workbook-dataconnections-member)|Represents all data connections in the workbook.|
||[functions](/javascript/api/excel/excel.workbook#excel-excel-workbook-functions-member)|Represents a collection of worksheet functions that can be used for computation.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicer-member(1))|Gets the currently active slicer in the workbook.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicerornullobject-member(1))|Gets the currently active slicer in the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-comments-member)|Returns a collection of all the Comments objects on the worksheet.|
||[getLast(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getlast-member(1))|Gets the last worksheet in the collection.|
||[getPrevious(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getprevious-member(1))|Gets the worksheet that precedes this one.|
||[getPreviousOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getpreviousornullobject-member(1))|Gets the worksheet that precedes this one.|
||[onAdded](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onadded-member)|Occurs when a new worksheet is added to the workbook.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member)|Occurs when one or more columns have been sorted.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member)|Occurs when one or more columns have been sorted.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member)|Occurs when the worksheet is deactivated.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member)|Occurs when any worksheet in the workbook is deactivated.|
||[onDeleted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeleted-member)|Occurs when a worksheet is deleted from the workbook.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member)|Occurs when format changed on a specific worksheet.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member)|Occurs when any worksheet in the workbook has a format changed.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showoutlinelevels-member(1))|Shows row or column groups by their outline levels.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-address-member)|Gets the range address that represents the sorted areas of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-worksheetid-member)|Gets the ID of the worksheet where the sorting happened.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowAutoFilter](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowautofilter-member)|Represents the worksheet protection option allowing use of the AutoFilter feature.|
||[allowDeleteColumns](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowdeletecolumns-member)|Represents the worksheet protection option allowing deleting of columns.|
||[allowDeleteRows](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowdeleterows-member)|Represents the worksheet protection option allowing deleting of rows.|
||[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditobjects-member)|Represents the worksheet protection option allowing editing of objects.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditscenarios-member)|Represents the worksheet protection option allowing editing of scenarios.|
||[allowFormatCells](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowformatcells-member)|Represents the worksheet protection option allowing formatting of cells.|
||[allowFormatColumns](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowformatcolumns-member)|Represents the worksheet protection option allowing formatting of columns.|
||[allowFormatRows](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowformatrows-member)|Represents the worksheet protection option allowing formatting of rows.|
||[allowInsertColumns](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowinsertcolumns-member)|Represents the worksheet protection option allowing inserting of columns.|
||[allowInsertHyperlinks](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowinserthyperlinks-member)|Represents the worksheet protection option allowing inserting of hyperlinks.|
||[allowInsertRows](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowinsertrows-member)|Represents the worksheet protection option allowing inserting of rows.|
||[allowPivotTables](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowpivottables-member)|Represents the worksheet protection option allowing use of the PivotTable feature.|
||[allowSort](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-allowsort-member)|Represents the worksheet protection option allowing use of the sort feature.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-selectionmode-member)|Represents the worksheet protection option of selection mode.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-address-member)|Gets the range address that represents the sorted areas of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-worksheetid-member)|Gets the ID of the worksheet where the sorting happened.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-address-member)|Gets the address that represents the cell which was left-clicked/tapped for a specific worksheet.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsetx-member)|The distance, in points, from the left-clicked/tapped point to the left (or right for right-to-left languages) gridline edge of the left-clicked/tapped cell.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsety-member)|The distance, in points, from the left-clicked/tapped point to the top gridline edge of the left-clicked/tapped cell.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the cell was left-clicked/tapped.|
