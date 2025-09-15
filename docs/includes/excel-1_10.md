| Class | Fields | Description |
|:---|:---|:---|
|[Comment](/.comment)|[authorEmail](/.comment#excel-javascript/api/excel/-comment-authoremail-member)|Gets the email of the comment's author.|
||[authorName](/.comment#excel-javascript/api/excel/-comment-authorname-member)|Gets the name of the comment's author.|
||[content](/.comment#excel-javascript/api/excel/-comment-content-member)|The comment's content.|
||[creationDate](/.comment#excel-javascript/api/excel/-comment-creationdate-member)|Gets the creation time of the comment.|
||[delete()](/.comment#excel-javascript/api/excel/-comment-delete-member(1))|Deletes the comment and all the connected replies.|
||[getLocation()](/.comment#excel-javascript/api/excel/-comment-getlocation-member(1))|Gets the cell where this comment is located.|
||[id](/.comment#excel-javascript/api/excel/-comment-id-member)|Specifies the comment identifier.|
||[replies](/.comment#excel-javascript/api/excel/-comment-replies-member)|Represents a collection of reply objects associated with the comment.|
|[CommentCollection](/.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/.commentcollection#excel-javascript/api/excel/-commentcollection-add-member(1))|Creates a new comment with the given content on the given cell.|
||[getCount()](/.commentcollection#excel-javascript/api/excel/-commentcollection-getcount-member(1))|Gets the number of comments in the collection.|
||[getItem(commentId: string)](/.commentcollection#excel-javascript/api/excel/-commentcollection-getitem-member(1))|Gets a comment from the collection based on its ID.|
||[getItemAt(index: number)](/.commentcollection#excel-javascript/api/excel/-commentcollection-getitemat-member(1))|Gets a comment from the collection based on its position.|
||[getItemByCell(cellAddress: Range \| string)](/.commentcollection#excel-javascript/api/excel/-commentcollection-getitembycell-member(1))|Gets the comment from the specified cell.|
||[getItemByReplyId(replyId: string)](/.commentcollection#excel-javascript/api/excel/-commentcollection-getitembyreplyid-member(1))|Gets the comment to which the given reply is connected.|
||[items](/.commentcollection#excel-javascript/api/excel/-commentcollection-items-member)|Gets the loaded child items in this collection.|
|[CommentReply](/.commentreply)|[authorEmail](/.commentreply#excel-javascript/api/excel/-commentreply-authoremail-member)|Gets the email of the comment reply's author.|
||[authorName](/.commentreply#excel-javascript/api/excel/-commentreply-authorname-member)|Gets the name of the comment reply's author.|
||[content](/.commentreply#excel-javascript/api/excel/-commentreply-content-member)|The comment reply's content.|
||[creationDate](/.commentreply#excel-javascript/api/excel/-commentreply-creationdate-member)|Gets the creation time of the comment reply.|
||[delete()](/.commentreply#excel-javascript/api/excel/-commentreply-delete-member(1))|Deletes the comment reply.|
||[getLocation()](/.commentreply#excel-javascript/api/excel/-commentreply-getlocation-member(1))|Gets the cell where this comment reply is located.|
||[getParentComment()](/.commentreply#excel-javascript/api/excel/-commentreply-getparentcomment-member(1))|Gets the parent comment of this reply.|
||[id](/.commentreply#excel-javascript/api/excel/-commentreply-id-member)|Specifies the comment reply identifier.|
|[CommentReplyCollection](/.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/.commentreplycollection#excel-javascript/api/excel/-commentreplycollection-add-member(1))|Creates a comment reply for a comment.|
||[getCount()](/.commentreplycollection#excel-javascript/api/excel/-commentreplycollection-getcount-member(1))|Gets the number of comment replies in the collection.|
||[getItem(commentReplyId: string)](/.commentreplycollection#excel-javascript/api/excel/-commentreplycollection-getitem-member(1))|Returns a comment reply identified by its ID.|
||[getItemAt(index: number)](/.commentreplycollection#excel-javascript/api/excel/-commentreplycollection-getitemat-member(1))|Gets a comment reply based on its position in the collection.|
||[items](/.commentreplycollection#excel-javascript/api/excel/-commentreplycollection-items-member)|Gets the loaded child items in this collection.|
|[PivotLayout](/.pivotlayout)|[enableFieldList](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-enablefieldlist-member)|Specifies if the field list can be shown in the UI.|
|[PivotTableStyle](/.pivottablestyle)|[delete()](/.pivottablestyle#excel-javascript/api/excel/-pivottablestyle-delete-member(1))|Deletes the PivotTable style.|
||[duplicate()](/.pivottablestyle#excel-javascript/api/excel/-pivottablestyle-duplicate-member(1))|Creates a duplicate of this PivotTable style with copies of all the style elements.|
||[name](/.pivottablestyle#excel-javascript/api/excel/-pivottablestyle-name-member)|Specifies the name of the PivotTable style.|
||[readOnly](/.pivottablestyle#excel-javascript/api/excel/-pivottablestyle-readonly-member)|Specifies if this `PivotTableStyle` object is read-only.|
|[PivotTableStyleCollection](/.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/.pivottablestylecollection#excel-javascript/api/excel/-pivottablestylecollection-add-member(1))|Creates a blank `PivotTableStyle` with the specified name.|
||[getCount()](/.pivottablestylecollection#excel-javascript/api/excel/-pivottablestylecollection-getcount-member(1))|Gets the number of PivotTable styles in the collection.|
||[getDefault()](/.pivottablestylecollection#excel-javascript/api/excel/-pivottablestylecollection-getdefault-member(1))|Gets the default PivotTable style for the parent object's scope.|
||[getItem(name: string)](/.pivottablestylecollection#excel-javascript/api/excel/-pivottablestylecollection-getitem-member(1))|Gets a `PivotTableStyle` by name.|
||[getItemOrNullObject(name: string)](/.pivottablestylecollection#excel-javascript/api/excel/-pivottablestylecollection-getitemornullobject-member(1))|Gets a `PivotTableStyle` by name.|
||[items](/.pivottablestylecollection#excel-javascript/api/excel/-pivottablestylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/.pivottablestylecollection#excel-javascript/api/excel/-pivottablestylecollection-setdefault-member(1))|Sets the default PivotTable style for use in the parent object's scope.|
|[Range](/.range)|[group(groupOption: Excel.GroupOption)](/.range#excel-javascript/api/excel/-range-group-member(1))|Groups columns and rows for an outline.|
||[height](/.range#excel-javascript/api/excel/-range-height-member)|Returns the distance in points, for 100% zoom, from the top edge of the range to the bottom edge of the range.|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/.range#excel-javascript/api/excel/-range-hidegroupdetails-member(1))|Hides the details of the row or column group.|
||[left](/.range#excel-javascript/api/excel/-range-left-member)|Returns the distance in points, for 100% zoom, from the left edge of the worksheet to the left edge of the range.|
||[showGroupDetails(groupOption: Excel.GroupOption)](/.range#excel-javascript/api/excel/-range-showgroupdetails-member(1))|Shows the details of the row or column group.|
||[top](/.range#excel-javascript/api/excel/-range-top-member)|Returns the distance in points, for 100% zoom, from the top edge of the worksheet to the top edge of the range.|
||[ungroup(groupOption: Excel.GroupOption)](/.range#excel-javascript/api/excel/-range-ungroup-member(1))|Ungroups columns and rows for an outline.|
||[width](/.range#excel-javascript/api/excel/-range-width-member)|Returns the distance in points, for 100% zoom, from the left edge of the range to the right edge of the range.|
|[Shape](/.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/.shape#excel-javascript/api/excel/-shape-copyto-member(1))|Copies and pastes a `Shape` object.|
||[placement](/.shape#excel-javascript/api/excel/-shape-placement-member)|Represents how the object is attached to the cells below it.|
|[Slicer](/.slicer)|[caption](/.slicer#excel-javascript/api/excel/-slicer-caption-member)|Represents the caption of the slicer.|
||[clearFilters()](/.slicer#excel-javascript/api/excel/-slicer-clearfilters-member(1))|Clears all the filters currently applied on the slicer.|
||[delete()](/.slicer#excel-javascript/api/excel/-slicer-delete-member(1))|Deletes the slicer.|
||[getSelectedItems()](/.slicer#excel-javascript/api/excel/-slicer-getselecteditems-member(1))|Returns an array of selected items' keys.|
||[height](/.slicer#excel-javascript/api/excel/-slicer-height-member)|Specifies the height, in points, of the slicer.|
||[id](/.slicer#excel-javascript/api/excel/-slicer-id-member)|Represents the unique ID of the slicer.|
||[isFilterCleared](/.slicer#excel-javascript/api/excel/-slicer-isfiltercleared-member)|Value is `true` if all filters currently applied on the slicer are cleared.|
||[left](/.slicer#excel-javascript/api/excel/-slicer-left-member)|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|
||[name](/.slicer#excel-javascript/api/excel/-slicer-name-member)|Represents the name of the slicer.|
||[selectItems(items?: string[])](/.slicer#excel-javascript/api/excel/-slicer-selectitems-member(1))|Selects slicer items based on their keys.|
||[slicerItems](/.slicer#excel-javascript/api/excel/-slicer-sliceritems-member)|Represents the collection of slicer items that are part of the slicer.|
||[sortBy](/.slicer#excel-javascript/api/excel/-slicer-sortby-member)|Specifies the sort order of the items in the slicer.|
||[style](/.slicer#excel-javascript/api/excel/-slicer-style-member)|Constant value that represents the slicer style.|
||[top](/.slicer#excel-javascript/api/excel/-slicer-top-member)|Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.|
||[width](/.slicer#excel-javascript/api/excel/-slicer-width-member)|Represents the width, in points, of the slicer.|
||[worksheet](/.slicer#excel-javascript/api/excel/-slicer-worksheet-member)|Represents the worksheet containing the slicer.|
|[SlicerCollection](/.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/.slicercollection#excel-javascript/api/excel/-slicercollection-add-member(1))|Adds a new slicer to the workbook.|
||[getCount()](/.slicercollection#excel-javascript/api/excel/-slicercollection-getcount-member(1))|Returns the number of slicers in the collection.|
||[getItem(key: string)](/.slicercollection#excel-javascript/api/excel/-slicercollection-getitem-member(1))|Gets a slicer object using its name or ID.|
||[getItemAt(index: number)](/.slicercollection#excel-javascript/api/excel/-slicercollection-getitemat-member(1))|Gets a slicer based on its position in the collection.|
||[getItemOrNullObject(key: string)](/.slicercollection#excel-javascript/api/excel/-slicercollection-getitemornullobject-member(1))|Gets a slicer using its name or ID.|
||[items](/.slicercollection#excel-javascript/api/excel/-slicercollection-items-member)|Gets the loaded child items in this collection.|
|[SlicerItem](/.sliceritem)|[hasData](/.sliceritem#excel-javascript/api/excel/-sliceritem-hasdata-member)|Value is `true` if the slicer item has data.|
||[isSelected](/.sliceritem#excel-javascript/api/excel/-sliceritem-isselected-member)|Value is `true` if the slicer item is selected.|
||[key](/.sliceritem#excel-javascript/api/excel/-sliceritem-key-member)|Represents the unique value representing the slicer item.|
||[name](/.sliceritem#excel-javascript/api/excel/-sliceritem-name-member)|Represents the title displayed in the Excel UI.|
|[SlicerItemCollection](/.sliceritemcollection)|[getCount()](/.sliceritemcollection#excel-javascript/api/excel/-sliceritemcollection-getcount-member(1))|Returns the number of slicer items in the slicer.|
||[getItem(key: string)](/.sliceritemcollection#excel-javascript/api/excel/-sliceritemcollection-getitem-member(1))|Gets a slicer item object using its key or name.|
||[getItemAt(index: number)](/.sliceritemcollection#excel-javascript/api/excel/-sliceritemcollection-getitemat-member(1))|Gets a slicer item based on its position in the collection.|
||[getItemOrNullObject(key: string)](/.sliceritemcollection#excel-javascript/api/excel/-sliceritemcollection-getitemornullobject-member(1))|Gets a slicer item using its key or name.|
||[items](/.sliceritemcollection#excel-javascript/api/excel/-sliceritemcollection-items-member)|Gets the loaded child items in this collection.|
|[SlicerStyle](/.slicerstyle)|[delete()](/.slicerstyle#excel-javascript/api/excel/-slicerstyle-delete-member(1))|Deletes the slicer style.|
||[duplicate()](/.slicerstyle#excel-javascript/api/excel/-slicerstyle-duplicate-member(1))|Creates a duplicate of this slicer style with copies of all the style elements.|
||[name](/.slicerstyle#excel-javascript/api/excel/-slicerstyle-name-member)|Specifies the name of the slicer style.|
||[readOnly](/.slicerstyle#excel-javascript/api/excel/-slicerstyle-readonly-member)|Specifies if this `SlicerStyle` object is read-only.|
|[SlicerStyleCollection](/.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/.slicerstylecollection#excel-javascript/api/excel/-slicerstylecollection-add-member(1))|Creates a blank slicer style with the specified name.|
||[getCount()](/.slicerstylecollection#excel-javascript/api/excel/-slicerstylecollection-getcount-member(1))|Gets the number of slicer styles in the collection.|
||[getDefault()](/.slicerstylecollection#excel-javascript/api/excel/-slicerstylecollection-getdefault-member(1))|Gets the default `SlicerStyle` for the parent object's scope.|
||[getItem(name: string)](/.slicerstylecollection#excel-javascript/api/excel/-slicerstylecollection-getitem-member(1))|Gets a `SlicerStyle` by name.|
||[getItemOrNullObject(name: string)](/.slicerstylecollection#excel-javascript/api/excel/-slicerstylecollection-getitemornullobject-member(1))|Gets a `SlicerStyle` by name.|
||[items](/.slicerstylecollection#excel-javascript/api/excel/-slicerstylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/.slicerstylecollection#excel-javascript/api/excel/-slicerstylecollection-setdefault-member(1))|Sets the default slicer style for use in the parent object's scope.|
|[TableStyle](/.tablestyle)|[delete()](/.tablestyle#excel-javascript/api/excel/-tablestyle-delete-member(1))|Deletes the table style.|
||[duplicate()](/.tablestyle#excel-javascript/api/excel/-tablestyle-duplicate-member(1))|Creates a duplicate of this table style with copies of all the style elements.|
||[name](/.tablestyle#excel-javascript/api/excel/-tablestyle-name-member)|Specifies the name of the table style.|
||[readOnly](/.tablestyle#excel-javascript/api/excel/-tablestyle-readonly-member)|Specifies if this `TableStyle` object is read-only.|
|[TableStyleCollection](/.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/.tablestylecollection#excel-javascript/api/excel/-tablestylecollection-add-member(1))|Creates a blank `TableStyle` with the specified name.|
||[getCount()](/.tablestylecollection#excel-javascript/api/excel/-tablestylecollection-getcount-member(1))|Gets the number of table styles in the collection.|
||[getDefault()](/.tablestylecollection#excel-javascript/api/excel/-tablestylecollection-getdefault-member(1))|Gets the default table style for the parent object's scope.|
||[getItem(name: string)](/.tablestylecollection#excel-javascript/api/excel/-tablestylecollection-getitem-member(1))|Gets a `TableStyle` by name.|
||[getItemOrNullObject(name: string)](/.tablestylecollection#excel-javascript/api/excel/-tablestylecollection-getitemornullobject-member(1))|Gets a `TableStyle` by name.|
||[items](/.tablestylecollection#excel-javascript/api/excel/-tablestylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/.tablestylecollection#excel-javascript/api/excel/-tablestylecollection-setdefault-member(1))|Sets the default table style for use in the parent object's scope.|
|[TimelineStyle](/.timelinestyle)|[delete()](/.timelinestyle#excel-javascript/api/excel/-timelinestyle-delete-member(1))|Deletes the table style.|
||[duplicate()](/.timelinestyle#excel-javascript/api/excel/-timelinestyle-duplicate-member(1))|Creates a duplicate of this timeline style with copies of all the style elements.|
||[name](/.timelinestyle#excel-javascript/api/excel/-timelinestyle-name-member)|Specifies the name of the timeline style.|
||[readOnly](/.timelinestyle#excel-javascript/api/excel/-timelinestyle-readonly-member)|Specifies if this `TimelineStyle` object is read-only.|
|[TimelineStyleCollection](/.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/.timelinestylecollection#excel-javascript/api/excel/-timelinestylecollection-add-member(1))|Creates a blank `TimelineStyle` with the specified name.|
||[getCount()](/.timelinestylecollection#excel-javascript/api/excel/-timelinestylecollection-getcount-member(1))|Gets the number of timeline styles in the collection.|
||[getDefault()](/.timelinestylecollection#excel-javascript/api/excel/-timelinestylecollection-getdefault-member(1))|Gets the default timeline style for the parent object's scope.|
||[getItem(name: string)](/.timelinestylecollection#excel-javascript/api/excel/-timelinestylecollection-getitem-member(1))|Gets a `TimelineStyle` by name.|
||[getItemOrNullObject(name: string)](/.timelinestylecollection#excel-javascript/api/excel/-timelinestylecollection-getitemornullobject-member(1))|Gets a `TimelineStyle` by name.|
||[items](/.timelinestylecollection#excel-javascript/api/excel/-timelinestylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/.timelinestylecollection#excel-javascript/api/excel/-timelinestylecollection-setdefault-member(1))|Sets the default timeline style for use in the parent object's scope.|
|[Workbook](/.workbook)|[comments](/.workbook#excel-javascript/api/excel/-workbook-comments-member)|Represents a collection of comments associated with the workbook.|
||[getActiveSlicer()](/.workbook#excel-javascript/api/excel/-workbook-getactiveslicer-member(1))|Gets the currently active slicer in the workbook.|
||[getActiveSlicerOrNullObject()](/.workbook#excel-javascript/api/excel/-workbook-getactiveslicerornullobject-member(1))|Gets the currently active slicer in the workbook.|
||[pivotTableStyles](/.workbook#excel-javascript/api/excel/-workbook-pivottablestyles-member)|Represents a collection of PivotTableStyles associated with the workbook.|
||[slicerStyles](/.workbook#excel-javascript/api/excel/-workbook-slicerstyles-member)|Represents a collection of SlicerStyles associated with the workbook.|
||[slicers](/.workbook#excel-javascript/api/excel/-workbook-slicers-member)|Represents a collection of slicers associated with the workbook.|
||[tableStyles](/.workbook#excel-javascript/api/excel/-workbook-tablestyles-member)|Represents a collection of TableStyles associated with the workbook.|
||[timelineStyles](/.workbook#excel-javascript/api/excel/-workbook-timelinestyles-member)|Represents a collection of TimelineStyles associated with the workbook.|
|[Worksheet](/.worksheet)|[comments](/.worksheet#excel-javascript/api/excel/-worksheet-comments-member)|Returns a collection of all the Comments objects on the worksheet.|
||[onColumnSorted](/.worksheet#excel-javascript/api/excel/-worksheet-oncolumnsorted-member)|Occurs when one or more columns have been sorted.|
||[onRowSorted](/.worksheet#excel-javascript/api/excel/-worksheet-onrowsorted-member)|Occurs when one or more rows have been sorted.|
||[onSingleClicked](/.worksheet#excel-javascript/api/excel/-worksheet-onsingleclicked-member)|Occurs when a left-clicked/tapped action happens in the worksheet.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/.worksheet#excel-javascript/api/excel/-worksheet-showoutlinelevels-member(1))|Shows row or column groups by their outline levels.|
||[slicers](/.worksheet#excel-javascript/api/excel/-worksheet-slicers-member)|Returns a collection of slicers that are part of the worksheet.|
|[WorksheetCollection](/.worksheetcollection)|[onColumnSorted](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-oncolumnsorted-member)|Occurs when one or more columns have been sorted.|
||[onRowSorted](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onrowsorted-member)|Occurs when one or more rows have been sorted.|
||[onSingleClicked](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onsingleclicked-member)|Occurs when left-clicked/tapped operation happens in the worksheet collection.|
|[WorksheetColumnSortedEventArgs](/.worksheetcolumnsortedeventargs)|[address](/.worksheetcolumnsortedeventargs#excel-javascript/api/excel/-worksheetcolumnsortedeventargs-address-member)|Gets the range address that represents the sorted areas of a specific worksheet.|
||[source](/.worksheetcolumnsortedeventargs#excel-javascript/api/excel/-worksheetcolumnsortedeventargs-source-member)|Gets the source of the event.|
||[type](/.worksheetcolumnsortedeventargs#excel-javascript/api/excel/-worksheetcolumnsortedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetcolumnsortedeventargs#excel-javascript/api/excel/-worksheetcolumnsortedeventargs-worksheetid-member)|Gets the ID of the worksheet where the sorting happened.|
|[WorksheetRowSortedEventArgs](/.worksheetrowsortedeventargs)|[address](/.worksheetrowsortedeventargs#excel-javascript/api/excel/-worksheetrowsortedeventargs-address-member)|Gets the range address that represents the sorted areas of a specific worksheet.|
||[source](/.worksheetrowsortedeventargs#excel-javascript/api/excel/-worksheetrowsortedeventargs-source-member)|Gets the source of the event.|
||[type](/.worksheetrowsortedeventargs#excel-javascript/api/excel/-worksheetrowsortedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetrowsortedeventargs#excel-javascript/api/excel/-worksheetrowsortedeventargs-worksheetid-member)|Gets the ID of the worksheet where the sorting happened.|
|[WorksheetSingleClickedEventArgs](/.worksheetsingleclickedeventargs)|[address](/.worksheetsingleclickedeventargs#excel-javascript/api/excel/-worksheetsingleclickedeventargs-address-member)|Gets the address that represents the cell which was left-clicked/tapped for a specific worksheet.|
||[offsetX](/.worksheetsingleclickedeventargs#excel-javascript/api/excel/-worksheetsingleclickedeventargs-offsetx-member)|The distance, in points, from the left-clicked/tapped point to the left (or right for right-to-left languages) gridline edge of the left-clicked/tapped cell.|
||[offsetY](/.worksheetsingleclickedeventargs#excel-javascript/api/excel/-worksheetsingleclickedeventargs-offsety-member)|The distance, in points, from the left-clicked/tapped point to the top gridline edge of the left-clicked/tapped cell.|
||[type](/.worksheetsingleclickedeventargs#excel-javascript/api/excel/-worksheetsingleclickedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetsingleclickedeventargs#excel-javascript/api/excel/-worksheetsingleclickedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the cell was left-clicked/tapped.|
