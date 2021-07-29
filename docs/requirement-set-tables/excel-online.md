| Class | Fields | Description |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|Clears the column filter criteria of the AutoFilter.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate__)|Activates this sheet view.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete__)|Removes the sheet view from the worksheet.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate_name_)|Creates a copy of this sheet view.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Gets or sets the name of the sheet view.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add_name_)|Creates a new sheet view with the given name.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#enterTemporary__)|Creates and activates a new temporary sheet view.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit__)|Exits the currently active sheet view.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getActive__)|Gets the worksheet's currently active sheet view.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getCount__)|Gets the number of sheet views in this worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItem_key_)|Gets a sheet view using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getItemAt_index_)|Gets a sheet view by its index in the collection.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Gets the loaded child items in this collection.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|Returns a collection of sheet views that are present in the worksheet.|
