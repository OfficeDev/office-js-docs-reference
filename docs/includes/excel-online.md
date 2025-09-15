| Class | Fields | Description |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-address-member)|Specifies the range associated with the object.|
||[delete()](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-delete-member(1))|Deletes the object from the `AllowEditRangeCollection`.|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-ispasswordprotected-member)|Specifies if the object is password protected.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-pauseprotection-member(1))|Pauses worksheet protection for the object for the user in the current session.|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-setpassword-member(1))|Changes the password associated with the object.|
||[title](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-title-member)|Specifies the title of the object.|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel.AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-add-member(1))|Adds an `AllowEditRange` object to the worksheet.|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getcount-member(1))|Returns the number of `AllowEditRange` objects in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitem-member(1))|Gets the `AllowEditRange` object by its title.|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemat-member(1))|Returns an `AllowEditRange` object by its index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemornullobject-member(1))|Gets the `AllowEditRange` object by its title.|
||[items](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-items-member)|Gets the loaded child items in this collection.|
||[pauseProtection(password: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-pauseprotection-member(1))|Pauses worksheet protection for all `AllowEditRange` objects found in this worksheet that have the given password for the user in the current session.|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[password](/javascript/api/excel/excel.alloweditrangeoptions#excel-excel-alloweditrangeoptions-password-member)|The password associated with the `AllowEditRange`.|
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
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterows-member(1))|Delete multiple rows from a table.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterowsat-member(1))|Delete a specified number of rows from a table, starting at a given index.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkedworkbooks-member)|Returns a collection of linked workbooks.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-namedsheetviews-member)|Returns a collection of sheet views that are present in the worksheet.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-alloweditranges-member)|Specifies the `AllowEditRangeCollection` object found in this worksheet.|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-canpauseprotection-member)|Specifies if protection can be paused for this worksheet.|
||[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-checkpassword-member(1))|Specifies if the password can be used to unlock worksheet protection.|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispasswordprotected-member)|Specifies if the sheet is password protected.|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispaused-member)|Specifies if worksheet protection is paused.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-pauseprotection-member(1))|Pauses worksheet protection for the given worksheet object for the user in the current session.|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-resumeprotection-member(1))|Resumes worksheet protection for the given worksheet object for the user in a given session.|
||[savedOptions](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-savedoptions-member)|Specifies the protection options saved in the worksheet.|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-setpassword-member(1))|Changes the password associated with the `WorksheetProtection` object.|
||[updateOptions(options: Excel.WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-updateoptions-member(1))|Change the worksheet protection options associated with the `WorksheetProtection` object.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-alloweditrangeschanged-member)|Specifies if any of the `AllowEditRange` objects have changed.|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-protectionoptionschanged-member)|Specifies if the `WorksheetProtectionOptions` have changed.|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-sheetpasswordchanged-member)|Specifies if the worksheet password has changed.|
