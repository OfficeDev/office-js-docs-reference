| Class | Fields | Description |
|:---|:---|:---|
|[AllowEditRange](/.alloweditrange)|[address](/.alloweditrange#excel-javascript/api/excel/-alloweditrange-address-member)|Specifies the range associated with the object.|
||[delete()](/.alloweditrange#excel-javascript/api/excel/-alloweditrange-delete-member(1))|Deletes the object from the `AllowEditRangeCollection`.|
||[isPasswordProtected](/.alloweditrange#excel-javascript/api/excel/-alloweditrange-ispasswordprotected-member)|Specifies if the object is password protected.|
||[pauseProtection(password?: string)](/.alloweditrange#excel-javascript/api/excel/-alloweditrange-pauseprotection-member(1))|Pauses worksheet protection for the object for the user in the current session.|
||[setPassword(password?: string)](/.alloweditrange#excel-javascript/api/excel/-alloweditrange-setpassword-member(1))|Changes the password associated with the object.|
||[title](/.alloweditrange#excel-javascript/api/excel/-alloweditrange-title-member)|Specifies the title of the object.|
|[AllowEditRangeCollection](/.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel.AllowEditRangeOptions)](/.alloweditrangecollection#excel-javascript/api/excel/-alloweditrangecollection-add-member(1))|Adds an `AllowEditRange` object to the worksheet.|
||[getCount()](/.alloweditrangecollection#excel-javascript/api/excel/-alloweditrangecollection-getcount-member(1))|Returns the number of `AllowEditRange` objects in the collection.|
||[getItem(key: string)](/.alloweditrangecollection#excel-javascript/api/excel/-alloweditrangecollection-getitem-member(1))|Gets the `AllowEditRange` object by its title.|
||[getItemAt(index: number)](/.alloweditrangecollection#excel-javascript/api/excel/-alloweditrangecollection-getitemat-member(1))|Returns an `AllowEditRange` object by its index in the collection.|
||[getItemOrNullObject(key: string)](/.alloweditrangecollection#excel-javascript/api/excel/-alloweditrangecollection-getitemornullobject-member(1))|Gets the `AllowEditRange` object by its title.|
||[items](/.alloweditrangecollection#excel-javascript/api/excel/-alloweditrangecollection-items-member)|Gets the loaded child items in this collection.|
||[pauseProtection(password: string)](/.alloweditrangecollection#excel-javascript/api/excel/-alloweditrangecollection-pauseprotection-member(1))|Pauses worksheet protection for all `AllowEditRange` objects found in this worksheet that have the given password for the user in the current session.|
|[AllowEditRangeOptions](/.alloweditrangeoptions)|[password](/.alloweditrangeoptions#excel-javascript/api/excel/-alloweditrangeoptions-password-member)|The password associated with the `AllowEditRange`.|
|[LinkedWorkbook](/.linkedworkbook)|[breakLinks()](/.linkedworkbook#excel-javascript/api/excel/-linkedworkbook-breaklinks-member(1))|Makes a request to break the links pointing to the linked workbook.|
||[id](/.linkedworkbook#excel-javascript/api/excel/-linkedworkbook-id-member)|The original URL pointing to the linked workbook.|
||[refresh()](/.linkedworkbook#excel-javascript/api/excel/-linkedworkbook-refresh-member(1))|Makes a request to refresh the data retrieved from the linked workbook.|
|[LinkedWorkbookCollection](/.linkedworkbookcollection)|[breakAllLinks()](/.linkedworkbookcollection#excel-javascript/api/excel/-linkedworkbookcollection-breakalllinks-member(1))|Breaks all the links to the linked workbooks.|
||[getItem(key: string)](/.linkedworkbookcollection#excel-javascript/api/excel/-linkedworkbookcollection-getitem-member(1))|Gets information about a linked workbook by its URL.|
||[getItemOrNullObject(key: string)](/.linkedworkbookcollection#excel-javascript/api/excel/-linkedworkbookcollection-getitemornullobject-member(1))|Gets information about a linked workbook by its URL.|
||[items](/.linkedworkbookcollection#excel-javascript/api/excel/-linkedworkbookcollection-items-member)|Gets the loaded child items in this collection.|
||[refreshAll()](/.linkedworkbookcollection#excel-javascript/api/excel/-linkedworkbookcollection-refreshall-member(1))|Makes a request to refresh all the workbook links.|
||[workbookLinksRefreshMode](/.linkedworkbookcollection#excel-javascript/api/excel/-linkedworkbookcollection-workbooklinksrefreshmode-member)|Represents the update mode of the workbook links.|
|[NamedSheetView](/.namedsheetview)|[activate()](/.namedsheetview#excel-javascript/api/excel/-namedsheetview-activate-member(1))|Activates this sheet view.|
||[delete()](/.namedsheetview#excel-javascript/api/excel/-namedsheetview-delete-member(1))|Removes the sheet view from the worksheet.|
||[duplicate(name?: string)](/.namedsheetview#excel-javascript/api/excel/-namedsheetview-duplicate-member(1))|Creates a copy of this sheet view.|
||[name](/.namedsheetview#excel-javascript/api/excel/-namedsheetview-name-member)|Gets or sets the name of the sheet view.|
|[NamedSheetViewCollection](/.namedsheetviewcollection)|[add(name: string)](/.namedsheetviewcollection#excel-javascript/api/excel/-namedsheetviewcollection-add-member(1))|Creates a new sheet view with the given name.|
||[enterTemporary()](/.namedsheetviewcollection#excel-javascript/api/excel/-namedsheetviewcollection-entertemporary-member(1))|Creates and activates a new temporary sheet view.|
||[exit()](/.namedsheetviewcollection#excel-javascript/api/excel/-namedsheetviewcollection-exit-member(1))|Exits the currently active sheet view.|
||[getActive()](/.namedsheetviewcollection#excel-javascript/api/excel/-namedsheetviewcollection-getactive-member(1))|Gets the worksheet's currently active sheet view.|
||[getCount()](/.namedsheetviewcollection#excel-javascript/api/excel/-namedsheetviewcollection-getcount-member(1))|Gets the number of sheet views in this worksheet.|
||[getItem(key: string)](/.namedsheetviewcollection#excel-javascript/api/excel/-namedsheetviewcollection-getitem-member(1))|Gets a sheet view using its name.|
||[getItemAt(index: number)](/.namedsheetviewcollection#excel-javascript/api/excel/-namedsheetviewcollection-getitemat-member(1))|Gets a sheet view by its index in the collection.|
||[items](/.namedsheetviewcollection#excel-javascript/api/excel/-namedsheetviewcollection-items-member)|Gets the loaded child items in this collection.|
|[TableRowCollection](/.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/.tablerowcollection#excel-javascript/api/excel/-tablerowcollection-deleterows-member(1))|Delete multiple rows from a table.|
||[deleteRowsAt(index: number, count?: number)](/.tablerowcollection#excel-javascript/api/excel/-tablerowcollection-deleterowsat-member(1))|Delete a specified number of rows from a table, starting at a given index.|
|[Workbook](/.workbook)|[linkedWorkbooks](/.workbook#excel-javascript/api/excel/-workbook-linkedworkbooks-member)|Returns a collection of linked workbooks.|
|[Worksheet](/.worksheet)|[namedSheetViews](/.worksheet#excel-javascript/api/excel/-worksheet-namedsheetviews-member)|Returns a collection of sheet views that are present in the worksheet.|
|[WorksheetProtection](/.worksheetprotection)|[allowEditRanges](/.worksheetprotection#excel-javascript/api/excel/-worksheetprotection-alloweditranges-member)|Specifies the `AllowEditRangeCollection` object found in this worksheet.|
||[canPauseProtection](/.worksheetprotection#excel-javascript/api/excel/-worksheetprotection-canpauseprotection-member)|Specifies if protection can be paused for this worksheet.|
||[checkPassword(password?: string)](/.worksheetprotection#excel-javascript/api/excel/-worksheetprotection-checkpassword-member(1))|Specifies if the password can be used to unlock worksheet protection.|
||[isPasswordProtected](/.worksheetprotection#excel-javascript/api/excel/-worksheetprotection-ispasswordprotected-member)|Specifies if the sheet is password protected.|
||[isPaused](/.worksheetprotection#excel-javascript/api/excel/-worksheetprotection-ispaused-member)|Specifies if worksheet protection is paused.|
||[pauseProtection(password?: string)](/.worksheetprotection#excel-javascript/api/excel/-worksheetprotection-pauseprotection-member(1))|Pauses worksheet protection for the given worksheet object for the user in the current session.|
||[resumeProtection()](/.worksheetprotection#excel-javascript/api/excel/-worksheetprotection-resumeprotection-member(1))|Resumes worksheet protection for the given worksheet object for the user in a given session.|
||[savedOptions](/.worksheetprotection#excel-javascript/api/excel/-worksheetprotection-savedoptions-member)|Specifies the protection options saved in the worksheet.|
||[setPassword(password?: string)](/.worksheetprotection#excel-javascript/api/excel/-worksheetprotection-setpassword-member(1))|Changes the password associated with the `WorksheetProtection` object.|
||[updateOptions(options: Excel.WorksheetProtectionOptions)](/.worksheetprotection#excel-javascript/api/excel/-worksheetprotection-updateoptions-member(1))|Change the worksheet protection options associated with the `WorksheetProtection` object.|
|[WorksheetProtectionChangedEventArgs](/.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/.worksheetprotectionchangedeventargs#excel-javascript/api/excel/-worksheetprotectionchangedeventargs-alloweditrangeschanged-member)|Specifies if any of the `AllowEditRange` objects have changed.|
||[protectionOptionsChanged](/.worksheetprotectionchangedeventargs#excel-javascript/api/excel/-worksheetprotectionchangedeventargs-protectionoptionschanged-member)|Specifies if the `WorksheetProtectionOptions` have changed.|
||[sheetPasswordChanged](/.worksheetprotectionchangedeventargs#excel-javascript/api/excel/-worksheetprotectionchangedeventargs-sheetpasswordchanged-member)|Specifies if the worksheet password has changed.|
