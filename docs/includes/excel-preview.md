| Class | Fields | Description |
|:---|:---|:---|
|[Application](/.application)|[formatStaleValues](/.application#excel-javascript/api/excel/-application-formatstalevalues-member)|Specifies whether the Format Stale Values option within Calculation Options is enabled or disabled.|
|[Base64EncodedImage](/.base64encodedimage)|[data](/.base64encodedimage#excel-javascript/api/excel/-base64encodedimage-data-member)|The Base64-encoded string.|
||[type](/.base64encodedimage#excel-javascript/api/excel/-base64encodedimage-type-member)|The file type of the Base64-encoded image.|
|[BlockedErrorCellValue](/.blockederrorcellvalue)|[errorSubType](/.blockederrorcellvalue#excel-javascript/api/excel/-blockederrorcellvalue-errorsubtype-member)|Represents the type of `BlockedErrorCellValue`.|
|[BooleanCellValue](/.booleancellvalue)|[type](/.booleancellvalue#excel-javascript/api/excel/-booleancellvalue-type-member)|Represents the type of this cell value.|
|[BusyErrorCellValue](/.busyerrorcellvalue)|[errorSubType](/.busyerrorcellvalue#excel-javascript/api/excel/-busyerrorcellvalue-errorsubtype-member)|Represents the type of `BusyErrorCellValue`.|
|[CalcErrorCellValue](/.calcerrorcellvalue)|[errorSubType](/.calcerrorcellvalue#excel-javascript/api/excel/-calcerrorcellvalue-errorsubtype-member)|Represents the type of `CalcErrorCellValue`.|
|[Chart](/.chart)|[getDataRange()](/.chart#excel-javascript/api/excel/-chart-getdatarange-member(1))|Gets the data source of the whole chart.|
||[getDataRangeOrNullObject()](/.chart#excel-javascript/api/excel/-chart-getdatarangeornullobject-member(1))|Gets the data source of the whole chart.|
|[Comment](/.comment)|[assignTask(assignee: Excel.EmailIdentity)](/.comment#excel-javascript/api/excel/-comment-assigntask-member(1))|Assigns the task attached to the comment to the given user as an assignee.|
||[getTask()](/.comment#excel-javascript/api/excel/-comment-gettask-member(1))|Gets the task associated with this comment.|
||[getTaskOrNullObject()](/.comment#excel-javascript/api/excel/-comment-gettaskornullobject-member(1))|Gets the task associated with this comment.|
|[CommentReply](/.commentreply)|[assignTask(assignee: Excel.EmailIdentity)](/.commentreply#excel-javascript/api/excel/-commentreply-assigntask-member(1))|Assigns the task attached to the comment to the given user as the sole assignee.|
||[getTask()](/.commentreply#excel-javascript/api/excel/-commentreply-gettask-member(1))|Gets the task associated with this comment reply's thread.|
||[getTaskOrNullObject()](/.commentreply#excel-javascript/api/excel/-commentreply-gettaskornullobject-member(1))|Gets the task associated with this comment reply's thread.|
|[ConnectErrorCellValue](/.connecterrorcellvalue)|[errorSubType](/.connecterrorcellvalue#excel-javascript/api/excel/-connecterrorcellvalue-errorsubtype-member)|Represents the type of `ConnectErrorCellValue`.|
|[DatetimeFormatInfo](/.datetimeformatinfo)|[shortDateTimePattern](/.datetimeformatinfo#excel-javascript/api/excel/-datetimeformatinfo-shortdatetimepattern-member)|Gets the format string for a short date and time value.|
|[DocumentTask](/.documenttask)|[assign(assignee: Excel.EmailIdentity)](/.documenttask#excel-javascript/api/excel/-documenttask-assign-member(1))|Adds the given user to the list of assignees attached to the task.|
||[assignees](/.documenttask#excel-javascript/api/excel/-documenttask-assignees-member)|Returns a collection of assignees of the task.|
||[changes](/.documenttask#excel-javascript/api/excel/-documenttask-changes-member)|Gets the change records of the task.|
||[comment](/.documenttask#excel-javascript/api/excel/-documenttask-comment-member)|Gets the comment associated with the task.|
||[completedBy](/.documenttask#excel-javascript/api/excel/-documenttask-completedby-member)|Gets the most recent user to have completed the task.|
||[completedDateTime](/.documenttask#excel-javascript/api/excel/-documenttask-completeddatetime-member)|Gets the date and time that the task was completed.|
||[createdBy](/.documenttask#excel-javascript/api/excel/-documenttask-createdby-member)|Gets the user who created the task.|
||[createdDateTime](/.documenttask#excel-javascript/api/excel/-documenttask-createddatetime-member)|Gets the date and time that the task was created.|
||[id](/.documenttask#excel-javascript/api/excel/-documenttask-id-member)|Gets the ID of the task.|
||[percentComplete](/.documenttask#excel-javascript/api/excel/-documenttask-percentcomplete-member)|Specifies the completion percentage of the task.|
||[priority](/.documenttask#excel-javascript/api/excel/-documenttask-priority-member)|Specifies the priority of the task.|
||[startAndDueDateTime](/.documenttask#excel-javascript/api/excel/-documenttask-startandduedatetime-member)|Gets or sets the date and time the task should start and is due.|
||[title](/.documenttask#excel-javascript/api/excel/-documenttask-title-member)|Specifies title of the task.|
||[unassign(assignee: Excel.EmailIdentity)](/.documenttask#excel-javascript/api/excel/-documenttask-unassign-member(1))|Removes the given user from the list of assignees attached to the task.|
||[unassignAll()](/.documenttask#excel-javascript/api/excel/-documenttask-unassignall-member(1))|Removes all users from the list of assignees attached to the task.|
|[DocumentTaskChange](/.documenttaskchange)|[assignee](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-assignee-member)|Represents the user assigned to the task for an `assign` change action, or the user unassigned from the task for an `unassign` change action.|
||[changedBy](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-changedby-member)|Represents the identity of the user who made the task change.|
||[commentId](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-commentid-member)|Represents the ID of the comment or comment reply to which the task change is anchored.|
||[createdDateTime](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-createddatetime-member)|Represents the creation date and time of the task change record.|
||[dueDateTime](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-duedatetime-member)|Represents the task's due date and time.|
||[id](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-id-member)|The unique GUID of the task change.|
||[percentComplete](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-percentcomplete-member)|Represents the task's completion percentage.|
||[priority](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-priority-member)|Represents the task's priority.|
||[startDateTime](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-startdatetime-member)|Represents the task's start date and time.|
||[title](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-title-member)|Represents the task's title.|
||[type](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-type-member)|Represents the action type of the task change record.|
||[undoChangeId](/.documenttaskchange#excel-javascript/api/excel/-documenttaskchange-undochangeid-member)|Represents the `DocumentTaskChange.id` property that was undone for the `undo` change action.|
|[DocumentTaskChangeCollection](/.documenttaskchangecollection)|[getCount()](/.documenttaskchangecollection#excel-javascript/api/excel/-documenttaskchangecollection-getcount-member(1))|Gets the number of change records in the collection for the task.|
||[getItemAt(index: number)](/.documenttaskchangecollection#excel-javascript/api/excel/-documenttaskchangecollection-getitemat-member(1))|Gets a task change record by using its index in the collection.|
||[items](/.documenttaskchangecollection#excel-javascript/api/excel/-documenttaskchangecollection-items-member)|Gets the loaded child items in this collection.|
|[DocumentTaskCollection](/.documenttaskcollection)|[getCount()](/.documenttaskcollection#excel-javascript/api/excel/-documenttaskcollection-getcount-member(1))|Gets the number of tasks in the collection.|
||[getItem(key: string)](/.documenttaskcollection#excel-javascript/api/excel/-documenttaskcollection-getitem-member(1))|Gets a task using its ID.|
||[getItemAt(index: number)](/.documenttaskcollection#excel-javascript/api/excel/-documenttaskcollection-getitemat-member(1))|Gets a task by its index in the collection.|
||[getItemOrNullObject(key: string)](/.documenttaskcollection#excel-javascript/api/excel/-documenttaskcollection-getitemornullobject-member(1))|Gets a task using its ID.|
||[items](/.documenttaskcollection#excel-javascript/api/excel/-documenttaskcollection-items-member)|Gets the loaded child items in this collection.|
|[DocumentTaskSchedule](/.documenttaskschedule)|[dueDateTime](/.documenttaskschedule#excel-javascript/api/excel/-documenttaskschedule-duedatetime-member)|Gets the date and time that the task is due.|
||[startDateTime](/.documenttaskschedule#excel-javascript/api/excel/-documenttaskschedule-startdatetime-member)|Gets the date and time that the task should start.|
|[DoubleCellValue](/.doublecellvalue)|[type](/.doublecellvalue#excel-javascript/api/excel/-doublecellvalue-type-member)|Represents the type of this cell value.|
|[EmailIdentity](/.emailidentity)|[displayName](/.emailidentity#excel-javascript/api/excel/-emailidentity-displayname-member)|Represents the user's display name.|
||[email](/.emailidentity#excel-javascript/api/excel/-emailidentity-email-member)|Represents the user's email.|
||[id](/.emailidentity#excel-javascript/api/excel/-emailidentity-id-member)|Represents the user's unique ID.|
|[EntityArrayCardLayout](/.entityarraycardlayout)|[arrayProperty](/.entityarraycardlayout#excel-javascript/api/excel/-entityarraycardlayout-arrayproperty-member)|Represents name of the property that contains the array shown in the card.|
||[columnsToReport](/.entityarraycardlayout#excel-javascript/api/excel/-entityarraycardlayout-columnstoreport-member)|Represents the count of columns which the card claims are in the array.|
||[displayName](/.entityarraycardlayout#excel-javascript/api/excel/-entityarraycardlayout-displayname-member)|Represents name of the property that contains the array shown in the card.|
||[firstRowIsHeader](/.entityarraycardlayout#excel-javascript/api/excel/-entityarraycardlayout-firstrowisheader-member)|Represents whether the first row of the array is treated as a header.|
||[layout](/.entityarraycardlayout#excel-javascript/api/excel/-entityarraycardlayout-layout-member)|Represents the type of this layout.|
||[rowsToReport](/.entityarraycardlayout#excel-javascript/api/excel/-entityarraycardlayout-rowstoreport-member)|Represents the count of rows which the card claims are in the array.|
|[EntityCardLayout](/.entitycardlayout)|[layout](/.entitycardlayout#excel-javascript/api/excel/-entitycardlayout-layout-member)|Represents the type of this layout.|
|[ExternalCodeServiceObjectCellValue](/.externalcodeserviceobjectcellvalue)|[Python_str](/.externalcodeserviceobjectcellvalue#excel-javascript/api/excel/-externalcodeserviceobjectcellvalue-python_str-member)|Represents the output of the `str()` function when used on this object.|
||[Python_type](/.externalcodeserviceobjectcellvalue#excel-javascript/api/excel/-externalcodeserviceobjectcellvalue-python_type-member)|Represents the full type name of this object.|
||[Python_typeName](/.externalcodeserviceobjectcellvalue#excel-javascript/api/excel/-externalcodeserviceobjectcellvalue-python_typename-member)|Represents the short type name of this object.|
||[basicType](/.externalcodeserviceobjectcellvalue#excel-javascript/api/excel/-externalcodeserviceobjectcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/.externalcodeserviceobjectcellvalue#excel-javascript/api/excel/-externalcodeserviceobjectcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[language](/.externalcodeserviceobjectcellvalue#excel-javascript/api/excel/-externalcodeserviceobjectcellvalue-language-member)|Represents the runtime language of this external code service.|
||[preview](/.externalcodeserviceobjectcellvalue#excel-javascript/api/excel/-externalcodeserviceobjectcellvalue-preview-member)|Represents the preview value shown in the cell.|
||[provider](/.externalcodeserviceobjectcellvalue#excel-javascript/api/excel/-externalcodeserviceobjectcellvalue-provider-member)|Represents information about the service that provided the data in this `ExternalCodeServiceObjectCellValue`.|
||[type](/.externalcodeserviceobjectcellvalue#excel-javascript/api/excel/-externalcodeserviceobjectcellvalue-type-member)|Represents the type of this cell value.|
|[Identity](/.identity)|[displayName](/.identity#excel-javascript/api/excel/-identity-displayname-member)|Represents the user's display name.|
||[id](/.identity#excel-javascript/api/excel/-identity-id-member)|Represents the user's unique ID.|
|[LinkedDataType](/.linkeddatatype)|[dataProvider](/.linkeddatatype#excel-javascript/api/excel/-linkeddatatype-dataprovider-member)|The name of the data provider for the linked data type.|
||[lastRefreshed](/.linkeddatatype#excel-javascript/api/excel/-linkeddatatype-lastrefreshed-member)|The local time-zone date and time since the workbook was opened when the linked data type was last refreshed.|
||[name](/.linkeddatatype#excel-javascript/api/excel/-linkeddatatype-name-member)|The name of the linked data type.|
||[periodicRefreshInterval](/.linkeddatatype#excel-javascript/api/excel/-linkeddatatype-periodicrefreshinterval-member)|The frequency, in seconds, at which the linked data type is refreshed if `refreshMode` is set to "Periodic".|
||[refreshMode](/.linkeddatatype#excel-javascript/api/excel/-linkeddatatype-refreshmode-member)|The mechanism by which the data for the linked data type is retrieved.|
||[requestRefresh()](/.linkeddatatype#excel-javascript/api/excel/-linkeddatatype-requestrefresh-member(1))|Makes a request to refresh the linked data type.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/.linkeddatatype#excel-javascript/api/excel/-linkeddatatype-requestsetrefreshmode-member(1))|Makes a request to change the refresh mode for this linked data type.|
||[serviceId](/.linkeddatatype#excel-javascript/api/excel/-linkeddatatype-serviceid-member)|The unique ID of the linked data type.|
||[supportedRefreshModes](/.linkeddatatype#excel-javascript/api/excel/-linkeddatatype-supportedrefreshmodes-member)|Returns an array with all the refresh modes supported by the linked data type.|
|[LinkedDataTypeAddedEventArgs](/.linkeddatatypeaddedeventargs)|[serviceId](/.linkeddatatypeaddedeventargs#excel-javascript/api/excel/-linkeddatatypeaddedeventargs-serviceid-member)|The unique ID of the new linked data type.|
||[source](/.linkeddatatypeaddedeventargs#excel-javascript/api/excel/-linkeddatatypeaddedeventargs-source-member)|Gets the source of the event.|
||[type](/.linkeddatatypeaddedeventargs#excel-javascript/api/excel/-linkeddatatypeaddedeventargs-type-member)|Gets the type of the event.|
|[LinkedDataTypeCollection](/.linkeddatatypecollection)|[getCount()](/.linkeddatatypecollection#excel-javascript/api/excel/-linkeddatatypecollection-getcount-member(1))|Gets the number of linked data types in the collection.|
||[getItem(key: number)](/.linkeddatatypecollection#excel-javascript/api/excel/-linkeddatatypecollection-getitem-member(1))|Gets a linked data type by service ID.|
||[getItemAt(index: number)](/.linkeddatatypecollection#excel-javascript/api/excel/-linkeddatatypecollection-getitemat-member(1))|Gets a linked data type by its index in the collection.|
||[getItemOrNullObject(key: number)](/.linkeddatatypecollection#excel-javascript/api/excel/-linkeddatatypecollection-getitemornullobject-member(1))|Gets a linked data type by ID.|
||[items](/.linkeddatatypecollection#excel-javascript/api/excel/-linkeddatatypecollection-items-member)|Gets the loaded child items in this collection.|
||[requestRefreshAll()](/.linkeddatatypecollection#excel-javascript/api/excel/-linkeddatatypecollection-requestrefreshall-member(1))|Makes a request to refresh all the linked data types in the collection.|
|[LocalImage](/.localimage)|[getBase64EncodedImageData(cacheUid: string)](/.localimage#excel-javascript/api/excel/-localimage-getbase64encodedimagedata-member(1))|Gets the Base64-encoded image data stored in the shared image cache with the cache unique identifier (UID).|
|[LocalImageCellValue](/.localimagecellvalue)|[altText](/.localimagecellvalue#excel-javascript/api/excel/-localimagecellvalue-alttext-member)|Represents the alternate text used in accessibility scenarios to describe what the image represents.|
||[attribution](/.localimagecellvalue#excel-javascript/api/excel/-localimagecellvalue-attribution-member)|Represents attribution information to describe the source and license requirements for this image.|
||[basicType](/.localimagecellvalue#excel-javascript/api/excel/-localimagecellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/.localimagecellvalue#excel-javascript/api/excel/-localimagecellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[image](/.localimagecellvalue#excel-javascript/api/excel/-localimagecellvalue-image-member)|Represents the image itself, either cached or encoded.|
||[provider](/.localimagecellvalue#excel-javascript/api/excel/-localimagecellvalue-provider-member)|Represents information that describes the entity or individual who provided the image.|
||[type](/.localimagecellvalue#excel-javascript/api/excel/-localimagecellvalue-type-member)|Represents the type of this cell value.|
|[LocalImageCellValueCacheId](/.localimagecellvaluecacheid)|[cachedUid](/.localimagecellvaluecacheid#excel-javascript/api/excel/-localimagecellvaluecacheid-cacheduid-member)|Represents the image's UID as it appears in the cache.|
|[NameErrorCellValue](/.nameerrorcellvalue)|[errorSubType](/.nameerrorcellvalue#excel-javascript/api/excel/-nameerrorcellvalue-errorsubtype-member)|Represents the type of `NameErrorCellValue`.|
|[NamedSheetViewCollection](/.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/.namedsheetviewcollection#excel-javascript/api/excel/-namedsheetviewcollection-getitemornullobject-member(1))|Gets a sheet view using its name.|
|[NotAvailableErrorCellValue](/.notavailableerrorcellvalue)|[errorSubType](/.notavailableerrorcellvalue#excel-javascript/api/excel/-notavailableerrorcellvalue-errorsubtype-member)|Represents the type of `NotAvailableErrorCellValue`.|
|[PivotLayout](/.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-getcell-member(1))|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies.|
||[pivotStyle](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-pivotstyle-member)|The style applied to the PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-setstyle-member(1))|Sets the style applied to the PivotTable.|
|[PivotTable](/.pivottable)|[autoRefresh](/.pivottable#excel-javascript/api/excel/-pivottable-autorefresh-member)|Specifies whether the PivotTable auto refreshes when the source data changes.|
|[PythonErrorCellValue](/.pythonerrorcellvalue)|[basicType](/.pythonerrorcellvalue#excel-javascript/api/excel/-pythonerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/.pythonerrorcellvalue#excel-javascript/api/excel/-pythonerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/.pythonerrorcellvalue#excel-javascript/api/excel/-pythonerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/.pythonerrorcellvalue#excel-javascript/api/excel/-pythonerrorcellvalue-type-member)|Represents the type of this cell value.|
|[Query](/.query)|[delete()](/.query#excel-javascript/api/excel/-query-delete-member(1))|Deletes the query and associated connection.|
||[refresh()](/.query#excel-javascript/api/excel/-query-refresh-member(1))|Refreshes the query.|
|[QueryCollection](/.querycollection)|[refreshAll()](/.querycollection#excel-javascript/api/excel/-querycollection-refreshall-member(1))|Refresh all queries.|
|[Range](/.range)|[togglePythonMarshalMode(marshalMode?: Excel.PythonMarshalMode)](/.range#excel-javascript/api/excel/-range-togglepythonmarshalmode-member(1))|Sets the marshaling mode of the Python in Excel formula =PY.|
|[RangeAreas](/.rangeareas)|||
|[RefErrorCellValue](/.referrorcellvalue)|[errorSubType](/.referrorcellvalue#excel-javascript/api/excel/-referrorcellvalue-errorsubtype-member)|Represents the type of `RefErrorCellValue`.|
|[RefreshModeChangedEventArgs](/.refreshmodechangedeventargs)|[refreshMode](/.refreshmodechangedeventargs#excel-javascript/api/excel/-refreshmodechangedeventargs-refreshmode-member)|The linked data type refresh mode.|
||[serviceId](/.refreshmodechangedeventargs#excel-javascript/api/excel/-refreshmodechangedeventargs-serviceid-member)|The unique ID of the object whose refresh mode was changed.|
||[source](/.refreshmodechangedeventargs#excel-javascript/api/excel/-refreshmodechangedeventargs-source-member)|Gets the source of the event.|
||[type](/.refreshmodechangedeventargs#excel-javascript/api/excel/-refreshmodechangedeventargs-type-member)|Gets the type of the event.|
|[RefreshRequestCompletedEventArgs](/.refreshrequestcompletedeventargs)|[refreshed](/.refreshrequestcompletedeventargs#excel-javascript/api/excel/-refreshrequestcompletedeventargs-refreshed-member)|Indicates if the request to refresh was successful.|
||[serviceId](/.refreshrequestcompletedeventargs#excel-javascript/api/excel/-refreshrequestcompletedeventargs-serviceid-member)|The unique ID of the object whose refresh request was completed.|
||[source](/.refreshrequestcompletedeventargs#excel-javascript/api/excel/-refreshrequestcompletedeventargs-source-member)|Gets the source of the event.|
||[type](/.refreshrequestcompletedeventargs#excel-javascript/api/excel/-refreshrequestcompletedeventargs-type-member)|Gets the type of the event.|
||[warnings](/.refreshrequestcompletedeventargs#excel-javascript/api/excel/-refreshrequestcompletedeventargs-warnings-member)|An array that contains any warnings generated from the refresh request.|
|[ShapeCollection](/.shapecollection)|[addLocalImageReference(address: string)](/.shapecollection#excel-javascript/api/excel/-shapecollection-addlocalimagereference-member(1))|Creates a reference for the local image stored in the cell address and displays it as a floating shape over cells.|
||[addSvg(xml: string)](/.shapecollection#excel-javascript/api/excel/-shapecollection-addsvg-member(1))|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet.|
|[Slicer](/.slicer)|[nameInFormula](/.slicer#excel-javascript/api/excel/-slicer-nameinformula-member)|Specifies the slicer name used in the formula.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/.slicer#excel-javascript/api/excel/-slicer-setstyle-member(1))|Sets the style applied to the slicer.|
||[slicerStyle](/.slicer#excel-javascript/api/excel/-slicer-slicerstyle-member)|The style applied to the slicer.|
|[StringCellValue](/.stringcellvalue)|[type](/.stringcellvalue#excel-javascript/api/excel/-stringcellvalue-type-member)|Represents the type of this cell value.|
|[Table](/.table)|[clearStyle()](/.table#excel-javascript/api/excel/-table-clearstyle-member(1))|Changes the table to use the default table style.|
||[onFiltered](/.table#excel-javascript/api/excel/-table-onfiltered-member)|Occurs when a filter is applied on a specific table.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/.table#excel-javascript/api/excel/-table-setstyle-member(1))|Sets the style applied to the table.|
||[tableStyle](/.table#excel-javascript/api/excel/-table-tablestyle-member)|The style applied to the table.|
|[TableCollection](/.tablecollection)|[onFiltered](/.tablecollection#excel-javascript/api/excel/-tablecollection-onfiltered-member)|Occurs when a filter is applied on any table in a workbook, or a worksheet.|
|[TableFilteredEventArgs](/.tablefilteredeventargs)|[tableId](/.tablefilteredeventargs#excel-javascript/api/excel/-tablefilteredeventargs-tableid-member)|Gets the ID of the table in which the filter is applied.|
||[type](/.tablefilteredeventargs#excel-javascript/api/excel/-tablefilteredeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.tablefilteredeventargs#excel-javascript/api/excel/-tablefilteredeventargs-worksheetid-member)|Gets the ID of the worksheet which contains the table.|
|[TimeoutErrorCellValue](/.timeouterrorcellvalue)|[basicType](/.timeouterrorcellvalue#excel-javascript/api/excel/-timeouterrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/.timeouterrorcellvalue#excel-javascript/api/excel/-timeouterrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/.timeouterrorcellvalue#excel-javascript/api/excel/-timeouterrorcellvalue-errorsubtype-member)|Represents the type of `TimeoutErrorCellValue`.|
||[errorType](/.timeouterrorcellvalue#excel-javascript/api/excel/-timeouterrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/.timeouterrorcellvalue#excel-javascript/api/excel/-timeouterrorcellvalue-type-member)|Represents the type of this cell value.|
|[ValueErrorCellValue](/.valueerrorcellvalue)|[errorSubType](/.valueerrorcellvalue#excel-javascript/api/excel/-valueerrorcellvalue-errorsubtype-member)|Represents the type of `ValueErrorCellValue`.|
|[Workbook](/.workbook)|[enterPreviewMode()](/.workbook#excel-javascript/api/excel/-workbook-enterpreviewmode-member(1))|Enters Scratchpad Preview Mode for the workbook, showing changes suggested by Copilot to the user.|
||[exitPreviewMode(applyChanges: boolean)](/.workbook#excel-javascript/api/excel/-workbook-exitpreviewmode-member(1))|Exits Scratchpad Preview Mode for the workbook.|
||[externalCodeServiceTimeout](/.workbook#excel-javascript/api/excel/-workbook-externalcodeservicetimeout-member)|Specifies the maximum length of time, in seconds, allotted for a formula that depends on an external code service to complete.|
||[linkedDataTypes](/.workbook#excel-javascript/api/excel/-workbook-linkeddatatypes-member)|Returns a collection of linked data types that are part of the workbook.|
||[localImage](/.workbook#excel-javascript/api/excel/-workbook-localimage-member)|Returns the `LocalImage` object associated with the workbook.|
||[showPivotFieldList](/.workbook#excel-javascript/api/excel/-workbook-showpivotfieldlist-member)|Specifies whether the PivotTable's field list pane is shown at the workbook level.|
||[tasks](/.workbook#excel-javascript/api/excel/-workbook-tasks-member)|Returns a collection of tasks that are present in the workbook.|
||[use1904DateSystem](/.workbook#excel-javascript/api/excel/-workbook-use1904datesystem-member)|True if the workbook uses the 1904 date system.|
|[Worksheet](/.worksheet)|[onFiltered](/.worksheet#excel-javascript/api/excel/-worksheet-onfiltered-member)|Occurs when a filter is applied on a specific worksheet.|
||[tasks](/.worksheet#excel-javascript/api/excel/-worksheet-tasks-member)|Returns a collection of tasks that are present in the worksheet.|
|[WorksheetCollection](/.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-addfrombase64-member(1))|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onfiltered-member)|Occurs when any worksheet's filter is applied in the workbook.|
|[WorksheetFilteredEventArgs](/.worksheetfilteredeventargs)|[type](/.worksheetfilteredeventargs#excel-javascript/api/excel/-worksheetfilteredeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetfilteredeventargs#excel-javascript/api/excel/-worksheetfilteredeventargs-worksheetid-member)|Gets the ID of the worksheet in which the filter is applied.|
