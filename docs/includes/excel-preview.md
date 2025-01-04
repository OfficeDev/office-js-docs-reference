| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[formatStaleValues](/javascript/api/excel/excel.application#excel-excel-application-formatstalevalues-member)|Specifies whether the Format Stale Values option within Calculation Options is enabled or disabled.|
|[Base64EncodedImage](/javascript/api/excel/excel.base64encodedimage)|[data](/javascript/api/excel/excel.base64encodedimage#excel-excel-base64encodedimage-data-member)|The Base64-encoded string.|
||[type](/javascript/api/excel/excel.base64encodedimage#excel-excel-base64encodedimage-type-member)|The file type of the Base64-encoded image.|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errorsubtype-member)|Represents the type of `BlockedErrorCellValue`.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errorsubtype-member)|Represents the type of `BusyErrorCellValue`.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errorsubtype-member)|Represents the type of `CalcErrorCellValue`.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataRange()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatarange-member(1))|Gets the data source of the whole chart.|
||[getDataRangeOrNullObject()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatarangeornullobject-member(1))|Gets the data source of the whole chart.|
|[CheckboxCellControl](/javascript/api/excel/excel.checkboxcellcontrol)|[type](/javascript/api/excel/excel.checkboxcellcontrol#excel-excel-checkboxcellcontrol-type-member)|Represents an interactable control inside of a cell.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Excel.EmailIdentity)](/javascript/api/excel/excel.comment#excel-excel-comment-assigntask-member(1))|Assigns the task attached to the comment to the given user as an assignee.|
||[getTask()](/javascript/api/excel/excel.comment#excel-excel-comment-gettask-member(1))|Gets the task associated with this comment.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#excel-excel-comment-gettaskornullobject-member(1))|Gets the task associated with this comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Excel.EmailIdentity)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-assigntask-member(1))|Assigns the task attached to the comment to the given user as the sole assignee.|
||[getTask()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettask-member(1))|Gets the task associated with this comment reply's thread.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettaskornullobject-member(1))|Gets the task associated with this comment reply's thread.|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errorsubtype-member)|Represents the type of `ConnectErrorCellValue`.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[shortDateTimePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-shortdatetimepattern-member)|Gets the format string for a short date and time value.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assign(assignee: Excel.EmailIdentity)](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-assign-member(1))|Adds the given user to the list of assignees attached to the task.|
||[assignees](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-assignees-member)|Returns a collection of assignees of the task.|
||[changes](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-changes-member)|Gets the change records of the task.|
||[comment](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-comment-member)|Gets the comment associated with the task.|
||[completedBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completedby-member)|Gets the most recent user to have completed the task.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completeddatetime-member)|Gets the date and time that the task was completed.|
||[createdBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createdby-member)|Gets the user who created the task.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createddatetime-member)|Gets the date and time that the task was created.|
||[id](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-id-member)|Gets the ID of the task.|
||[percentComplete](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-percentcomplete-member)|Specifies the completion percentage of the task.|
||[priority](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-priority-member)|Specifies the priority of the task.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-startandduedatetime-member)|Gets or sets the date and time the task should start and is due.|
||[title](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-title-member)|Specifies title of the task.|
||[unassign(assignee: Excel.EmailIdentity)](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-unassign-member(1))|Removes the given user from the list of assignees attached to the task.|
||[unassignAll()](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-unassignall-member(1))|Removes all users from the list of assignees attached to the task.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-assignee-member)|Represents the user assigned to the task for an `assign` change action, or the user unassigned from the task for an `unassign` change action.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-changedby-member)|Represents the identity of the user who made the task change.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-commentid-member)|Represents the ID of the comment or comment reply to which the task change is anchored.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-createddatetime-member)|Represents the creation date and time of the task change record.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-duedatetime-member)|Represents the task's due date and time.|
||[id](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-id-member)|The unique GUID of the task change.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-percentcomplete-member)|Represents the task's completion percentage.|
||[priority](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-priority-member)|Represents the task's priority.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-startdatetime-member)|Represents the task's start date and time.|
||[title](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-title-member)|Represents the task's title.|
||[type](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-type-member)|Represents the action type of the task change record.|
||[undoChangeId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-undochangeid-member)|Represents the `DocumentTaskChange.id` property that was undone for the `undo` change action.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getcount-member(1))|Gets the number of change records in the collection for the task.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getitemat-member(1))|Gets a task change record by using its index in the collection.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-items-member)|Gets the loaded child items in this collection.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getcount-member(1))|Gets the number of tasks in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitem-member(1))|Gets a task using its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemat-member(1))|Gets a task by its index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemornullobject-member(1))|Gets a task using its ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-items-member)|Gets the loaded child items in this collection.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-duedatetime-member)|Gets the date and time that the task is due.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-startdatetime-member)|Gets the date and time that the task should start.|
|[EmailIdentity](/javascript/api/excel/excel.emailidentity)|[displayName](/javascript/api/excel/excel.emailidentity#excel-excel-emailidentity-displayname-member)|Represents the user's display name.|
||[email](/javascript/api/excel/excel.emailidentity#excel-excel-emailidentity-email-member)|Represents the user's email.|
||[id](/javascript/api/excel/excel.emailidentity#excel-excel-emailidentity-id-member)|Represents the user's unique ID.|
|[EmptyCellControl](/javascript/api/excel/excel.emptycellcontrol)|[type](/javascript/api/excel/excel.emptycellcontrol#excel-excel-emptycellcontrol-type-member)||
|[EntityArrayCardLayout](/javascript/api/excel/excel.entityarraycardlayout)|[arrayProperty](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-arrayproperty-member)|Represents name of the property that contains the array shown in the card.|
||[columnsToReport](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-columnstoreport-member)|Represents the count of columns which the card claims are in the array.|
||[displayName](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-displayname-member)|Represents name of the property that contains the array shown in the card.|
||[firstRowIsHeader](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-firstrowisheader-member)|Represents whether the first row of the array is treated as a header.|
||[layout](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-layout-member)|Represents the type of this layout.|
||[rowsToReport](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-rowstoreport-member)|Represents the count of rows which the card claims are in the array.|
|[EntityCardLayout](/javascript/api/excel/excel.entitycardlayout)|[layout](/javascript/api/excel/excel.entitycardlayout#excel-excel-entitycardlayout-layout-member)|Represents the type of this layout.|
|[ExternalCodeServiceObjectCellValue](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue)|[Python_str](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-python_str-member)|Represents the output of str() function when used on this object.|
||[Python_type](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-python_type-member)|Represents the full type name of this object.|
||[Python_typeName](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-python_typename-member)|Represents the short type name of this object.|
||[basicType](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[language](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-language-member)|Represents the runtime language of this external code service.|
||[preview](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-preview-member)|Represents the preview value shown in the cell.|
||[provider](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-provider-member)|Represents information about the service that provided the data in this `ExternalCodeServiceObjectCellValue`.|
||[type](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-type-member)|Represents the type of this cell value.|
|[Identity](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#excel-excel-identity-displayname-member)|Represents the user's display name.|
||[id](/javascript/api/excel/excel.identity#excel-excel-identity-id-member)|Represents the user's unique ID.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-dataprovider-member)|The name of the data provider for the linked data type.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-lastrefreshed-member)|The local time-zone date and time since the workbook was opened when the linked data type was last refreshed.|
||[name](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-name-member)|The name of the linked data type.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-periodicrefreshinterval-member)|The frequency, in seconds, at which the linked data type is refreshed if `refreshMode` is set to "Periodic".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-refreshmode-member)|The mechanism by which the data for the linked data type is retrieved.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestrefresh-member(1))|Makes a request to refresh the linked data type.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestsetrefreshmode-member(1))|Makes a request to change the refresh mode for this linked data type.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-serviceid-member)|The unique ID of the linked data type.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-supportedrefreshmodes-member)|Returns an array with all the refresh modes supported by the linked data type.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-serviceid-member)|The unique ID of the new linked data type.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-type-member)|Gets the type of the event.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getcount-member(1))|Gets the number of linked data types in the collection.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitem-member(1))|Gets a linked data type by service ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemat-member(1))|Gets a linked data type by its index in the collection.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemornullobject-member(1))|Gets a linked data type by ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-items-member)|Gets the loaded child items in this collection.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-requestrefreshall-member(1))|Makes a request to refresh all the linked data types in the collection.|
|[LocalImage](/javascript/api/excel/excel.localimage)|[getBase64EncodedImageData(cacheUid: string)](/javascript/api/excel/excel.localimage#excel-excel-localimage-getbase64encodedimagedata-member(1))|Gets the Base64-encoded image data stored in the shared image cache with the cache unique identifier (UID).|
|[LocalImageCellValue](/javascript/api/excel/excel.localimagecellvalue)|[altText](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-alttext-member)|Represents the alternate text used in accessibility scenarios to describe what the image represents.|
||[attribution](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-attribution-member)|Represents attribution information to describe the source and license requirements for this image.|
||[basicType](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[image](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-image-member)|Represents the image itself, either cached or encoded.|
||[provider](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-provider-member)|Represents information that describes the entity or individual who provided the image.|
||[type](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-type-member)|Represents the type of this cell value.|
|[LocalImageCellValueCacheId](/javascript/api/excel/excel.localimagecellvaluecacheid)|[cachedUid](/javascript/api/excel/excel.localimagecellvaluecacheid#excel-excel-localimagecellvaluecacheid-cacheduid-member)|Represents the image's UID as it appears in the cache.|
|[MixedCellControl](/javascript/api/excel/excel.mixedcellcontrol)|[type](/javascript/api/excel/excel.mixedcellcontrol#excel-excel-mixedcellcontrol-type-member)||
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-errorsubtype-member)|Represents the type of `NameErrorCellValue`.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemornullobject-member(1))|Gets a sheet view using its name.|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-errorsubtype-member)|Represents the type of `NotAvailableErrorCellValue`.|
|[Note](/javascript/api/excel/excel.note)|[authorName](/javascript/api/excel/excel.note#excel-excel-note-authorname-member)|Gets the author of the note.|
||[content](/javascript/api/excel/excel.note#excel-excel-note-content-member)|Gets or sets the text of the note.|
||[delete()](/javascript/api/excel/excel.note#excel-excel-note-delete-member(1))|Deletes the note.|
||[getLocation()](/javascript/api/excel/excel.note#excel-excel-note-getlocation-member(1))|Gets the cell where this note is located.|
||[height](/javascript/api/excel/excel.note#excel-excel-note-height-member)|Specifies the height of the note.|
||[visible](/javascript/api/excel/excel.note#excel-excel-note-visible-member)|Specifies the visibility of the note.|
||[width](/javascript/api/excel/excel.note#excel-excel-note-width-member)|Specifies the width of the note.|
|[NoteCollection](/javascript/api/excel/excel.notecollection)|[add(cellAddress: Range \| string, content: any)](/javascript/api/excel/excel.notecollection#excel-excel-notecollection-add-member(1))|Adds a new note with the given content on the given cell.|
||[getCount()](/javascript/api/excel/excel.notecollection#excel-excel-notecollection-getcount-member(1))|Gets the number of notes in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.notecollection#excel-excel-notecollection-getitemat-member(1))|Gets a note object by its index in the collection.|
||[items](/javascript/api/excel/excel.notecollection#excel-excel-notecollection-items-member)|Gets the loaded child items in this collection.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcell-member(1))|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-pivotstyle-member)|The style applied to the PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setstyle-member(1))|Sets the style applied to the PivotTable.|
|[PythonErrorCellValue](/javascript/api/excel/excel.pythonerrorcellvalue)|[basicType](/javascript/api/excel/excel.pythonerrorcellvalue#excel-excel-pythonerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.pythonerrorcellvalue#excel-excel-pythonerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.pythonerrorcellvalue#excel-excel-pythonerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.pythonerrorcellvalue#excel-excel-pythonerrorcellvalue-type-member)|Represents the type of this cell value.|
|[Query](/javascript/api/excel/excel.query)|[delete()](/javascript/api/excel/excel.query#excel-excel-query-delete-member(1))|Deletes the query and associated connection.|
||[refresh()](/javascript/api/excel/excel.query#excel-excel-query-refresh-member(1))|Refreshes the query.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[refreshAll()](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-refreshall-member(1))|Refresh all queries.|
|[Range](/javascript/api/excel/excel.range)|[clearOrResetContents()](/javascript/api/excel/excel.range#excel-excel-range-clearorresetcontents-member(1))|Clears the values of the cells in the range, with special consideration given to cells containing controls.|
||[control](/javascript/api/excel/excel.range#excel-excel-range-control-member)|Accesses the cell control applied to this range.|
||[getDisplayedCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getdisplayedcellproperties-member(1))|Returns a 2D array, encapsulating the display data for each cell's font, fill, borders, alignment, and other properties.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[clearOrResetContents()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-clearorresetcontents-member(1))|Clears the values of the cells in the ranges, with special consideration given to cells containing controls.|
||[select()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-select-member(1))|Selects the specified range areas in the Excel UI.|
|[RangeTextRun](/javascript/api/excel/excel.rangetextrun)|[font](/javascript/api/excel/excel.rangetextrun#excel-excel-rangetextrun-font-member)||
||[text](/javascript/api/excel/excel.rangetextrun#excel-excel-rangetextrun-text-member)||
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[errorSubType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errorsubtype-member)|Represents the type of `RefErrorCellValue`.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-refreshmode-member)|The linked data type refresh mode.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-serviceid-member)|The unique ID of the object whose refresh mode was changed.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-type-member)|Gets the type of the event.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[refreshed](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-refreshed-member)|Indicates if the request to refresh was successful.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-serviceid-member)|The unique ID of the object whose refresh request was completed.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-type-member)|Gets the type of the event.|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-warnings-member)|An array that contains any warnings generated from the refresh request.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[textRuns](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-textruns-member)|Represents the `textRuns` property.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addLocalImageReference(address: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addlocalimagereference-member(1))|Creates a reference for the local image stored in the cell address and displays it as a floating shape over cells.|
||[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1))|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#excel-excel-slicer-nameinformula-member)|Represents the slicer name used in the formula.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#excel-excel-slicer-setstyle-member(1))|Sets the style applied to the slicer.|
||[slicerStyle](/javascript/api/excel/excel.slicer#excel-excel-slicer-slicerstyle-member)|The style applied to the slicer.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#excel-excel-table-clearstyle-member(1))|Changes the table to use the default table style.|
||[onFiltered](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member)|Occurs when a filter is applied on a specific table.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#excel-excel-table-setstyle-member(1))|Sets the style applied to the table.|
||[tableStyle](/javascript/api/excel/excel.table#excel-excel-table-tablestyle-member)|The style applied to the table.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member)|Occurs when a filter is applied on any table in a workbook, or a worksheet.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-tableid-member)|Gets the ID of the table in which the filter is applied.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-worksheetid-member)|Gets the ID of the worksheet which contains the table.|
|[TimeoutErrorCellValue](/javascript/api/excel/excel.timeouterrorcellvalue)|[basicType](/javascript/api/excel/excel.timeouterrorcellvalue#excel-excel-timeouterrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.timeouterrorcellvalue#excel-excel-timeouterrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.timeouterrorcellvalue#excel-excel-timeouterrorcellvalue-errorsubtype-member)|Represents the type of `TimeoutErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.timeouterrorcellvalue#excel-excel-timeouterrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.timeouterrorcellvalue#excel-excel-timeouterrorcellvalue-type-member)|Represents the type of this cell value.|
|[UnknownCellControl](/javascript/api/excel/excel.unknowncellcontrol)|[type](/javascript/api/excel/excel.unknowncellcontrol#excel-excel-unknowncellcontrol-type-member)||
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errorsubtype-member)|Represents the type of `ValueErrorCellValue`.|
|[Workbook](/javascript/api/excel/excel.workbook)|[externalCodeServiceTimeout](/javascript/api/excel/excel.workbook#excel-excel-workbook-externalcodeservicetimeout-member)|Specifies the maximum length of time, in seconds, allotted for a formula that depends on an external code service to complete.|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkeddatatypes-member)|Returns a collection of linked data types that are part of the workbook.|
||[localImage](/javascript/api/excel/excel.workbook#excel-excel-workbook-localimage-member)|Returns the `LocalImage` object associated with the workbook.|
||[notes](/javascript/api/excel/excel.workbook#excel-excel-workbook-notes-member)|Returns a collection of all the notes objects in the workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#excel-excel-workbook-showpivotfieldlist-member)|Specifies whether the PivotTable's field list pane is shown at the workbook level.|
||[tasks](/javascript/api/excel/excel.workbook#excel-excel-workbook-tasks-member)|Returns a collection of tasks that are present in the workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#excel-excel-workbook-use1904datesystem-member)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[notes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-notes-member)|Returns a collection of all the notes objects in the worksheet.|
||[onFiltered](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member)|Occurs when a filter is applied on a specific worksheet.|
||[tasks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tasks-member)|Returns a collection of tasks that are present in the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-addfrombase64-member(1))|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member)|Occurs when any worksheet's filter is applied in the workbook.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-worksheetid-member)|Gets the ID of the worksheet in which the filter is applied.|
