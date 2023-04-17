| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[formatStaleValues](/javascript/api/excel/excel.application#excel-excel-application-formatstalevalues-member)|Specifies whether the Format Stale Values option within Calculation Options is enabled or disabled.|
|[Base64EncodedImage](/javascript/api/excel/excel.base64encodedimage)|[data](/javascript/api/excel/excel.base64encodedimage#excel-excel-base64encodedimage-data-member)|The base64 string encoding.|
||[type](/javascript/api/excel/excel.base64encodedimage#excel-excel-base64encodedimage-type-member)|The file type of the encoded image.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataRange()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatarange-member(1))|Gets the data source of the whole chart.|
||[getDataRangeOrNullObject()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatarangeornullobject-member(1))|Gets the data source of the whole chart.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Excel.EmailIdentity)](/javascript/api/excel/excel.comment#excel-excel-comment-assigntask-member(1))|Assigns the task attached to the comment to the given user as an assignee.|
||[getTask()](/javascript/api/excel/excel.comment#excel-excel-comment-gettask-member(1))|Gets the task associated with this comment.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#excel-excel-comment-gettaskornullobject-member(1))|Gets the task associated with this comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Excel.EmailIdentity)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-assigntask-member(1))|Assigns the task attached to the comment to the given user as the sole assignee.|
||[getTask()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettask-member(1))|Gets the task associated with this comment reply's thread.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettaskornullobject-member(1))|Gets the task associated with this comment reply's thread.|
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
||[commentId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-commentid-member)|Represents the ID of the `comment` or `commentReply` to which the task change is anchored.|
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
|[DocumentTaskChangeProperties](/javascript/api/excel/excel.documenttaskchangeproperties)|[assignee](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-assignee-member)|Represents the user assigned to the task for an `assign` change action, or the user unassigned from the task for an `unassign` change action.|
||[changedBy](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-changedby-member)|Represents the identity of the user who made the task change.|
||[commentId](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-commentid-member)|Represents the ID of the `comment` or `commentReply` to which the task change is anchored.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-createddatetime-member)|Represents the creation date and time of the task change record.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-duedatetime-member)|Represents the task's due date and time.|
||[id](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-id-member)|The unique GUID of the task change.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-percentcomplete-member)|Represents the task's completion percentage.|
||[priority](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-priority-member)|Represents the task's priority.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-startdatetime-member)|Represents the task's start date and time.|
||[title](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-title-member)|Represents the task's title.|
||[type](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-type-member)|Represents the action type of the task change record.|
||[undoChangeId](/javascript/api/excel/excel.documenttaskchangeproperties#excel-excel-documenttaskchangeproperties-undochangeid-member)|Represents the `DocumentTaskChange.id` property that was undone for the `undo` change action.|
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
|[EntityArrayCardLayout](/javascript/api/excel/excel.entityarraycardlayout)|[arrayProperty](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-arrayproperty-member)|Represents name of the property that contains the array shown in the card.|
||[columnsToReport](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-columnstoreport-member)|Represents the count of columns which the card claims are in the array.|
||[displayName](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-displayname-member)|Represents name of the property that contains the array shown in the card.|
||[firstRowIsHeader](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-firstrowisheader-member)|Represents whether the first row of the array is treated as a header.|
||[layout](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-layout-member)|Represents the type of this layout.|
||[rowsToReport](/javascript/api/excel/excel.entityarraycardlayout#excel-excel-entityarraycardlayout-rowstoreport-member)|Represents the count of rows which the card claims are in the array.|
|[EntityCardLayout](/javascript/api/excel/excel.entitycardlayout)|[layout](/javascript/api/excel/excel.entitycardlayout#excel-excel-entitycardlayout-layout-member)|Represents the type of this layout.|
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
|[LocalImageCellValue](/javascript/api/excel/excel.localimagecellvalue)|[altText](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-alttext-member)|Represents the alternate text used in accessibility scenarios to describe what the image represents.|
||[attribution](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-attribution-member)|Represents attribution information to describe the source and license requirements for this image.|
||[basicType](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[image](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-image-member)|Represents the image itself, either cached or encoded.|
||[provider](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-provider-member)|Represents information that describes the entity or individual who provided the image.|
||[type](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-type-member)|Represents the type of this cell value.|
|[LocalImageCellValueCacheId](/javascript/api/excel/excel.localimagecellvaluecacheid)|[cacheUid](/javascript/api/excel/excel.localimagecellvaluecacheid#excel-excel-localimagecellvaluecacheid-cacheuid-member)|Represents the image's UID as it appears in the cache.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemornullobject-member(1))|Gets a sheet view using its name.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcell-member(1))|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-pivotstyle-member)|The style applied to the PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setstyle-member(1))|Sets the style applied to the PivotTable.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-refreshmode-member)|The linked data type refresh mode.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-serviceid-member)|The unique ID of the object whose refresh mode was changed.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-type-member)|Gets the type of the event.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[refreshed](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-refreshed-member)|Indicates if the request to refresh was successful.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-serviceid-member)|The unique ID of the object whose refresh request was completed.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-type-member)|Gets the type of the event.|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-warnings-member)|An array that contains any warnings generated from the refresh request.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1))|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet.|
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
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkeddatatypes-member)|Returns a collection of linked data types that are part of the workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#excel-excel-workbook-showpivotfieldlist-member)|Specifies whether the PivotTable's field list pane is shown at the workbook level.|
||[tasks](/javascript/api/excel/excel.workbook#excel-excel-workbook-tasks-member)|Returns a collection of tasks that are present in the workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#excel-excel-workbook-use1904datesystem-member)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member)|Occurs when a filter is applied on a specific worksheet.|
||[tasks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tasks-member)|Returns a collection of tasks that are present in the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-addfrombase64-member(1))|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member)|Occurs when any worksheet's filter is applied in the workbook.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-worksheetid-member)|Gets the ID of the worksheet in which the filter is applied.|
