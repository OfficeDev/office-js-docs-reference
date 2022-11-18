| Class | Fields | Description |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[getDataRange()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatarange-member(1))|Gets the data source of the whole chart.|
||[getDataRangeOrNullObject()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatarangeornullobject-member(1))|Gets the data source of the whole chart.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Excel.Identity)](/javascript/api/excel/excel.comment#excel-excel-comment-assigntask-member(1))|Assigns the task attached to the comment to the given user as an assignee.|
||[getTask()](/javascript/api/excel/excel.comment#excel-excel-comment-gettask-member(1))|Gets the task associated with this comment.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#excel-excel-comment-gettaskornullobject-member(1))|Gets the task associated with this comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Excel.Identity)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-assigntask-member(1))|Assigns the task attached to the comment to the given user as the sole assignee.|
||[getTask()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettask-member(1))|Gets the task associated with this comment reply's thread.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettaskornullobject-member(1))|Gets the task associated with this comment reply's thread.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[changeRuleToCellValue(properties: Excel.ConditionalCellValueRule)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocellvalue-member(1))|Change the conditional format rule type to cell value.|
||[changeRuleToColorScale()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocolorscale-member(1))|Change the conditional format rule type to color scale.|
||[changeRuleToContainsText(properties: Excel.ConditionalTextComparisonRule)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocontainstext-member(1))|Change the conditional format rule type to text comparison.|
||[changeRuleToCustom(formula: string)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocustom-member(1))|Change the conditional format rule type to custom.|
||[changeRuleToDataBar()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletodatabar-member(1))|Change the conditional format rule type to data bar.|
||[changeRuleToIconSet()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletoiconset-member(1))|Change the conditional format rule type to icon set.|
||[changeRuleToPresetCriteria(properties: Excel.ConditionalPresetCriteriaRule)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletopresetcriteria-member(1))|Change the conditional format rule type to preset criteria.|
||[changeRuleToTopBottom(properties: Excel.ConditionalTopBottomRule)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletotopbottom-member(1))|Change the conditional format rule type to top/bottom.|
||[setRanges(ranges: Range \| RangeAreas \| string)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-setranges-member(1))|Set the ranges that the conditonal format rule is applied to.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[clearFormat()](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-clearformat-member(1))|Remove the format properties from a conditional format rule.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assign(assignee: Excel.Identity)](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-assign-member(1))|Adds the given user to the list of assignees attached to the task.|
||[assignees](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-assignees-member)|Returns a collection of assignees of the task.|
||[changes](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-changes-member)|Gets the change records of the task.|
||[comment](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-comment-member)|Gets the comment associated with the task.|
||[completedBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completedby-member)|Gets the most recent user to have completed the task.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completeddatetime-member)|Gets the date and time that the task was completed.|
||[createdBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createdby-member)|Gets the user who created the task.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createddatetime-member)|Gets the date and time that the task was created.|
||[dueDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-duedatetime-member)|Gets or sets the date and time the task is due.|
||[id](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-id-member)|Gets the ID of the task.|
||[percentComplete](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-percentcomplete-member)|Specifies the completion percentage of the task.|
||[priority](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-priority-member)|Specifies the priority of the task.|
||[startDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-startdatetime-member)|Gets or sets the date and time the task starts.|
||[title](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-title-member)|Specifies title of the task.|
||[unassign(assignee: Excel.Identity)](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-unassign-member(1))|Removes the given user from the list of assignees attached to the task.|
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
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getcount-member(1))|Gets the number of tasks in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitem-member(1))|Gets a task using its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemat-member(1))|Gets a task by its index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemornullobject-member(1))|Gets a task using its ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-items-member)|Gets the loaded child items in this collection.|
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
