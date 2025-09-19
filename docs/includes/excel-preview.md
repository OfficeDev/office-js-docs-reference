| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[activeWindow](/javascript/api/excel/excel.application#excel-excel-application-activewindow-member)|Returns a `window` object that represents the active window (the window on top).|
||[checkSpelling(word: string, options?: Excel.CheckSpellingOptions)](/javascript/api/excel/excel.application#excel-excel-application-checkspelling-member(1))|Checks the spelling of a single word.|
||[enterEditingMode()](/javascript/api/excel/excel.application#excel-excel-application-entereditingmode-member(1))|Enters editing mode for the selected range in the active worksheet.|
||[formatStaleValues](/javascript/api/excel/excel.application#excel-excel-application-formatstalevalues-member)|Specifies whether the Format Stale Values option within Calculation Options is enabled or disabled.|
||[union(firstRange: Range \| RangeAreas, secondRange: Range \| RangeAreas, ...additionalRanges: (Range \| RangeAreas)[])](/javascript/api/excel/excel.application#excel-excel-application-union-member(1))|Returns a `RangeAreas` object that represents the union of two or more `Range` or `RangeAreas` objects.|
||[windows](/javascript/api/excel/excel.application#excel-excel-application-windows-member)|Returns all the Excel windows.|
|[Base64EncodedImage](/javascript/api/excel/excel.base64encodedimage)|[data](/javascript/api/excel/excel.base64encodedimage#excel-excel-base64encodedimage-data-member)|The Base64-encoded string.|
||[type](/javascript/api/excel/excel.base64encodedimage#excel-excel-base64encodedimage-type-member)|The file type of the Base64-encoded image.|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errorsubtype-member)|Represents the type of `BlockedErrorCellValue`.|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[type](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-type-member)|Represents the type of this cell value.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errorsubtype-member)|Represents the type of `BusyErrorCellValue`.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errorsubtype-member)|Represents the type of `CalcErrorCellValue`.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataRange()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatarange-member(1))|Gets the data source of the whole chart.|
||[getDataRangeOrNullObject()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatarangeornullobject-member(1))|Gets the data source of the whole chart.|
|[CheckSpellingOptions](/javascript/api/excel/excel.checkspellingoptions)|[customDictionary](/javascript/api/excel/excel.checkspellingoptions#excel-excel-checkspellingoptions-customdictionary-member)|Optional.|
||[ignoreUppercase](/javascript/api/excel/excel.checkspellingoptions#excel-excel-checkspellingoptions-ignoreuppercase-member)|Optional.|
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
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-startandduedatetime-member)|Specifies the date and time the task should start and is due.|
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
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[type](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-type-member)|Represents the type of this cell value.|
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
|[ExternalCodeServiceObjectCellValue](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue)|[Python_str](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-python_str-member)|Represents the output of the `str()` function when used on this object.|
||[Python_type](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-python_type-member)|Represents the full type name of this object.|
||[Python_typeName](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-python_typename-member)|Represents the short type name of this object.|
||[basicType](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[language](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-language-member)|Represents the runtime language of this external code service.|
||[preview](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-preview-member)|Represents the preview value shown in the cell.|
||[provider](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-provider-member)|Represents information about the service that provided the data in this `ExternalCodeServiceObjectCellValue`.|
||[type](/javascript/api/excel/excel.externalcodeserviceobjectcellvalue#excel-excel-externalcodeserviceobjectcellvalue-type-member)|Represents the type of this cell value.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooterPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerfooterpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the center section of the footer.|
||[centerHeaderPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerheaderpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the center section of the header.|
||[leftFooterPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftfooterpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the left section of the footer.|
||[leftHeaderPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftheaderpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the left section of the header.|
||[rightFooterPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightfooterpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the right section of the footer.|
||[rightHeaderPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightheaderpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the right section of the header.|
|[HeaderFooterPicture](/javascript/api/excel/excel.headerfooterpicture)|[brightness](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-brightness-member)|Specifies the brightness of the picture.|
||[colorType](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-colortype-member)|Specifies the type of color transformation of the picture.|
||[contrast](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-contrast-member)|Specifies the contrast of the picture.|
||[cropBottom](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-cropbottom-member)|Specifies the number of points that are cropped off the bottom of the picture.|
||[cropLeft](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-cropleft-member)|Specifies the number of points that are cropped off the left side of the picture.|
||[cropRight](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-cropright-member)|Specifies the number of points that are cropped off the right side of the picture.|
||[cropTop](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-croptop-member)|Specifies the number of points that are cropped off the top of the picture.|
||[filename](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-filename-member)|Specifies the URL (on the intranet or the web) or path (local or network) to the location where the source object is saved.|
||[height](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-height-member)|Specifies the height of the picture in points.|
||[lockAspectRatio](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-lockaspectratio-member)|Specifies a value that indicates whether the picture retains its original proportions when resized.|
||[width](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-width-member)|Specifies the width of the picture in points.|
|[Identity](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#excel-excel-identity-displayname-member)|Represents the user's display name.|
||[id](/javascript/api/excel/excel.identity#excel-excel-identity-id-member)|Represents the user's unique ID.|
|[Image](/javascript/api/excel/excel.image)|[brightness](/javascript/api/excel/excel.image#excel-excel-image-brightness-member)|Specifies the brightness of the image.|
||[colorType](/javascript/api/excel/excel.image#excel-excel-image-colortype-member)|Specifies the type of color transformation applied to the image.|
||[contrast](/javascript/api/excel/excel.image#excel-excel-image-contrast-member)|Specifies the contrast of the image.|
||[cropBottom](/javascript/api/excel/excel.image#excel-excel-image-cropbottom-member)|Specifies the number of points that are cropped off the bottom of the image.|
||[cropLeft](/javascript/api/excel/excel.image#excel-excel-image-cropleft-member)|Specifies the number of points that are cropped off the left side of the image.|
||[cropRight](/javascript/api/excel/excel.image#excel-excel-image-cropright-member)|Specifies the number of points that are cropped off the right side of the image.|
||[cropTop](/javascript/api/excel/excel.image#excel-excel-image-croptop-member)|Specifies the number of points that are cropped off the top of the image.|
||[incrementBrightness(increment: number)](/javascript/api/excel/excel.image#excel-excel-image-incrementbrightness-member(1))|Increments the brightness of the image by a specified amount.|
||[incrementContrast(increment: number)](/javascript/api/excel/excel.image#excel-excel-image-incrementcontrast-member(1))|Increments the contrast of the image by a specified amount.|
|[LocalImage](/javascript/api/excel/excel.localimage)|[getBase64EncodedImageData(cacheUid: string)](/javascript/api/excel/excel.localimage#excel-excel-localimage-getbase64encodedimagedata-member(1))|Gets the Base64-encoded image data stored in the shared image cache with the cache unique identifier (UID).|
|[LocalImageCellValue](/javascript/api/excel/excel.localimagecellvalue)|[altText](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-alttext-member)|Represents the alternate text used in accessibility scenarios to describe what the image represents.|
||[attribution](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-attribution-member)|Represents attribution information to describe the source and license requirements for this image.|
||[basicType](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[image](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-image-member)|Represents the image itself, either cached or encoded.|
||[provider](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-provider-member)|Represents information that describes the entity or individual who provided the image.|
||[type](/javascript/api/excel/excel.localimagecellvalue#excel-excel-localimagecellvalue-type-member)|Represents the type of this cell value.|
|[LocalImageCellValueCacheId](/javascript/api/excel/excel.localimagecellvaluecacheid)|[cachedUid](/javascript/api/excel/excel.localimagecellvaluecacheid#excel-excel-localimagecellvaluecacheid-cacheduid-member)|Represents the image's UID as it appears in the cache.|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-errorsubtype-member)|Represents the type of `NameErrorCellValue`.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemornullobject-member(1))|Gets a sheet view using its name.|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-errorsubtype-member)|Represents the type of `NotAvailableErrorCellValue`.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[alignMarginsHeaderFooter](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-alignmarginsheaderfooter-member)|Specifies whether Excel aligns the header and the footer with the margins set in the page setup options.|
||[printQuality](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printquality-member)|Specifies a two-element array that contains both horizontal and vertical print quality values.|
|[Pane](/javascript/api/excel/excel.pane)|[index](/javascript/api/excel/excel.pane#excel-excel-pane-index-member)|Returns index of the pane.|
|[PaneCollection](/javascript/api/excel/excel.panecollection)|[getCount()](/javascript/api/excel/excel.panecollection#excel-excel-panecollection-getcount-member(1))|Returns the number of bindings in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.panecollection#excel-excel-panecollection-getitemat-member(1))|Gets the Pane in the collection by index.|
||[items](/javascript/api/excel/excel.panecollection#excel-excel-panecollection-items-member)|Gets the loaded child items in this collection.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcell-member(1))|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-pivotstyle-member)|The style applied to the PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setstyle-member(1))|Sets the style applied to the PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[autoRefresh](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-autorefresh-member)|Specifies whether the PivotTable auto refreshes when the source data changes.|
|[PythonErrorCellValue](/javascript/api/excel/excel.pythonerrorcellvalue)|[basicType](/javascript/api/excel/excel.pythonerrorcellvalue#excel-excel-pythonerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.pythonerrorcellvalue#excel-excel-pythonerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.pythonerrorcellvalue#excel-excel-pythonerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.pythonerrorcellvalue#excel-excel-pythonerrorcellvalue-type-member)|Represents the type of this cell value.|
|[Query](/javascript/api/excel/excel.query)|[delete()](/javascript/api/excel/excel.query#excel-excel-query-delete-member(1))|Deletes the query and associated connection.|
||[refresh()](/javascript/api/excel/excel.query#excel-excel-query-refresh-member(1))|Refreshes the query.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[refreshAll()](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-refreshall-member(1))|Refresh all queries.|
|[Range](/javascript/api/excel/excel.range)|[checkSpelling(options?: Excel.CheckSpellingOptions)](/javascript/api/excel/excel.range#excel-excel-range-checkspelling-member(1))|Checks the spelling of words in this range.|
||[formulaArray](/javascript/api/excel/excel.range#excel-excel-range-formulaarray-member)|Specifies the array formula of a range.|
||[showDependents(remove?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-showdependents-member(1))|Draws tracer arrows to the direct dependents of the range.|
||[showPrecedents(remove?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-showprecedents-member(1))|Draws tracer arrows to the direct precedents of the range.|
||[togglePythonMarshalMode(marshalMode?: Excel.PythonMarshalMode)](/javascript/api/excel/excel.range#excel-excel-range-togglepythonmarshalmode-member(1))|Sets the marshaling mode of the Python in Excel formula =PY.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|||
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[errorSubType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errorsubtype-member)|Represents the type of `RefErrorCellValue`.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[setInvocation(invocation: { invocationId: number isInCFSyncScenario: boolean })](/javascript/api/excel/excel.requestcontext#excel-excel-requestcontext-setinvocation-member(1))||
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addLocalImageReference(address: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addlocalimagereference-member(1))|Creates a reference for the local image stored in the cell address and displays it as a floating shape over cells.|
||[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1))|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[strikethrough](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-strikethrough-member)|Specifies the strikethrough status of font.|
||[subscript](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-subscript-member)|Specifies the subscript status of font.|
||[superscript](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-superscript-member)|Specifies the superscript status of font.|
||[tintAndShade](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-tintandshade-member)|Specifies a double that lightens or darkens a color for the range font.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#excel-excel-slicer-nameinformula-member)|Specifies the slicer name used in the formula.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#excel-excel-slicer-setstyle-member(1))|Sets the style applied to the slicer.|
||[slicerStyle](/javascript/api/excel/excel.slicer#excel-excel-slicer-slicerstyle-member)|The style applied to the slicer.|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[type](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-type-member)|Represents the type of this cell value.|
|[Table](/javascript/api/excel/excel.table)|[altTextDescription](/javascript/api/excel/excel.table#excel-excel-table-alttextdescription-member)|Specifies the alternative text for accessibility.|
||[altTextTitle](/javascript/api/excel/excel.table#excel-excel-table-alttexttitle-member)|Specifies a summary for the table, such as one used by screen readers.|
||[clearStyle()](/javascript/api/excel/excel.table#excel-excel-table-clearstyle-member(1))|Changes the table to use the default table style.|
||[comment](/javascript/api/excel/excel.table#excel-excel-table-comment-member)|Specifies a comment associated with the table.|
||[isActive](/javascript/api/excel/excel.table#excel-excel-table-isactive-member)|Retrieves whether the table is currently active.|
||[onFiltered](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member)|Occurs when a filter is applied on a specific table.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#excel-excel-table-setstyle-member(1))|Sets the style applied to the table.|
||[source](/javascript/api/excel/excel.table#excel-excel-table-source-member)|Retrieves the data source type from which the table originates.|
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
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errorsubtype-member)|Represents the type of `ValueErrorCellValue`.|
|[Window](/javascript/api/excel/excel.window)|[activate()](/javascript/api/excel/excel.window#excel-excel-window-activate-member(1))|Activates the window.|
||[activateNext()](/javascript/api/excel/excel.window#excel-excel-window-activatenext-member(1))|Activates the next window.|
||[activatePrevious()](/javascript/api/excel/excel.window#excel-excel-window-activateprevious-member(1))|Activates the previous window.|
||[activeCell](/javascript/api/excel/excel.window#excel-excel-window-activecell-member)|Specifies the active cell in the window.|
||[activePane](/javascript/api/excel/excel.window#excel-excel-window-activepane-member)|Specifies the active pane in the window.|
||[activeWorksheet](/javascript/api/excel/excel.window#excel-excel-window-activeworksheet-member)|Specifies the active sheet in the window.|
||[autoFilterDateGroupingEnabled](/javascript/api/excel/excel.window#excel-excel-window-autofilterdategroupingenabled-member)|Specifies whether AutoFilter date grouping is enabled in the window.|
||[close()](/javascript/api/excel/excel.window#excel-excel-window-close-member(1))|Closes the window.|
||[enableResize](/javascript/api/excel/excel.window#excel-excel-window-enableresize-member)|Specifies a value indicating whether resizing is enabled for the window.|
||[freezePanes](/javascript/api/excel/excel.window#excel-excel-window-freezepanes-member)|Specifies a value indicating whether panes are frozen in the window.|
||[height](/javascript/api/excel/excel.window#excel-excel-window-height-member)|Specifies the height of the window.|
||[index](/javascript/api/excel/excel.window#excel-excel-window-index-member)|Gets the index of the window.|
||[isVisible](/javascript/api/excel/excel.window#excel-excel-window-isvisible-member)|Specifies the visibility of the window.|
||[largeScroll(Down: number, Up: number, ToRight: number, ToLeft: number)](/javascript/api/excel/excel.window#excel-excel-window-largescroll-member(1))|Scrolls the window by a large amount.|
||[left](/javascript/api/excel/excel.window#excel-excel-window-left-member)|Specifies the left position of the window.|
||[name](/javascript/api/excel/excel.window#excel-excel-window-name-member)|Specifies the name of the window.|
||[newWindow()](/javascript/api/excel/excel.window#excel-excel-window-newwindow-member(1))|Open a new window|
||[panes](/javascript/api/excel/excel.window#excel-excel-window-panes-member)|Gets the panes associated with the window.|
||[pointsToScreenPixelsX(Points: number)](/javascript/api/excel/excel.window#excel-excel-window-pointstoscreenpixelsx-member(1))|Converts horizontal points to screen pixels.|
||[pointsToScreenPixelsY(Points: number)](/javascript/api/excel/excel.window#excel-excel-window-pointstoscreenpixelsy-member(1))|Converts vertical points to screen pixels.|
||[rangeSelection](/javascript/api/excel/excel.window#excel-excel-window-rangeselection-member)|Gets the range selection in the window.|
||[scrollColumn](/javascript/api/excel/excel.window#excel-excel-window-scrollcolumn-member)|Specifies the scroll column of the window.|
||[scrollIntoView(Left: number, Top: number, Width: number, Height: number, Start?: boolean)](/javascript/api/excel/excel.window#excel-excel-window-scrollintoview-member(1))|Scrolls the window to bring the specified range into view.|
||[scrollRow](/javascript/api/excel/excel.window#excel-excel-window-scrollrow-member)|Specifies the scroll row of the window.|
||[scrollWorkbookTabs(Sheets?: number, Position?: Excel.ScrollWorkbookTabPosition)](/javascript/api/excel/excel.window#excel-excel-window-scrollworkbooktabs-member(1))|Scrolls the workbook tabs.|
||[showFormulas](/javascript/api/excel/excel.window#excel-excel-window-showformulas-member)|Specifies the display of formulas in the window.|
||[showGridlines](/javascript/api/excel/excel.window#excel-excel-window-showgridlines-member)|Specifies the display of gridlines in the window.|
||[showHeadings](/javascript/api/excel/excel.window#excel-excel-window-showheadings-member)|Specifies the display of headings in the window.|
||[showHorizontalScrollBar](/javascript/api/excel/excel.window#excel-excel-window-showhorizontalscrollbar-member)|Specifies the display of the horizontal scroll bar in the window.|
||[showOutline](/javascript/api/excel/excel.window#excel-excel-window-showoutline-member)|Specifies the display of the outline in the window.|
||[showRightToLeft](/javascript/api/excel/excel.window#excel-excel-window-showrighttoleft-member)|Specifies the display of the right-to-left layout in the window.|
||[showRuler](/javascript/api/excel/excel.window#excel-excel-window-showruler-member)|Specifies the display of the ruler in the window.|
||[showVerticalScrollBar](/javascript/api/excel/excel.window#excel-excel-window-showverticalscrollbar-member)|Specifies the display of the vertical scroll bar in the window.|
||[showWhitespace](/javascript/api/excel/excel.window#excel-excel-window-showwhitespace-member)|Specifies the display of whitespace in the window.|
||[showWorkbookTabs](/javascript/api/excel/excel.window#excel-excel-window-showworkbooktabs-member)|Specifies the display of workbook tabs in the window.|
||[showZeros](/javascript/api/excel/excel.window#excel-excel-window-showzeros-member)|Specifies the display of zeros in the window.|
||[smallScroll(Down: number, Up: number, ToRight: number, ToLeft: number)](/javascript/api/excel/excel.window#excel-excel-window-smallscroll-member(1))|Scrolls the window by a small amount.|
||[split](/javascript/api/excel/excel.window#excel-excel-window-split-member)|Specifies the split state of the window.|
||[splitColumn](/javascript/api/excel/excel.window#excel-excel-window-splitcolumn-member)|Specifies the split column of the window.|
||[splitHorizontal](/javascript/api/excel/excel.window#excel-excel-window-splithorizontal-member)|Specifies the horizontal split of the window.|
||[splitRow](/javascript/api/excel/excel.window#excel-excel-window-splitrow-member)|Specifies the split row of the window.|
||[splitVertical](/javascript/api/excel/excel.window#excel-excel-window-splitvertical-member)|Specifies the vertical split of the window.|
||[tabRatio](/javascript/api/excel/excel.window#excel-excel-window-tabratio-member)|Specifies the tab ratio of the window.|
||[top](/javascript/api/excel/excel.window#excel-excel-window-top-member)|Specifies the top position of the window.|
||[type](/javascript/api/excel/excel.window#excel-excel-window-type-member)|Specifies the type of the window.|
||[usableHeight](/javascript/api/excel/excel.window#excel-excel-window-usableheight-member)|Specifies the usable height of the window.|
||[usableWidth](/javascript/api/excel/excel.window#excel-excel-window-usablewidth-member)|Specifies the usable width of the window.|
||[view](/javascript/api/excel/excel.window#excel-excel-window-view-member)|Specifies the view of the window.|
||[visibleRange](/javascript/api/excel/excel.window#excel-excel-window-visiblerange-member)|Gets the visible range of the window.|
||[width](/javascript/api/excel/excel.window#excel-excel-window-width-member)|Returns or sets an integer value that represents the display size of the window.|
||[windowNumber](/javascript/api/excel/excel.window#excel-excel-window-windownumber-member)|Specifies the window number.|
||[windowState](/javascript/api/excel/excel.window#excel-excel-window-windowstate-member)|Returns or sets an integer value that represents the display size of the window.|
||[zoom](/javascript/api/excel/excel.window#excel-excel-window-zoom-member)|Specifies an integer value that represents the display size of the window.|
|[WindowCollection](/javascript/api/excel/excel.windowcollection)|[breakSideBySide()](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-breaksidebyside-member(1))|Breaks the side-by-side view of windows.|
||[compareCurrentSideBySideWith(windowName: string)](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-comparecurrentsidebysidewith-member(1))|Compares the current window side by side with the specified window.|
||[getCount()](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-getcount-member(1))|Gets the number of windows in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-getitemat-member(1))|Gets the Window in the collection by index.|
||[items](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-items-member)|Gets the loaded child items in this collection.|
||[resetPositionsSideBySide()](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-resetpositionssidebyside-member(1))|Resets the positions of windows in side-by-side view.|
|[Workbook](/javascript/api/excel/excel.workbook)|[enterPreviewMode()](/javascript/api/excel/excel.workbook#excel-excel-workbook-enterpreviewmode-member(1))|Enters Scratchpad Preview Mode for the workbook, showing changes suggested by Copilot to the user.|
||[exitPreviewMode(applyChanges: boolean)](/javascript/api/excel/excel.workbook#excel-excel-workbook-exitpreviewmode-member(1))|Exits Scratchpad Preview Mode for the workbook.|
||[externalCodeServiceTimeout](/javascript/api/excel/excel.workbook#excel-excel-workbook-externalcodeservicetimeout-member)|Specifies the maximum length of time, in seconds, allotted for a formula that depends on an external code service to complete.|
||[focus()](/javascript/api/excel/excel.workbook#excel-excel-workbook-focus-member(1))|Sets focus on the workbook.|
||[localImage](/javascript/api/excel/excel.workbook#excel-excel-workbook-localimage-member)|Returns the `LocalImage` object associated with the workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#excel-excel-workbook-showpivotfieldlist-member)|Specifies whether the PivotTable's field list pane is shown at the workbook level.|
||[tasks](/javascript/api/excel/excel.workbook#excel-excel-workbook-tasks-member)|Returns a collection of tasks that are present in the workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#excel-excel-workbook-use1904datesystem-member)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[checkSpelling(options?: Excel.CheckSpellingOptions)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-checkspelling-member(1))|Checks the spelling of words in this worksheet.|
||[clearArrows()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-cleararrows-member(1))|Clears the tracer arrows from the worksheet.|
||[evaluate(name: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-evaluate-member(1))|Returns the evaluation result of a formula string.|
||[onFiltered](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member)|Occurs when a filter is applied on a specific worksheet.|
||[tasks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tasks-member)|Returns a collection of tasks that are present in the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-addfrombase64-member(1))|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member)|Occurs when any worksheet's filter is applied in the workbook.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-worksheetid-member)|Gets the ID of the worksheet in which the filter is applied.|
