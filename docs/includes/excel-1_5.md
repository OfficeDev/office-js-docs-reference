| Class | Fields | Description |
|:---|:---|:---|
|[CustomXmlPart](/.customxmlpart)|[delete()](/.customxmlpart#excel-javascript/api/excel/-customxmlpart-delete-member(1))|Deletes the custom XML part.|
||[getXml()](/.customxmlpart#excel-javascript/api/excel/-customxmlpart-getxml-member(1))|Gets the custom XML part's full XML content.|
||[id](/.customxmlpart#excel-javascript/api/excel/-customxmlpart-id-member)|The custom XML part's ID.|
||[namespaceUri](/.customxmlpart#excel-javascript/api/excel/-customxmlpart-namespaceuri-member)|The custom XML part's namespace URI.|
||[setXml(xml: string)](/.customxmlpart#excel-javascript/api/excel/-customxmlpart-setxml-member(1))|Sets the custom XML part's full XML content.|
|[CustomXmlPartCollection](/.customxmlpartcollection)|[add(xml: string)](/.customxmlpartcollection#excel-javascript/api/excel/-customxmlpartcollection-add-member(1))|Adds a new custom XML part to the workbook.|
||[getByNamespace(namespaceUri: string)](/.customxmlpartcollection#excel-javascript/api/excel/-customxmlpartcollection-getbynamespace-member(1))|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/.customxmlpartcollection#excel-javascript/api/excel/-customxmlpartcollection-getcount-member(1))|Gets the number of custom XML parts in the collection.|
||[getItem(id: string)](/.customxmlpartcollection#excel-javascript/api/excel/-customxmlpartcollection-getitem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/.customxmlpartcollection#excel-javascript/api/excel/-customxmlpartcollection-getitemornullobject-member(1))|Gets a custom XML part based on its ID.|
||[items](/.customxmlpartcollection#excel-javascript/api/excel/-customxmlpartcollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/.customxmlpartscopedcollection)|[getCount()](/.customxmlpartscopedcollection#excel-javascript/api/excel/-customxmlpartscopedcollection-getcount-member(1))|Gets the number of CustomXML parts in this collection.|
||[getItem(id: string)](/.customxmlpartscopedcollection#excel-javascript/api/excel/-customxmlpartscopedcollection-getitem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/.customxmlpartscopedcollection#excel-javascript/api/excel/-customxmlpartscopedcollection-getitemornullobject-member(1))|Gets a custom XML part based on its ID.|
||[getOnlyItem()](/.customxmlpartscopedcollection#excel-javascript/api/excel/-customxmlpartscopedcollection-getonlyitem-member(1))|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/.customxmlpartscopedcollection#excel-javascript/api/excel/-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|If the collection contains exactly one item, this method returns it.|
||[items](/.customxmlpartscopedcollection#excel-javascript/api/excel/-customxmlpartscopedcollection-items-member)|Gets the loaded child items in this collection.|
|[PivotTable](/.pivottable)|[id](/.pivottable#excel-javascript/api/excel/-pivottable-id-member)|ID of the PivotTable.|
|[RequestContext](/.requestcontext)|[runtime](/.requestcontext#excel-javascript/api/excel/-requestcontext-runtime-member)|[Api set: ExcelApi 1.5]|
|[Runtime](/.runtime)|||
|[Workbook](/.workbook)|[customXmlParts](/.workbook#excel-javascript/api/excel/-workbook-customxmlparts-member)|Represents the collection of custom XML parts contained by this workbook.|
|[Worksheet](/.worksheet)|[getNext(visibleOnly?: boolean)](/.worksheet#excel-javascript/api/excel/-worksheet-getnext-member(1))|Gets the worksheet that follows this one.|
||[getNextOrNullObject(visibleOnly?: boolean)](/.worksheet#excel-javascript/api/excel/-worksheet-getnextornullobject-member(1))|Gets the worksheet that follows this one.|
||[getPrevious(visibleOnly?: boolean)](/.worksheet#excel-javascript/api/excel/-worksheet-getprevious-member(1))|Gets the worksheet that precedes this one.|
||[getPreviousOrNullObject(visibleOnly?: boolean)](/.worksheet#excel-javascript/api/excel/-worksheet-getpreviousornullobject-member(1))|Gets the worksheet that precedes this one.|
|[WorksheetCollection](/.worksheetcollection)|[getFirst(visibleOnly?: boolean)](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-getfirst-member(1))|Gets the first worksheet in the collection.|
||[getLast(visibleOnly?: boolean)](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-getlast-member(1))|Gets the last worksheet in the collection.|
