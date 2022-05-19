| Class | Fields | Description |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-delete-member(1))|Deletes the custom XML part.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-getxml-member(1))|Gets the custom XML part's full XML content.|
||[id](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-id-member)|The custom XML part's ID.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-namespaceuri-member)|The custom XML part's namespace URI.|
||[setXml(xml: string)](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-setxml-member(1))|Sets the custom XML part's full XML content.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-add-member(1))|Adds a new custom XML part to the workbook.|
||[getByNamespace(namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getbynamespace-member(1))|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getcount-member(1))|Gets the number of custom XML parts in the collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getitem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getitemornullobject-member(1))|Gets a custom XML part based on its ID.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getcount-member(1))|Gets the number of CustomXML parts in this collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getitem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getitemornullobject-member(1))|Gets a custom XML part based on its ID.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getonlyitem-member(1))|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|If the collection contains exactly one item, this method returns it.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-items-member)|Gets the loaded child items in this collection.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#excel-excel-requestcontext-runtime-member)|[Api set: ExcelApi 1.5]|
|[RunOptions](/javascript/api/excel/excel.runoptions)|[delayForCellEdit](/javascript/api/excel/excel.runoptions#excel-excel-runoptions-delayforcelledit-member)|Determines whether Excel will delay the batch request until the user exits cell edit mode.|
