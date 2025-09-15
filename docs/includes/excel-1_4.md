| Class | Fields | Description |
|:---|:---|:---|
|[BindingCollection](/.bindingcollection)|[getCount()](/.bindingcollection#excel-javascript/api/excel/-bindingcollection-getcount-member(1))|Gets the number of bindings in the collection.|
||[getItemOrNullObject(id: string)](/.bindingcollection#excel-javascript/api/excel/-bindingcollection-getitemornullobject-member(1))|Gets a binding object by ID.|
|[ChartCollection](/.chartcollection)|[getCount()](/.chartcollection#excel-javascript/api/excel/-chartcollection-getcount-member(1))|Returns the number of charts in the worksheet.|
||[getItemOrNullObject(name: string)](/.chartcollection#excel-javascript/api/excel/-chartcollection-getitemornullobject-member(1))|Gets a chart using its name.|
|[ChartPointsCollection](/.chartpointscollection)|[getCount()](/.chartpointscollection#excel-javascript/api/excel/-chartpointscollection-getcount-member(1))|Returns the number of chart points in the series.|
|[ChartSeriesCollection](/.chartseriescollection)|[getCount()](/.chartseriescollection#excel-javascript/api/excel/-chartseriescollection-getcount-member(1))|Returns the number of series in the collection.|
|[NamedItem](/.nameditem)|[comment](/.nameditem#excel-javascript/api/excel/-nameditem-comment-member)|Specifies the comment associated with this name.|
||[delete()](/.nameditem#excel-javascript/api/excel/-nameditem-delete-member(1))|Deletes the given name.|
||[getRangeOrNullObject()](/.nameditem#excel-javascript/api/excel/-nameditem-getrangeornullobject-member(1))|Returns the range object that is associated with the name.|
||[scope](/.nameditem#excel-javascript/api/excel/-nameditem-scope-member)|Specifies if the name is scoped to the workbook or to a specific worksheet.|
||[worksheet](/.nameditem#excel-javascript/api/excel/-nameditem-worksheet-member)|Returns the worksheet on which the named item is scoped to.|
||[worksheetOrNullObject](/.nameditem#excel-javascript/api/excel/-nameditem-worksheetornullobject-member)|Returns the worksheet to which the named item is scoped.|
|[NamedItemCollection](/.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/.nameditemcollection#excel-javascript/api/excel/-nameditemcollection-add-member(1))|Adds a new name to the collection of the given scope.|
||[addFormulaLocal(name: string, formula: string, comment?: string)](/.nameditemcollection#excel-javascript/api/excel/-nameditemcollection-addformulalocal-member(1))|Adds a new name to the collection of the given scope using the user's locale for the formula.|
||[getCount()](/.nameditemcollection#excel-javascript/api/excel/-nameditemcollection-getcount-member(1))|Gets the number of named items in the collection.|
||[getItemOrNullObject(name: string)](/.nameditemcollection#excel-javascript/api/excel/-nameditemcollection-getitemornullobject-member(1))|Gets a `NamedItem` object using its name.|
|[PivotTableCollection](/.pivottablecollection)|[getCount()](/.pivottablecollection#excel-javascript/api/excel/-pivottablecollection-getcount-member(1))|Gets the number of pivot tables in the collection.|
||[getItemOrNullObject(name: string)](/.pivottablecollection#excel-javascript/api/excel/-pivottablecollection-getitemornullobject-member(1))|Gets a PivotTable by name.|
|[Range](/.range)|[getIntersectionOrNullObject(anotherRange: Range \| string)](/.range#excel-javascript/api/excel/-range-getintersectionornullobject-member(1))|Gets the range object that represents the rectangular intersection of the given ranges.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/.range#excel-javascript/api/excel/-range-getusedrangeornullobject-member(1))|Returns the used range of the given range object.|
|[RangeViewCollection](/.rangeviewcollection)|[getCount()](/.rangeviewcollection#excel-javascript/api/excel/-rangeviewcollection-getcount-member(1))|Gets the number of `RangeView` objects in the collection.|
|[Setting](/.setting)|[delete()](/.setting#excel-javascript/api/excel/-setting-delete-member(1))|Deletes the setting.|
||[key](/.setting#excel-javascript/api/excel/-setting-key-member)|The key that represents the ID of the setting.|
||[value](/.setting#excel-javascript/api/excel/-setting-value-member)|Represents the value stored for this setting.|
|[SettingCollection](/.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date \| any[] \| any)](/.settingcollection#excel-javascript/api/excel/-settingcollection-add-member(1))|Sets or adds the specified setting to the workbook.|
||[getCount()](/.settingcollection#excel-javascript/api/excel/-settingcollection-getcount-member(1))|Gets the number of settings in the collection.|
||[getItem(key: string)](/.settingcollection#excel-javascript/api/excel/-settingcollection-getitem-member(1))|Gets a setting entry via the key.|
||[getItemOrNullObject(key: string)](/.settingcollection#excel-javascript/api/excel/-settingcollection-getitemornullobject-member(1))|Gets a setting entry via the key.|
||[items](/.settingcollection#excel-javascript/api/excel/-settingcollection-items-member)|Gets the loaded child items in this collection.|
||[onSettingsChanged](/.settingcollection#excel-javascript/api/excel/-settingcollection-onsettingschanged-member)|Occurs when the settings in the document are changed.|
|[SettingsChangedEventArgs](/.settingschangedeventargs)|[settings](/.settingschangedeventargs#excel-javascript/api/excel/-settingschangedeventargs-settings-member)|Gets the `Setting` object that represents the binding that raised the settings changed event|
|[TableCollection](/.tablecollection)|[getCount()](/.tablecollection#excel-javascript/api/excel/-tablecollection-getcount-member(1))|Gets the number of tables in the collection.|
||[getItemOrNullObject(key: string)](/.tablecollection#excel-javascript/api/excel/-tablecollection-getitemornullobject-member(1))|Gets a table by name or ID.|
|[TableColumnCollection](/.tablecolumncollection)|[getCount()](/.tablecolumncollection#excel-javascript/api/excel/-tablecolumncollection-getcount-member(1))|Gets the number of columns in the table.|
||[getItemOrNullObject(key: number \| string)](/.tablecolumncollection#excel-javascript/api/excel/-tablecolumncollection-getitemornullobject-member(1))|Gets a column object by name or ID.|
|[TableRowCollection](/.tablerowcollection)|[getCount()](/.tablerowcollection#excel-javascript/api/excel/-tablerowcollection-getcount-member(1))|Gets the number of rows in the table.|
|[Workbook](/.workbook)|[settings](/.workbook#excel-javascript/api/excel/-workbook-settings-member)|Represents a collection of settings associated with the workbook.|
|[Worksheet](/.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/.worksheet#excel-javascript/api/excel/-worksheet-getusedrangeornullobject-member(1))|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them.|
||[names](/.worksheet#excel-javascript/api/excel/-worksheet-names-member)|Collection of names scoped to the current worksheet.|
|[WorksheetCollection](/.worksheetcollection)|[getCount(visibleOnly?: boolean)](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-getcount-member(1))|Gets the number of worksheets in the collection.|
||[getItemOrNullObject(key: string)](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-getitemornullobject-member(1))|Gets a worksheet object using its name or ID.|
