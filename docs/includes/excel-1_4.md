| Class | Fields | Description |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getcount-member(1))|Gets the number of bindings in the collection.|
||[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitem-member(1))|Gets a binding object by ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemat-member(1))|Gets a binding object based on its position in the items array.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemornullobject-member(1))|Gets a binding object by ID.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getcount-member(1))|Returns the number of charts in the worksheet.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitem-member(1))|Gets a chart using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemat-member(1))|Gets a chart based on its position in the collection.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemornullobject-member(1))|Gets a chart using its name.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|||
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getcount-member(1))|Returns the number of chart points in the series.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getitemat-member(1))|Retrieve a point based on its position within the series.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-comment-member)|Specifies the comment associated with this name.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-add-member(1))|Adds a new name to the collection of the given scope.|
||[addFormulaLocal(name: string, formula: string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-addformulalocal-member(1))|Adds a new name to the collection of the given scope using the user's locale for the formula.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getcount-member(1))|Gets the number of named items in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitem-member(1))|Gets a `NamedItem` object using its name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitemornullobject-member(1))|Gets a `NamedItem` object using its name.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getcount-member(1))|Gets the number of `RangeView` objects in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getitemat-member(1))|Gets a `RangeView` row via its index.|
|[Setting](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#excel-excel-setting-delete-member(1))|Deletes the setting.|
||[key](/javascript/api/excel/excel.setting#excel-excel-setting-key-member)|The key that represents the ID of the setting.|
||[value](/javascript/api/excel/excel.setting#excel-excel-setting-value-member)|Represents the value stored for this setting.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date \| Array \| any)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-add-member(1))|Sets or adds the specified setting to the workbook.|
||[getCount()](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getcount-member(1))|Gets the number of settings in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitem-member(1))|Gets a setting entry via the key.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitemornullobject-member(1))|Gets a setting entry via the key.|
||[items](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-items-member)|Gets the loaded child items in this collection.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member)|Occurs when the settings in the document are changed.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#excel-excel-settingschangedeventargs-settings-member)|Gets the `Setting` object that represents the binding that raised the settings changed event|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getcount-member(1))|Gets the number of tables in the collection.|
||[getCount()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getcount-member(1))|Gets the number of columns in the table.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitem-member(1))|Gets a column object by name or ID.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitem-member(1))|Gets a table by name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemat-member(1))|Gets a table based on its position in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemat-member(1))|Gets a column based on its position in the collection.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemornullobject-member(1))|Gets a column object by name or ID.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemornullobject-member(1))|Gets a table by name or ID.|
