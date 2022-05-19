| Class | Fields | Description |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#excel-excel-binding-delete-member(1))|Deletes the binding.|
||[getRange()](/javascript/api/excel/excel.binding#excel-excel-binding-getrange-member(1))|Returns the range represented by the binding.|
||[getTable()](/javascript/api/excel/excel.binding#excel-excel-binding-gettable-member(1))|Returns the table represented by the binding.|
||[getText()](/javascript/api/excel/excel.binding#excel-excel-binding-gettext-member(1))|Returns the text represented by the binding.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-add-member(1))|Add a new binding to a particular Range.|
||[addFromNamedItem(name: string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-addfromnameditem-member(1))|Add a new binding based on a named item in the workbook.|
||[addFromSelection(bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-addfromselection-member(1))|Add a new binding based on the current selection.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[items](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-items-member)|Gets the loaded child items in this collection.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[items](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-items-member)|Gets the loaded child items in this collection.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-refreshall-member(1))|Refreshes all the pivot tables in the collection.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[items](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-items-member)|Gets the loaded child items in this collection.|
||[items](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-items-member)|Gets the loaded child items in this collection.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-add-member(1))|Creates a new table.|
||[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, name?: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-add-member(1))|Adds a new column to the table.|
||[clearFilters()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-clearfilters-member(1))|Clears all the filters currently applied on the table.|
||[convertToRange()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-converttorange-member(1))|Converts the table into a normal range of cells.|
||[count](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-count-member)|Returns the number of tables in the workbook.|
||[count](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-count-member)|Returns the number of columns in the table.|
||[delete()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-delete-member(1))|Deletes the table.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getdatabodyrange-member(1))|Gets the range object associated with the data body of the table.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getheaderrowrange-member(1))|Gets the range object associated with the header row of the table.|
||[getRange()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getrange-member(1))|Gets the range object associated with the entire table.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-gettotalrowrange-member(1))|Gets the range object associated with the totals row of the table.|
||[items](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-items-member)|Gets the loaded child items in this collection.|
||[items](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-items-member)|Gets the loaded child items in this collection.|
||[items](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-items-member)|Gets the loaded child items in this collection.|
||[reapplyFilters()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-reapplyfilters-member(1))|Reapplies all the filters currently on the table.|
||[showBandedColumns](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-showbandedcolumns-member)|Specifies if the columns show banded formatting in which odd columns are highlighted differently from even ones, to make reading the table easier.|
||[showBandedRows](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-showbandedrows-member)|Specifies if the rows show banded formatting in which odd rows are highlighted differently from even ones, to make reading the table easier.|
||[showFilterButton](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-showfilterbutton-member)|Specifies if the filter buttons are visible at the top of each column header.|
||[showHeaders](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-showheaders-member)|Specifies if the header row is visible.|
||[showTotals](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-showtotals-member)|Specifies if the total row is visible.|
||[style](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-style-member)|Constant value that represents the table style.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-delete-member(1))|Deletes the column from the table.|
||[filter](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-filter-member)|Retrieves the filter applied to the column.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getdatabodyrange-member(1))|Gets the range object associated with the data body of the column.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getheaderrowrange-member(1))|Gets the range object associated with the header row of the column.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getrange-member(1))|Gets the range object associated with the entire column.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-gettotalrowrange-member(1))|Gets the range object associated with the totals row of the column.|
||[id](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-id-member)|Returns a unique key that identifies the column within the table.|
||[index](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-index-member)|Returns the index number of the column within the columns collection of the table.|
||[name](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-name-member)|Specifies the name of the table column.|
||[values](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-values-member)|Represents the raw values of the specified range.|
