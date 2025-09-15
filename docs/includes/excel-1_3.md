| Class | Fields | Description |
|:---|:---|:---|
|[Binding](/.binding)|[delete()](/.binding#excel-javascript/api/excel/-binding-delete-member(1))|Deletes the binding.|
|[BindingCollection](/.bindingcollection)|[add(range: Range \| string, bindingType: Excel.BindingType, id: string)](/.bindingcollection#excel-javascript/api/excel/-bindingcollection-add-member(1))|Add a new binding to a particular Range.|
||[addFromNamedItem(name: string, bindingType: Excel.BindingType, id: string)](/.bindingcollection#excel-javascript/api/excel/-bindingcollection-addfromnameditem-member(1))|Add a new binding based on a named item in the workbook.|
||[addFromSelection(bindingType: Excel.BindingType, id: string)](/.bindingcollection#excel-javascript/api/excel/-bindingcollection-addfromselection-member(1))|Add a new binding based on the current selection.|
|[PivotTable](/.pivottable)|[name](/.pivottable#excel-javascript/api/excel/-pivottable-name-member)|Name of the PivotTable.|
||[refresh()](/.pivottable#excel-javascript/api/excel/-pivottable-refresh-member(1))|Refreshes the PivotTable.|
||[worksheet](/.pivottable#excel-javascript/api/excel/-pivottable-worksheet-member)|The worksheet containing the current PivotTable.|
|[PivotTableCollection](/.pivottablecollection)|[getItem(name: string)](/.pivottablecollection#excel-javascript/api/excel/-pivottablecollection-getitem-member(1))|Gets a PivotTable by name.|
||[items](/.pivottablecollection#excel-javascript/api/excel/-pivottablecollection-items-member)|Gets the loaded child items in this collection.|
||[refreshAll()](/.pivottablecollection#excel-javascript/api/excel/-pivottablecollection-refreshall-member(1))|Refreshes all the pivot tables in the collection.|
|[Range](/.range)|[getVisibleView()](/.range#excel-javascript/api/excel/-range-getvisibleview-member(1))|Represents the visible rows of the current range.|
|[RangeView](/.rangeview)|[cellAddresses](/.rangeview#excel-javascript/api/excel/-rangeview-celladdresses-member)|Represents the cell addresses of the `RangeView`.|
||[columnCount](/.rangeview#excel-javascript/api/excel/-rangeview-columncount-member)|The number of visible columns.|
||[formulas](/.rangeview#excel-javascript/api/excel/-rangeview-formulas-member)|Represents the formula in A1-style notation.|
||[formulasLocal](/.rangeview#excel-javascript/api/excel/-rangeview-formulaslocal-member)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.|
||[formulasR1C1](/.rangeview#excel-javascript/api/excel/-rangeview-formulasr1c1-member)|Represents the formula in R1C1-style notation.|
||[getRange()](/.rangeview#excel-javascript/api/excel/-rangeview-getrange-member(1))|Gets the parent range associated with the current `RangeView`.|
||[index](/.rangeview#excel-javascript/api/excel/-rangeview-index-member)|Returns a value that represents the index of the `RangeView`.|
||[numberFormat](/.rangeview#excel-javascript/api/excel/-rangeview-numberformat-member)|Represents Excel's number format code for the given cell.|
||[rowCount](/.rangeview#excel-javascript/api/excel/-rangeview-rowcount-member)|The number of visible rows.|
||[rows](/.rangeview#excel-javascript/api/excel/-rangeview-rows-member)|Represents a collection of range views associated with the range.|
||[text](/.rangeview#excel-javascript/api/excel/-rangeview-text-member)|Text values of the specified range.|
||[valueTypes](/.rangeview#excel-javascript/api/excel/-rangeview-valuetypes-member)|Represents the type of data of each cell.|
||[values](/.rangeview#excel-javascript/api/excel/-rangeview-values-member)|Represents the raw values of the specified range view.|
|[RangeViewCollection](/.rangeviewcollection)|[getItemAt(index: number)](/.rangeviewcollection#excel-javascript/api/excel/-rangeviewcollection-getitemat-member(1))|Gets a `RangeView` row via its index.|
||[items](/.rangeviewcollection#excel-javascript/api/excel/-rangeviewcollection-items-member)|Gets the loaded child items in this collection.|
|[Table](/.table)|[highlightFirstColumn](/.table#excel-javascript/api/excel/-table-highlightfirstcolumn-member)|Specifies if the first column contains special formatting.|
||[highlightLastColumn](/.table#excel-javascript/api/excel/-table-highlightlastcolumn-member)|Specifies if the last column contains special formatting.|
||[showBandedColumns](/.table#excel-javascript/api/excel/-table-showbandedcolumns-member)|Specifies if the columns show banded formatting in which odd columns are highlighted differently from even ones, to make reading the table easier.|
||[showBandedRows](/.table#excel-javascript/api/excel/-table-showbandedrows-member)|Specifies if the rows show banded formatting in which odd rows are highlighted differently from even ones, to make reading the table easier.|
||[showFilterButton](/.table#excel-javascript/api/excel/-table-showfilterbutton-member)|Specifies if the filter buttons are visible at the top of each column header.|
|[Workbook](/.workbook)|[pivotTables](/.workbook#excel-javascript/api/excel/-workbook-pivottables-member)|Represents a collection of PivotTables associated with the workbook.|
|[Worksheet](/.worksheet)|[pivotTables](/.worksheet#excel-javascript/api/excel/-worksheet-pivottables-member)|Collection of PivotTables that are part of the worksheet.|
