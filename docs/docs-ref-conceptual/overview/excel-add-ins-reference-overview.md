# Excel JavaScript API overview

You can use the Excel JavaScript API to build add-ins for Excel 2016. The following list shows the high-level Excel objects that are available in the API. Each object page link contains a description of the properties, relationships, and methods that are available on the object. Explore the links from the menu to learn more.

Note that the relationships section within the document lists the properties that are used to navigate from the main object to another related object. These are non-scalar objects that themselves may contain other properties, methods and relationships.

[Workbook](../../api/excel/excel.workbook)

[WorksheetCollection](../../api/excel/excel.worksheetcollection)

Some of the core Excel objects are listed below for convenience: 

* [Workbook](../../api/excel/excel.workbook): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.

* [Worksheet](../../api/excel/excel.worksheet): Represents a worksheet in a workbook. 
  * [WorksheetCollection](../../api/excel/excel.worksheetcollection): A collection of the **Worksheet** objects in a workbook.

* [Range](../../api/excel/excel.range): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.

* [Table](../../api/excel/excel.table): Represents a collection of organized cells designed to make management of the data easy.
  * [TableCollection](../../api/excel/excel.tablecollection): A collection of tables in a workbook or worksheet.
  * [TableColumnCollection](../../api/excel/excel.tablecolumncollection): A collection of all the columns in a table.
  * [TableRowCollection](../../api/excel/excel.tablerowcollection): A collection of all the rows in a table.

* [Chart](../../api/excel/excel.chart): Represents a chart object in a worksheet, which is a visual representation of underlying data.
  * [ChartCollection](../../api/excel/excel.chartcollection): A collection of charts in a worksheet.

* [TableSort](../../api/excel/excel.tablesort): Represents an object that manages sorting operations on **Table** objects.

* [RangeSort](../../api/excel/excel.rangesort): Represents a object that manages sorting operations on **Range** objects.

* [Filter](../../api/excel/excel.filter): Represents an object that manages the filtering of a table's column.

* [WorksheetProtection](../../api/excel/excel.worksheetprotection): Represents the protection of a **Worksheet** object.

* [NamedItem](../../api/excel/excel.nameditem): Represents a defined name for a range of cells or a value. 
  * [NamedItemCollection](../../api/excel/excel.nameditemcollection): A collection of the **NamedItem** objects in a workbook.

* [Binding](../../api/excel/excel.binding): An abstract class that represents a binding to a section of the workbook.
  * [BindingCollection](../../api/excel/excel.bindingcollection): A collection of the **Binding** objects in a workbook.

## Excel JavaScript API open specifications

As we design and develop new APIs for Excel add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline for the Excel JavaScript APIs, and provide your input on our design specifications.

## Additional resources

* [Excel add-ins overview](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-overview)
* [Office Add-ins platform overview](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)
* [Word add-in samples on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Excel)
