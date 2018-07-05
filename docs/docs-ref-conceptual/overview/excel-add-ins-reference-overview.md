# Excel JavaScript API overview

You can use the Excel JavaScript API to build add-ins for Excel 2016. The following list shows the high-level Excel objects that are available in the API. Each object page link contains a description of the properties, relationships, and methods that are available on the object. Explore the links from the menu to learn more.

Note that the relationships section within the document lists the properties that are used to navigate from the main object to another related object. These are non-scalar objects that themselves may contain other properties, methods and relationships.

Some of the core Excel objects are listed below for convenience: 

- [Workbook](../../docs-ref-autogen/excel/excel.workbook.yml): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.

- [Worksheet](../../docs-ref-autogen/excel/excel.worksheet.yml): Represents a worksheet in a workbook. 
    - [WorksheetCollection](../../docs-ref-autogen/excel/excel.worksheetcollection.yml): A collection of the **Worksheet** objects in a workbook.

- [Range](../../docs-ref-autogen/excel/excel.range.yml): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.

- [Table](../../docs-ref-autogen/excel/excel.table.yml): Represents a collection of organized cells designed to make management of the data easy.
    - [TableCollection](../../docs-ref-autogen/excel/excel.tablecollection.yml): A collection of tables in a workbook or worksheet.
    - [TableColumnCollection](../../docs-ref-autogen/excel/excel.tablecolumncollection.yml): A collection of all the columns in a table.
    - [TableRowCollection](../../docs-ref-autogen/excel/excel.tablerowcollection.yml): A collection of all the rows in a table.

- [Chart](../../docs-ref-autogen/excel/excel.chart.yml): Represents a chart object in a worksheet, which is a visual representation of underlying data.
    - [ChartCollection](../../docs-ref-autogen/excel/excel.chartcollection.yml): A collection of charts in a worksheet.

- [TableSort](../../docs-ref-autogen/excel/excel.tablesort.yml): Represents an object that manages sorting operations on **Table** objects.

- [RangeSort](../../docs-ref-autogen/excel/excel.rangesort.yml): Represents a object that manages sorting operations on **Range** objects.

- [Filter](../../docs-ref-autogen/excel/excel.filter.yml): Represents an object that manages the filtering of a table's column.

- [WorksheetProtection](../../docs-ref-autogen/excel/excel.worksheetprotection.yml): Represents the protection of a **Worksheet** object.

- [NamedItem](../../docs-ref-autogen/excel/excel.nameditem.yml): Represents a defined name for a range of cells or a value. 
    - [NamedItemCollection](../../docs-ref-autogen/excel/excel.nameditemcollection.yml): A collection of the **NamedItem** objects in a workbook.

- [Binding](../../docs-ref-autogen/excel/excel.binding.yml): An abstract class that represents a binding to a section of the workbook.
    - [BindingCollection](../../docs-ref-autogen/excel/excel.bindingcollection.yml): A collection of the **Binding** objects in a workbook.

## Excel JavaScript API open specifications

As we design and develop new APIs for Excel add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline for the Excel JavaScript APIs, and provide your input on our design specifications.

## Excel JavaScript API reference

For detailed information about Excel JavaScript API, see the [Excel JavaScript API reference documentation](../../docs-ref-autogen/excel.yml).

## See also

- [Excel add-ins overview](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Office Add-ins platform overview](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Word add-in samples on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Excel)
