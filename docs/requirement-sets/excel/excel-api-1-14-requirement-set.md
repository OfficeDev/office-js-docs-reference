---
title: Excel JavaScript API requirement set 1.14
description: 'Details about the ExcelApi 1.14 requirement set.'
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.14

The ExcelApi 1.14 added objects to control the data table feature of a chart, a method to locate all the precedent cells of a formula, and worksheet protection events to track changes to the protection state of a worksheet. It also added multiple [`getItemOrNullObject`](/office/dev/add-ins/develop/application-specific-api-model.md#ornullobject-methods-and-properties) methods for objects like `CommentCollection`, `ShapeCollection`, and `StyleCollection` to improve error handling.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Chart data tables](/office/dev/add-ins/excel/excel-add-ins-charts.md#add-and-format-a-chart-data-table) | Control appearance, formatting, and visibility of data tables on charts. | [Chart](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| [Formula precedents](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents-dependents.md#get-the-precedents-of-a-formula) | Return all the precedent cells of a formula. | [Range](/javascript/api/excel/excel.range) |
| Queries | Retrieve Power Query attributes like name, refresh date, and query count. | [Query](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| [Worksheet protection events](/office/dev/add-ins/excel/excel-add-ins-worksheets.md#detect-changes-to-the-worksheet-protection-state) | Track changes to the protection state of a worksheet and the source of those changes. | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [Worksheet](/javascript/api/excel/excel.worksheet), [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.14. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.14 or earlier, see [Excel APIs in requirement set 1.14 or earlier](/javascript/api/excel?view=excel-js-1.14&preserve-view=true).

[!INCLUDE[API table](../includes/excel-1_14.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
