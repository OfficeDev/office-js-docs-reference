---
title: Excel JavaScript API requirement set 1.12
description: 'Details about the ExcelApi 1.12 requirement set.'
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.12

The ExcelApi 1.12 increased support for formulas in ranges by adding APIs for tracking dynamic arrays and finding a formula's direct precedents. It also added API control of PivotTable filters. Improvements were also made in the comment, culture settings, and custom properties feature areas.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Comment events](../../excel/excel-add-ins-comments.md#comment-events) | Adds events for add, change, and delete to the comment collection.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Date and time [culture settings](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Gives access to additional cultural settings around date and time formatting. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Direct precedents](../../excel/excel-add-ins-ranges-precedents.md) | Returns ranges that are used to evaluate a cell's formula.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Pivot Filters | Applies value-driven filters to the fields of a PivotTable. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotfilters) |
| [Range spilling](../../excel/excel-add-ins-ranges-dynamic-arrays.md) | Lets add-ins find ranges associated with [dynamic array](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) results. | [Range](/javascript/api/excel/excel.range) |
| [Worksheet-level custom properties](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Lets custom properties be scoped to the worksheet-level, in addition to being scoped to the workbook-level. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.12. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.12 or earlier, see [Excel APIs in requirement set 1.12 or earlier](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

[!INCLUDE[API table](../includes/excel-1-12.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
