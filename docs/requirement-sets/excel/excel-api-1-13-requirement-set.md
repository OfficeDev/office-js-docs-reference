---
title: Excel JavaScript API requirement set 1.13
description: Details about the ExcelApi 1.13 requirement set.
ms.date: 07/09/2021
ms.topic: whats-new
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.13

The ExcelApi 1.13 added a method to insert worksheets into a workbook from a Base64-encoded string and an event to detect workbook activation. It also increased support for formulas in ranges by adding APIs to track changes to formulas and locate a formula's direct dependent cells. Additionally, it expanded PivotTable support by adding PivotLayout APIs for alt text, style, and empty cell management.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Formula changed events](/office/dev/add-ins/excel/excel-add-ins-worksheets#detect-formula-changes) | Track changes to formulas, including the source and type of event that caused a change. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|
| [Formula dependents](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents-dependents#get-the-direct-dependents-of-a-formula) | Locate the direct dependent cells of a formula. | [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)) |
| [Insert worksheets](/office/dev/add-ins/excel/excel-add-ins-workbooks#insert-a-copy-of-an-existing-workbook-into-the-current-one) | Insert worksheets from another workbook into the current workbook as a Base64-encoded string. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1)) |
| [PivotTable PivotLayout](/office/dev/add-ins/excel/excel-add-ins-pivottables#other-pivotlayout-functions) | An expansion of the PivotLayout class, including new support for alt text and empty cell management. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.13. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.13 or earlier, see [Excel APIs in requirement set 1.13 or earlier](/javascript/api/excel?view=excel-js-1.13&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_13.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
