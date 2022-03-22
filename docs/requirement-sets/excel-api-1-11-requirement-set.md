---
title: Excel JavaScript API requirement set 1.11
description: 'Details about the ExcelApi 1.11 requirement set.'
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.11

The ExcelApi 1.11 improved support for comments and workbook-level controls (such as saving and closing the workbook). It also added access to culture settings to help account for localization.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Comment [Mentions](../../excel/excel-add-ins-comments.md#mentions) |Tags and notifies other workbook users through comments. | [Comment](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| Comment [Resolution](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | Resolve comment threads and get the resolution status. | [Comment](/javascript/api/excel/excel.comment) |
| [Culture settings](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Gets cultural system settings for the workbook, such as number formatting. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Cut and paste (moveTo)](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Replicates the cut-and-paste functionality in Excel for a Range. | [Range](/javascript/api/excel/excel.range) |
| Workbook [Save](../../excel/excel-add-ins-workbooks.md#save-the-workbook) and [Close](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | Save and close workbooks. | [Workbook](/javascript/api/excel/excel.workbook) |
| Worksheet events | Additional events and event information for worksheet calculations and hidden rows. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.11. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.11 or earlier, see [Excel APIs in requirement set 1.11 or earlier](/javascript/api/excel?view=excel-js-1.11&preserve-view=true).

[!INCLUDE[API table](../includes/excel-1-11.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
