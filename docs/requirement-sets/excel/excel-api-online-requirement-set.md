---
title: Excel JavaScript API online-only requirement set
description: Details about the ExcelApiOnline requirement set.
ms.date: 08/29/2024
ms.topic: whats-new
ms.localizationpriority: medium
---

# Excel JavaScript API online-only requirement set

The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web. APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web application. `ExcelApiOnline` APIs are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.

When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will be added to the next released requirement set (`ExcelApi 1.[NEXT]`). Once that new requirement set is public, those APIs will be removed from `ExcelApiOnline`. Think of this as a similar promotion process to an API moving from preview to release.

> [!IMPORTANT]
> `ExcelApiOnline` is a superset of the latest numbered requirement set.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` is the only version of the online-only APIs. This is because Excel on the web will always have a single version available to users that is the latest version.

The following table provides a concise summary of the APIs, while the subsequent [API list](#api-list) table gives a detailed list of the current `ExcelApiOnline` APIs.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Linked workbooks | Manage links between workbooks, including support for refreshing and breaking workbook links. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Named sheet views | Gives programmatic control of per-user worksheet views. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |
| Worksheet move events | Detect when worksheets are moved within a collection, the position of the worksheet, and the source of the change. | [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection), [WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs) |
| Worksheet protection | Prevent unauthorized users from making changes to specified ranges within a worksheet. | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [AllowEditRange](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## Recommended usage

Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs. This avoids calling an online-only API on a different platform.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check. This will enable your add-in's feature on other platforms. Be sure to test the feature on those platforms when making this change.

> [!IMPORTANT]
> Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement. It is not a valid value to use in the [Set element](/javascript/api/manifest/set).

## API list

The following table lists the Excel JavaScript APIs currently included in the `ExcelApiOnline` requirement set. For a complete list of all Excel JavaScript APIs (including `ExcelApiOnline` APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-online&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-online.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript preview APIs](excel-preview-apis.md)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
- [Understanding platform-specific requirement sets](https://aka.ms/PlatformSpecificReqtSets)
