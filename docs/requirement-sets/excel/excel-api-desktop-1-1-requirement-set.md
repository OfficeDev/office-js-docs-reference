---
title: Excel JavaScript API desktop-only requirement set 1.1
description: Details about the ExcelApiDesktop 1.1 requirement set.
ms.date: 10/27/2025
ms.topic: whats-new
ms.localizationpriority: medium
---

# Excel JavaScript API desktop-only requirement set 1.1

The `ExcelApiDesktop` requirement set is a special requirement set that includes features that are only available in Excel for Windows and Excel for Mac. APIs in this requirement set are considered to be production APIs for the Excel application on Windows and Mac. They follow [Microsoft 365 developer support policies](/office/dev/add-ins/publish/maintain-breaking-changes). `ExcelApiDesktop` APIs are considered to be "preview" APIs for other platforms (such as web and iPad) and may not be supported by any of those platforms.

When APIs in the `ExcelApiDesktop` requirement set are supported across all platforms, they'll be added to the next released requirement set (`ExcelApi 1.[NEXT]`). Once that new requirement set is public, those APIs will also continue to be tagged in this `ExcelApiDesktop` requirement set. To learn more about platform-specific requirements in general, see [Understanding platform-specific requirement sets](https://aka.ms/PlatformSpecificReqtSets).

> [!IMPORTANT]
> `ExcelApiDesktop 1.1` is a desktop-only requirement set. It's a superset of the ExcelApi 1.20.

## Recommended usage

Because the `ExcelApiDesktop 1.1` APIs are only supported by Excel on Windows and Mac, your add-in should check if the requirement set is supported before calling these APIs. This avoids any attempt to use desktop-only APIs on an unsupported platform.

```js
if (Office.context.requirements.isSetSupported("ExcelApiDesktop", "1.1")) {
   // Any API exclusive to this ExcelApiDesktop requirement set.
}
```

Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check. This will enable your add-in's feature on other platforms. Be sure to test the feature on those platforms when making this change.

> [!IMPORTANT]
> Your manifest cannot specify `ExcelApiDesktop 1.1` as an activation requirement. It isn't a valid value to use in the [Set element](/javascript/api/manifest/set).

## API list

The following table lists the Excel JavaScript APIs currently included in the `ExcelApiDesktop 1.1` requirement set. For a complete list of all Excel JavaScript APIs (including `ExcelApiDesktop 1.1` APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-desktop-1.1&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-desktop-1_1.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-desktop-1.1&preserve-view=true)
- [Excel JavaScript preview APIs](excel-preview-apis.md)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
- [Understanding platform-specific requirement sets](https://aka.ms/PlatformSpecificReqtSets)
