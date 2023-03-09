---
title: Excel JavaScript API requirement sets
description: Office Add-in requirement set information for Excel builds.
ms.date: 11/10/2022
ms.topic: overview
ms.prod: excel
ms.localizationpriority: high
---

# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## Requirement set availability

Excel add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, Mac, and iPad. The following table lists the Excel requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

> [!NOTE]
> To use APIs in any of the numbered requirement sets or `ExcelApiOnline`, you should reference the **production** library on the [Office.js content delivery network (CDN)](https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> For information about using preview APIs, see the [Excel JavaScript preview APIs](excel-preview-apis.md) article.

| Requirement set | Office on Windows<br>- Microsoft 365 subscription<br>- retail perpetual Office 2016 and later | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iPad | Office on the web |
|:-----|:-----|:-----|:-----|:-----|:-----|
| [Preview](excel-preview-apis.md)  | Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com)). |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | Not applicable | Not applicable | Not applicable | Not applicable | Latest (see [requirement set page](excel-api-online-requirement-set.md)) |
| [ExcelApi 1.16](excel-api-1-16-requirement-set.md) | Version 2208 (Build 15601.20148) | Not available | 16.64 | 16.66 | Supported |
| [ExcelApi 1.15](excel-api-1-15-requirement-set.md) | Version 2202 (Build 14931.20132) | Not available | 16.58 | 16.59 | Supported |
| [ExcelApi 1.14](excel-api-1-14-requirement-set.md) | Version 2108 (Build 14326.20508) | Office 2021: Version 2108 (Build 14326.20508) | 16.52 | 16.53 | Supported |
| [ExcelApi 1.13](excel-api-1-13-requirement-set.md) | Version 2102 (Build 13801.20738) | Office 2021: Version 2102 (Build 13801.20738) | 16.50 | 16.50 | Supported |
| [ExcelApi 1.12](excel-api-1-12-requirement-set.md) | Version 2008 (Build 13127.20408) | Office 2021: Version 2008 (Build 13127.20408) | 16.40 | 16.40 | Supported |
| [ExcelApi 1.11](excel-api-1-11-requirement-set.md) | Version 2002 (Build 12527.20470) | Office 2021: Version 2002 (Build 12527.20470) | 16.33 | 16.35 | Supported |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | Version 1907 (Build 11929.20306) | Office 2021: Version 1907 (Build 11929.20306) | 16.30 | 16.0 | Supported |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md) | Version 1903 (Build 11425.20204) | Office 2021: Version 1903 (Build 11425.20204) | 16.24 | 16.0 | Supported |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md) | Version 1808 (Build 10730.20102) | Office 2021: Version 1808 (Build 10730.20102) | 16.17 | 16.0 | Supported |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md) | Version 1801 (Build 9001.2171) | Office 2019: Version 1801 (Build 9001.2171) | 16.9  | 16.0  | Supported |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md) | Version 1704 (Build 8201.2001) | Office 2019: Version 1704 (Build 8201.2001) | 15.36  | 15.0 | Supported |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md) | Version 1703 (Build 8067.2070) | Office 2019: Version 1703 (Build 8067.2070) | 15.36  | 15.0 | Supported |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md) | Version 1701 (Build 7870.2024) | Office 2019: Version 1701 (Build 7870.2024) | 15.36  | 15.0 | Supported |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md) | Version 1608 (Build 7369.2055) | Office 2019: Version 1608 (Build 7369.2055) | 15.27 | 15.0 | Supported |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md) | Version 1601 (Build 6741.2088) | Office 2019: Version 1601 (Build 6741.2088) | 15.22 | 15.0 | Supported |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md) | Version 1509 (Build 4266.1001) | Office 2016: Version 1509 (Build 4266.1001) | 15.20 | 15.0 | Supported |

## Office versions and build numbers

For more information about Office versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## How to use Excel requirement sets at runtime and in the manifest

> [!NOTE]
> This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets) and [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.

### Checking for requirement set support at runtime

The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

### Defining requirement set support in the manifest

You can use the [Requirements element](/javascript/api/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins**. If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that do not support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.

The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support ExcelApi requirement set version 1.3 or greater.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
