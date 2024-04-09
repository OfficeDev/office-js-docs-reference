---
title: Word JavaScript API requirement sets
description: Office Add-in requirement set information for Word.
ms.date: 04/09/2024
ms.topic: overview
ms.localizationpriority: high
---

# Word JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## Requirement set availability

Word add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, iPad, and Mac. The following table lists the Word requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

> [!NOTE]
> To use APIs in any of the numbered requirement sets, `WordApiOnline`, or `WordApiHiddenDocument`, you should reference the **production** library on the [Office.js content delivery network (CDN)](https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> For information about using preview APIs, see the [Word JavaScript preview APIs](word-preview-apis.md) article.

| Requirement set | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual Office 2016 and later</li></ul> | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iPad |
|:-----|:-----|:-----|:-----|:-----|:-----|
| [Preview](word-preview-apis.md) | Please use the latest Office version to try preview APIs (you may need to join the [Microsoft 365 Insider program](https://insider.microsoft365.com/join)) |
| [WordApiOnline 1.1](word-api-online-requirement-set.md) | Latest (see [requirement set page](word-api-online-requirement-set.md)) | Not applicable | Not applicable | Not applicable | Not applicable |
| [WordApiHiddenDocument 1.5](word-api-1.5-hidden-document-requirement-set.md) (Desktop only) | Not applicable | Version 2302 (Build 16130.20332) | Not available | 16.70 | Not applicable |
| [WordApiHiddenDocument 1.4](word-api-1.4-hidden-document-requirement-set.md) (Desktop only) | Not applicable | Version 2208 (Build 15601.20148) | Not available | 16.64 | Not applicable |
| [WordApiHiddenDocument 1.3](word-api-1.3-hidden-document-requirement-set.md) (Desktop only) | Not applicable | Version 1612 (Build 7668.1000) | Office 2019: Version 1612 (Build 7668.1000) | 15.32 | Not applicable |
| [WordApi 1.7](word-api-1-7-requirement-set.md) | Supported | Version 2311 (Build 17029.20068) | Not available | 16.79 | Not available |
| [WordApi 1.6](word-api-1-6-requirement-set.md) | Supported | Version 2308 (Build 16731.20234) | Not available | 16.76 | 16.76 |
| [WordApi 1.5](word-api-1-5-requirement-set.md) | Supported | Version 2302 (Build 16130.20332) | Not available | 16.70 | 16.70 |
| [WordApi 1.4](word-api-1-4-requirement-set.md) | Supported | Version 2208 (Build 15601.20148) | Not available | 16.64 | 16.64 |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | Supported | Version 1612 (Build 7668.1000) | Office 2019: Version 1612 (Build 7668.1000) | 15.32 | 2.22 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | Supported | Version 1601 (Build 6568.1000) | Office 2019: Version 1601 (Build 6568.1000) | 15.19 | 1.18 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | Supported | Version 1509 (Build 4266.1001) | Office 2016: Version 1509 (Build 4266.1001) | 15.19 | 1.18 |

## Office versions and build numbers

For more information about Office versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
