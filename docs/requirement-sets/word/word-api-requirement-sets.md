---
title: Word JavaScript API requirement sets
description: Office Add-in requirement set information for Word.
ms.date: 12/09/2024
ms.topic: overview
ms.localizationpriority: high
---

# Word JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## Requirement set availability

Word add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, iPad, and Mac. The following table lists the Word requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

> [!NOTE]
> To use APIs in any of the numbered requirement sets, `WordApiOnline`, `WordApiDesktop`, or `WordApiHiddenDocument`, you should reference the **production** library on the [Office.js content delivery network (CDN)](https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> For information about using preview APIs, see the [Word JavaScript preview APIs](word-preview-apis.md) article.

| Requirement set | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual</li></ul> | Office on Windows<ul><li>volume-licensed perpetual</li></ul> | Office on Mac | Office on iPad |
|:-----|:-----|:-----|:-----|:-----|:-----|
| [Preview](word-preview-apis.md) | Please use the latest Office version to try preview APIs (you may need to join the [Microsoft 365 Insider program](https://insider.microsoft365.com/join)) |
| [WordApiOnline 1.1](word-api-online-requirement-set.md) | Latest (see [requirement set page](word-api-online-requirement-set.md)) | Not applicable | Not applicable | Not applicable | Not applicable |
| [WordApiDesktop 1.1](word-api-desktop-1.1-requirement-set.md) | Not applicable | Version 2408 (Build 17928.20114) | Office 2024: Version 2408 (Build 17928.20114) | Version 16.88 (24081116) | Version 16.88 |
| [WordApiHiddenDocument 1.5](word-api-1.5-hidden-document-requirement-set.md) (Desktop only) | Not applicable | Version 2302 (Build 16130.20332) | Office 2024: Version 2302 (Build 16130.20332) | Version 16.70 (23021201) | Not applicable |
| [WordApiHiddenDocument 1.4](word-api-1.4-hidden-document-requirement-set.md) (Desktop only) | Not applicable | Version 2208 (Build 15601.20148) | Office 2024: Version 2208 (Build 15601.20148) | Version 16.64 (22081401) | Not applicable |
| [WordApiHiddenDocument 1.3](word-api-1.3-hidden-document-requirement-set.md) (Desktop only) | Not applicable | Version 1612 (Build 7668.1000) | Office 2019: Version 1612 (Build 7668.1000) | Version 15.32 (17030901) | Not applicable |
| [WordApi 1.9](word-api-1-9-requirement-set.md) | Supported | Version 2411 (Build 18227.20152) | Not available | Version 16.91 (24111020) | Version 16.91 |
| [WordApi 1.8](word-api-1-8-requirement-set.md) | Supported | Version 2405 (Build 17628.20110) | Office 2024: Version 2405 (Build 17628.20110) | Version 16.85 (24051214) | Version 16.85 |
| [WordApi 1.7](word-api-1-7-requirement-set.md) | Supported | Version 2311 (Build 17029.20068) | Office 2024: Version 2311 (Build 17029.20068) | Version 16.79 (23111019) | Version 16.79 |
| [WordApi 1.6](word-api-1-6-requirement-set.md) | Supported | Version 2308 (Build 16731.20234) | Office 2024: Version 2308 (Build 16731.20234) | Version 16.76 (23081101) | Version 16.76 |
| [WordApi 1.5](word-api-1-5-requirement-set.md) | Supported | Version 2302 (Build 16130.20332) | Office 2024: Version 2302 (Build 16130.20332) | Version 16.70 (23021201) | Version 16.70 |
| [WordApi 1.4](word-api-1-4-requirement-set.md) | Supported | Version 2208 (Build 15601.20148) | Office 2024: Version 2208 (Build 15601.20148) | Version 16.64 (22081401) | Version 16.64 |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | Supported | Version 1612 (Build 7668.1000) | Office 2019: Version 1612 (Build 7668.1000) | Version 15.32 (17030901) | Version 2.22 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | Supported | Version 1601 (Build 6568.1000) | Office 2019: Version 1601 (Build 6568.1000) | Version 15.19 | Version 1.18 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | Supported | Version 1509 (Build 4266.1001) | Office 2016: Version 1509 (Build 4266.1001) | Version 15.19 | Version 1.18 |

## Office versions and build numbers

For more information about Office versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
