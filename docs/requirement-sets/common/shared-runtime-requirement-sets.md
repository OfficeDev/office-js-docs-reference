---
title: Shared runtime requirement sets
description: Specifies the platforms and Office applications that support the SharedRuntime APIs.
ms.date: 07/06/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Shared runtime requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Parts of an Office Add-in that run JavaScript code, such as task panes, function files launched from add-in commands, and Excel custom functions, can share a single JavaScript runtime. This enables all the parts to share a set of global variables, to share a set of loaded libraries, and to communicate with each other without having to pass messages through a persisted storage. For more information, see [Configure your Office Add-in to use a shared JavaScript runtime](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime).

The following table lists the SharedRuntime 1.1 requirement set, the Office client applications that support that requirement set, and the build or version numbers for the Office application.

| Requirement set | Office 2021 or later on Windows<br>(one-time purchase) | Office on Windows<br>(connected to a Microsoft 365 subscription) | Office on iPad<br>(connected to a Microsoft 365 subscription) | Office on Mac<br>(both subscription<br> and one-time purchase Office on Mac 2019 and later)  | Office on the web | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | **Excel:** Build 16.0.14326.20454 or later<br> **Word:** ???<br> **PowerPoint:** ??? | **Excel:** Version 2002 (Build 12527.20092) or later<br> **Word:** ???<br> **PowerPoint:** ??? | N/A | **Excel:** 16.35 or later<br> **Word:** Build 16.46.207.0 or later<br> **PowerPoint:** Build 16.46.207.0 or later | **Excel:** February 2020<br> **Word:** July 2022<br> **PowerPoint:** February 2022 | N/A |

> [!IMPORTANT]
> At this time, the shared JavaScript runtime is not supported on iPad or in one-time purchase versions of Office 2019 or earlier. For additional support details, see the following sections.

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Configure your Office Add-in to use a shared JavaScript runtime](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
