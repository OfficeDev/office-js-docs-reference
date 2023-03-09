---
title: Dialog Origin requirement sets
description: Learn more about the Dialog Origin requirement sets.
ms.date: 09/28/2022
ms.topic: overview
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Dialog Origin requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office Add-ins run across multiple versions of Office. The following table lists the Dialog Origin requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Office on Windows<br>- Microsoft 365 subscription<br>- retail perpetual Office 2016 and later | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iPad | Office on the web | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogOrigin 1.1\* | Supported | Office 2013: Build 15.0.5371.1000 | 16.52 | 2.52 | Supported | Version 2108 (Build 10377.1000) |

> [!NOTE]
> \* Retail perpetual Office 2013 on Windows also supports DialogOrigin 1.1.

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## Dialog Origin 1.1

The Dialog Origin 1.1 is the first version of the API. It provides support for cross-domain messaging between a dialog and its parent page. For details about these APIs, see the [Office.ui](/javascript/api/office/office.ui) reference topic.

## See also

- [Use the Office dialog API in Office Add-ins](/office/dev/add-ins/develop/dialog-api-in-office-add-ins)
- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
