---
title: Ribbon API requirement sets
description: Specifies which Office platforms and builds support the dynamic ribbon APIs.
ms.date: 09/09/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Ribbon API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

The Ribbon API set supports programmatic control of when custom add-in commands (that is, custom ribbon buttons and menu items) are enabled and disabled and when contextual tabs appear on the ribbon.

Office Add-ins run across multiple versions of Office. The following table lists the Ribbon API requirement sets, the Office client applications that support that requirement set, and the **minimum** builds or versions for those applications.

| Requirement set | Office on Windows<br>(subscription) | Office on Windows<br>(retail perpetual Office 2016 or later) | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iPad | Office on the web | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.2  | Version 2102 (Build 13801.20294) | Version 2102 (Build 13801.20294) | Office 2021: Version 2108 (Build 14326.20454) | 16.53.806.0 | Not supported | Supported | Not supported |
| RibbonApi 1.1  | See [support](#support-for-version-11-in-office-on-windows-subscription)<br>[section](#support-for-version-11-in-office-on-windows-subscription) | Version 2006 (Build 20266.20266) | Office 2021: Version 2108 (Build 14326.20454) | 16.38 | Not supported | Supported | Not supported |

> [!IMPORTANT]
>
> - The RibbonApi requirement sets are supported only on task pane add-ins.
> - The RibbonApi requirement sets are supported for production add-ins only in Excel.
> - RibbonApi 1.1 (not 1.2) is available as a preview in PowerPoint and Word, but only in Office on Windows (subscription) and Office on Mac. It is not available in Office on the web.

## Support for version 1.1 in Office on Windows (subscription)

The 1.1 version of the RibbonApi requirement set is supported in the Consumer Channel from Version 2006 (Build 13001.20498). That requirement set is also supported in the Semi-Annual Channel and Monthly Enterprise Channel builds available from July 14, 2020. The **minimum** supported builds for each channel are as follows:  

| Channel | Minimum version | Minimum build |
|:-----|:-----|:-----|
| Current Channel | 2006 | 20266.20266 |
| Monthly Enterprise Channel | 2005 | 12827.20538 |
| Monthly Enterprise Channel | 2004 | 12730.20602 |
| Semi-Annual Enterprise Channel | 2002 | 12527.20880 |

## More information

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## Ribbon API 1.1

The Ribbon API 1.1 is the first version of the API. For details about the API, see the [Office.ribbon](/javascript/api/office/office.ribbon) reference topic.

## Ribbon API 1.2

The Ribbon API 1.2 adds support for contextual tabs. For more information, see [Create custom contextual tabs in Office Add-ins](/office/dev/add-ins/design/contextual-tabs).

> [!NOTE]
> The **RibbonApi 1.2** requirement set is not yet supported in the manifest, so you shouldn't specify it in the manifest's `<Requirements>` section.

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
