---
title: Ribbon API requirement sets
description: Specifies which Office platforms and builds support the dynamic ribbon APIs.
ms.date: 10/22/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Ribbon API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

The Ribbon API set supports programmatic control of when custom add-in commands (that is, custom ribbon buttons and menu items) are enabled and disabled and when contextual tabs appear on the ribbon.

## Support

`RibbonApi 1.1` is available with **Excel**, **PowerPoint**, and **Word**.  `RibbonApi 1.2` is only available with **Excel**. Both are only for use with **task pane add-ins**. The following table lists the Ribbon API requirement sets, the supported platforms, and the **minimum** builds or versions where applicable.

| Requirement set | Office on the web | Office on Windows<br>(Microsoft 365 subscription) | Office on Windows<br>(retail perpetual) | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iOS | Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.2 | Supported | Version 2102 (Build 13801.20294) | Version 2102 (Build 13801.20294) | Office 2021: Version 2108 (Build 14326.20454) | Version 16.53 (21080600) | Not supported | Not supported |
| RibbonApi 1.1 | Supported | Version 2002 (12527.20880) | Version 2006 (Build 13001.20266) | Office 2021: Version 2108 (Build 14326.20454) | Version 16.38 (20061401) | Not supported | Not supported |

To find out more about versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## Ribbon API 1.1

The Ribbon API 1.1 includes support for enabling and disabling Add-in Commands. To learn the patterns for this functionality, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands). For details about the API, see the [Office.ribbon](/javascript/api/office/office.ribbon) reference topic.

## Ribbon API 1.2

The Ribbon API 1.2 adds support for contextual tabs. For more information, see [Create custom contextual tabs in Office Add-ins](/office/dev/add-ins/design/contextual-tabs).

> [!NOTE]
> The **RibbonApi 1.2** requirement set isn't yet supported in the manifest, so you shouldn't specify it in the manifest's **\<Requirements\>** section.

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
