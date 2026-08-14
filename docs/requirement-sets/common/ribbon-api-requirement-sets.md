---
title: Ribbon API requirement sets
description: Specifies which Office platforms and builds support the dynamic ribbon APIs.
ms.date: 08/13/2026
ms.topic: overview
ms.localizationpriority: medium
---

# Ribbon API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

The Ribbon API requirement sets let add-ins programmatically control custom ribbon elements. Add-ins can:

- Enable or disable add-in commands.
- Show contextual tabs on the ribbon.
- Show or hide commands on a custom tab.

## Support

The following table lists the Ribbon API requirement sets, supported applications and platforms, and the **minimum** builds or versions where applicable.

| Requirement set | Applications | Office on the web | Office on Windows<br>(Microsoft 365 subscription) | Office on Windows<br>(retail perpetual) | Office on Windows<br>(volume-licensed perpetual/[LTSC](/office/dev/add-ins/resources/resources-glossary#long-term-service-channel-ltsc)) | Office on Mac | Office on iOS | Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.3 | <ul><li>Excel</li><li>PowerPoint</li><li>Word</li></ul> | Supported | Version 2606 (Build 20112.15190) | Version 2606 (Build 20112.15190) | Not supported | Version 16.109.1 (260512.115) | Not supported | Not supported |
| RibbonApi 1.2 | <ul><li>Excel</li></ul> | Supported | Version 2102 (Build 13801.20294) | Version 2102 (Build 13801.20294) | Office 2021: Version 2108 (Build 14326.20454) | Version 16.53 (21080600) | Not supported | Not supported |
| RibbonApi 1.1 | <ul><li>Excel</li><li>PowerPoint</li><li>Word</li></ul> | Supported | Version 2002 (12527.20880) | Version 2006 (Build 13001.20266) | Office 2021: Version 2108 (Build 14326.20454) | Version 16.38 (20061401) | Not supported | Not supported |

> [!NOTE]
> `RibbonApi 1.1` and `RibbonApi 1.2` are only for use with task pane add-ins.

To find out more about versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## Ribbon API 1.1

Ribbon API 1.1 includes support for enabling and disabling Add-in Commands. To learn the patterns for this functionality, see [Change the availability of add-in commands](/office/dev/add-ins/design/disable-add-in-commands). For details about the API, see the [Office.ribbon](/javascript/api/office/office.ribbon) reference topic.

## Ribbon API 1.2

Ribbon API 1.2 adds support for contextual tabs. For more information, see [Create custom contextual tabs in Office Add-ins](/office/dev/add-ins/design/contextual-tabs).

> [!NOTE]
> The **RibbonApi 1.2** requirement set isn't yet supported in the manifest, so you shouldn't specify it in the manifest's **\<Requirements\>** section.

## Ribbon API 1.3

Ribbon API 1.3 adds support to dynamically show or hide add-in commands on a custom tab. For more information, see [Show or hide add-in commands on a custom tab](/office/dev/add-ins/design/show-hide-controls-custom-tab).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
