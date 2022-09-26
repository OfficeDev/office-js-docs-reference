---
title: Add-in commands requirement sets
description: Overview of Office Add-in commands requirement sets.
ms.date: 09/26/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Add-in commands requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands for Excel, Word, and PowerPoint](/office/dev/add-ins/design/add-in-commands) and [Add-in commands for Outlook](/office/dev/add-ins/outlook/add-in-commands-for-outlook).

> [!NOTE]
> Outlook add-ins support add-in commands, but the APIs and manifest elements that enable add-in commands in Outlook are in the [Mailbox 1.3](../outlook/requirement-set-1.3/outlook-requirement-set-1.3.md) requirement set. The AddinCommands requirement sets are not applicable to Outlook.

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddinCommands 1.0 requirement set). The following table lists the Office client applications that support the initial release version, and the **minimum** builds or versions for those applications.  

| Release | Office on Windows<br>- subscription<br>- retail perpetual Office 2016 and later | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iPad | Office on the web |
|:-----|:-----|:-----|:-----|:-----|:-----|
| Add-in commands (initial release, no requirement set) | Version 1603 (Build 6769.0000) | Office 2021: Version 1809 (Build 10827.20150) | 15.33 | Not supported | Supported |

The add-in commands **1.1** requirement set introduces the ability to [autoopen a task pane with documents](/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

The add-in commands **1.3** requirement set introduces manifest markup that enables an add-in to customize the placement of a custom tab on the Office ribbon and to insert built-in Office ribbon controls into custom control groups.

The following table lists the add-in commands requirement sets, the Office client applications that support that requirement set, and the **minimum** builds or versions for those applications.

| Requirement set | Office on Windows<br>- subscription<br>- retail perpetual Office 2016 and later | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iPad | Office on the web |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.3 | Version 2204 (Build 14827.10000) | Not supported | 16.57.105.0 | Not supported | Supported |
| AddinCommands 1.1 | Version 1705 (Build 8121.1000)&dagger; | Office 2021: Version 1809 (Build 10827.20150)&dagger; | 15.34&dagger;\* | Not supported | Supported |

> \* The [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) method will erroneously return `false` for versions 16.9 &ndash; 16.14 (inclusive), but the requirement set *is* supported on these versions.
>
> &dagger; OneNote is supported only in Office on the web.

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
