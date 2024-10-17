---
title: Add-in commands requirement sets
description: Overview of Office Add-in commands requirement sets.
ms.date: 04/15/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Add-in commands requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands](/office/dev/add-ins/design/add-in-commands) and [Create add-in commands](/office/dev/add-ins/develop/create-addin-commands).

> [!NOTE]
> Outlook add-ins support add-in commands, but the APIs and manifest elements that enable add-in commands in Outlook are in the [Mailbox 1.3](../outlook/requirement-set-1.3/outlook-requirement-set-1.3.md) requirement set. The AddinCommands requirement sets aren't applicable to Outlook.

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddinCommands 1.0 requirement set). The following table lists the Office client applications that support the initial release version, and the **minimum** builds or versions for those applications where applicable.  

| Release | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual</li></ul> | Office on Windows<ul><li>volume-licensed perpetual</li></ul> | Office on Mac | Office on iPad |
|:-----|:-----|:-----|:-----|:-----|:-----|
| Add-in commands (initial release, no requirement set) | Supported | Version 1603 (Build 6769.0000) | Office 2021: Version 1809 (Build 10827.20150) | Version 15.33 (17040900) | Not supported |

The AddinCommands **1.1** requirement set introduces the ability to [autoopen a task pane with documents](/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

The AddinCommands **1.3** requirement set introduces manifest markup that enables an add-in to customize the placement of a custom tab on the Office ribbon and to insert built-in Office ribbon controls into custom control groups. At present, PowerPoint is the only Office client that supports this requirement set.

The following table lists the add-in commands requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual</li></ul> | Office on Windows<ul><li>volume-licensed perpetual</li></ul> | Office on Mac | Office on iPad |
|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.3 | PowerPoint: Supported | PowerPoint: Version 2204 (Build 14827.10000) | Not supported | PowerPoint: 16.57.105.0 | Not supported |
| AddinCommands 1.1 | Supported | Version 1705 (Build 8121.1000)&dagger; | Office 2021: Version 1809 (Build 10827.20150)&dagger; | Version 15.34 (17051500)&dagger;\* | Not supported |

> \* The [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) method will erroneously return `false` for versions 16.9 &ndash; 16.14 (inclusive), but the requirement set *is* supported on these versions.
>
> &dagger; OneNote is supported only in Office on the web.

## Office versions and build numbers

To find out more about versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
