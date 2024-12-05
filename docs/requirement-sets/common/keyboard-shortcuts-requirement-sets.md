---
title: Keyboard Shortcuts requirement sets
description: Keyboard Shortcuts requirement set information for Office Add-ins.
ms.date: 12/05/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Keyboard Shortcuts requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office Add-ins run across multiple versions of Office. The following table lists the Keyboard Shortcuts requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual</li></ul> | Office on Windows<ul><li>volume-licensed perpetual</li></ul> | Office on Mac | Office on iOS | Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| KeyboardShortcuts 1.1 | <ul><li>Excel: Supported</li><li>Word: Rollout in progress*</li></ul> | <ul><li>Excel: Version 2111 (Build 14701.10000)</li><li>Word: Version 2408 (Build 17928.20114)</li></ul> | Not available | <ul><li>Excel: Version 16.55 (21111400)</li><li>Word: Version 16.88 (24081116)</li></ul>| Not supported | Not supported |

> [!NOTE]
> The **KeyboardShortcuts 1.1** requirement set is supported only in **Excel** and **Word**.
>
> \* The keyboard shortcut feature is currently being rolled out to Word on the web. If you test the feature in Word on the web at this time, the shortcuts may not work if they're activated from within the add-in's task pane. We recommend to periodically check this page to find out when the feature is fully supported.

## Office versions and build numbers

To find out more about versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## KeyboardShortcuts 1.1

To learn about the keyboard shortcuts feature, see [Add custom keyboard shortcuts to your Office Add-ins](/office/dev/add-ins/design/keyboard-shortcuts). For details about the APIs in this requirement set, see [Office.actions](/javascript/api/office/office.actions).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
