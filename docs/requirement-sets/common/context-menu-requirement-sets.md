---
title: Context Menu requirement sets
description: Learn more about the Context Menu API requirement sets and the platforms it supports.
ms.date: 01/23/2025
ms.topic: overview
ms.localizationpriority: medium
---

# Context Menu requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Use the Context Menu requirement set to configure the availability of custom items on a context menu in Office.

The following table lists the Context Menu requirement sets, its supported Office client applications, and the minimum builds or versions for those applications, where applicable.

| Requirement set | Office on the web | Office on Windows<br>(Microsoft 365 subscription) | Office on Windows<br>(retail perpetual) | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iOS | Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| ContextMenu 1.1 | Supported | TBD | TBD | Not supported | TBD | Not supported | Not supported |

To find out more about versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## ContextMenu 1.1

To learn how to programmatically configure the availability of custom items on a context menu, see [Enable or disable add-in commands](/office/dev/add-ins/design/disable-add-in-commands). For details about the API, see [Office.ContextMenu](/javascript/api/office/office.contextmenu).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
