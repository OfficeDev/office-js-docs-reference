---
title: Task Pane API requirement sets
description: Learn more about the Task Pane API requirement sets and the platforms it supports.
ms.date: 11/18/2025
ms.topic: overview
ms.localizationpriority: medium
---

# Task Pane API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

With the Task Pane API requirement set, you can manage the task pane of an add-in. For example, you can change the width of an add-in's task pane.

## Support

`TaskPaneApi 1.1` is available with **Excel**, **PowerPoint**, and **Word**. The following table lists the Task Pane API requirement sets, its supported Office client applications, and the minimum builds or versions for those applications, where applicable.

| Requirement set | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual</li><li>volume-licensed perpetual</li></ul> | Office on Mac | Office on iOS | Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|
| TaskPaneApi 1.1 | <ul><li>**Excel**: Supported</li><li>**PowerPoint**: Not supported</li><li>**Word**: Supported</li></ul> | Version 2507 (Build 19029.20004) | Version 16.100.4 (Build 25083118) | Not supported | Not supported |

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## TaskPaneApi 1.1

For details about the API, see [Office.TaskPane](/javascript/api/office/office.taskpane).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
