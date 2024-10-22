---
title: Shared runtime requirement sets
description: Specifies the platforms and Office applications that support the SharedRuntime APIs.
ms.date: 10/22/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Shared runtime requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Parts of an Office Add-in that run JavaScript code, such as task panes, function files launched from add-in commands, and Excel custom functions, can share a single JavaScript runtime. This enables all the parts to share a set of global variables, to share a set of loaded libraries, and to communicate with each other without having to pass messages through a persisted storage. For more information, see [Configure your Office Add-in to use a shared JavaScript runtime](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime).

The following table lists the Shared Runtime requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual</li></ul> | Office on Windows<ul><li>volume-licensed perpetual</li></ul> | Office on Mac | Office on iOS | Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.2 | Excel: Supported | Excel: Version 2108 (Build 14326.20508) | Excel 2024: Version 2108 (Build 14326.20508) | Excel: Version 16.52 (21080801) | Not supported | Not supported |
| SharedRuntime 1.1  | Excel, PowerPoint, Word: Supported | <ul><li>Excel: Version 2002 (Build 12527.20092)</li><li>PowerPoint: Version 2102 (Build 13722.10000)</li><li>Word: Version 2205 (Build 15202.10000)</li></ul> | <ul><li>Excel 2021: Version 2108 (Build 12527.20092)</li><li>PowerPoint 2021: Version 2108 (13722.10000)</li><li>Word 2024: Version 2205 (Build 15202.10000)</li></ul> | <ul><li>Excel: Version 16.35 (20030802)</li><li>PowerPoint: Version 16.46 (21012000)</li><li>Word: 16.61 (22040100)</li></ul> | Not supported | Not supported |

## Office versions and build numbers

To find out more about versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## SharedRuntime API 1.1

The SharedRuntime API 1.1 is the first version of the API. For details, see the [Office.Addin](/javascript/api/office/office.addin) reference topic.

## SharedRuntime API 1.2

The SharedRuntime API 1.2 adds the [Office.BeforeDocumentCloseNotification](/javascript/api/office/office.beforedocumentclosenotification) interface, which helps ensure that workbooks don't close while an add-in process is running.

> [!IMPORTANT]
> SharedRuntime 1.2 is only supported in Excel.

## See also

- [Configure your Office Add-in to use a shared JavaScript runtime](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
