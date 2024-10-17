---
title: Nested app auth requirement sets
description: Nested app auth requirement set information for Office Add-ins.
ms.date: 10/15/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Nested app auth requirement set

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office Add-ins run across multiple versions of Office. The following table lists the nested app auth requirement set, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual</li></ul> | Office on Windows<ul><li>volume-licensed perpetual</li></ul> | Office on Mac | Office on iPad and Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|
| NestedAppAuth 1.1 | Supported | Version 2409 (Build 18025.20000) | not available | Version 16.89 (Build 24090815) | Build v4.2433.0 |

> [!IMPORTANT]
>
> - Currently, the NestedAppAuth 1.1 requirement set is supported in Office on the web only for documents that are opened from Microsoft SharePoint Online and OneDrive.
> - In Outlook, the NestedAppAuth 1.1 requirement set isn't supported if the add-in is loaded in an Outlook.com or Gmail mailbox.

## Outlook and NestedAppAuth requirement set

To require the NestedAppAuth requirement set 1.1 in your Outlook add-in code, check if it's supported by calling `isSetSupported('NestedAppAuth', '1.1')`.
Declaring it in the Outlook add-in's manifest isn't supported. You can also determine if the API is supported by checking that it's not `undefined`.
For further details, see [Using APIs from later requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#using-apis-from-later-requirement-sets).

## Office versions and build numbers

To find out more about versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
- [Enable SSO in an Office Add-in using nested app authentication](/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in)
