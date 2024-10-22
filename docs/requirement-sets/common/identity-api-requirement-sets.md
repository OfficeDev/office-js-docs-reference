---
title: Identity API requirement sets
description: Identity API requirement set information for Office Add-ins.
ms.date: 10/22/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Identity API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office Add-ins run across multiple versions of Office. The following table lists the Identity API requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual</li></ul> | Office on Windows<ul><li>volume-licensed perpetual</li></ul> | Office on Mac | Office on iOS | Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3 | Microsoft SharePoint Online and OneDrive\* | Version 2008 (Build 13127.20000) | Office 2021: Version 2108 (Build 14326.20454) | Version 16.40 (20081000) | Not supported | Not supported |

> \* Currently, the IdentityAPI 1.3 requirement set is supported in Office on the web only for documents that are opened from Microsoft SharePoint Online and OneDrive.

> [!IMPORTANT]
>
> - In Outlook, the Identity API requirement set isn't supported if the add-in is loaded in an Outlook.com or Gmail mailbox.

## Outlook and Identity API requirement sets

[!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]

> [!NOTE]
>
> - IdentityAPI 1.3 isn't supported in Outlook on Android or on iOS.
> - In an Outlook add-in using event-based activation, the [OfficeRuntime.Auth interface](/javascript/api/office-runtime/officeruntime.auth) is supported in Outlook from Version 2108 (Build 14326.20258) on Windows. The [Office.Auth interface](/javascript/api/office/office.auth) is supported from Version 2111 (Build 14701.20000). For more details according to your version, see the update history page for [Office 2021](/officeupdates/update-history-office-2021) or [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) and how to [find your Office client version and update channel](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

## Office versions and build numbers

To find out more about versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
