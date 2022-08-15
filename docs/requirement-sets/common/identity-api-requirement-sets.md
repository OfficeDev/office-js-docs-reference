---
title: Identity API requirement sets
description: Identity API requirement set information for Office Add-ins.
ms.date: 08/12/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Identity API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office Add-ins run across multiple versions of Office. The following table lists the Identity API requirement sets, the Office client applications that support that requirement set, and the minimum build or version numbers for the Office application primarily when connected to a Microsoft 365 subscription.

|  Requirement set  | Office on Windows\* |  Office on iPad  |  Office on Mac   | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | Version 2008 (Build 13127.20000) | Not supported | 16.40 | Microsoft SharePoint Online and OneDrive\*\* |

\* Windows: New requirement sets usually get deployed with feature updates to Office (subscription) and Office 2016 (retail perpetual) or later and so are available to users who adopt updated builds. Typically, new requirement sets don't get deployed to Office 2016 (volume-licensed perpetual) or later, nor to Office 2013. IdentityAPI 1.3 is available in Office 2021 (volume-licensed perpetual) or later from Build 16.0.14326.20454.

\*\* Currently, the requirement set is supported in Office on the web only for documents that are opened from Microsoft SharePoint Online and OneDrive.

## Outlook and Identity API requirement sets

[!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]

> [!NOTE]
> In an Outlook add-in using event-based activation, the [OfficeRuntime.Auth interface](/javascript/api/office-runtime/officeruntime.auth) is supported in Outlook version 2108 (build 14326.20258) or later on Windows. The [Office.Auth interface](/javascript/api/office/office.auth) is supported in version 2111 (build 14701.20000) or later. For more details according to your version, see the update history page for [Office 2021](/officeupdates/update-history-office-2021) or [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) and how to [find your Office client version and update channel](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

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
