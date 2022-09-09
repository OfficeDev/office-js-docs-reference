---
title: Open Browser Window requirement sets
description: Specifies which Office platforms and builds support the openBrowserWindow API.
ms.date: 09/09/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Open Browser Window API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

The OpenBrowserWindow API set enables add-ins to open a browser to accomplish tasks that cannot always be done in the sandboxed webview control within the add-in itself; for example, downloading a PDF file when the webview control is provided by Microsoft Edge.

Office Add-ins run across multiple versions of Office. The following table lists the OpenBrowserWindow API requirement sets, Office client applications that support that requirement set, and the **minimum** builds or versions for those applications.

| Requirement set | Office on Windows<br>(subscription) | Office on Windows<br>(retail perpetual Office 2016 or later) | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iPad | Office on the web | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1 | Version 1810 (Build 11001.20074) | Version 1810 (Build 11001.20074) | Office 2021: Version 2108 (Build 14326.20454) | 16.0.0.0 | 16.0.0.0 | Not supported | Not supported |

> [!IMPORTANT]
> The OpenBrowserWindowApi 1.1 requirement set is only available as follows:
>
> - Excel, PowerPoint, Word: Windows, Mac, iPad
> - Outlook: Windows, Mac

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## OpenBrowserWindowApi 1.1

The OpenBrowserWindowApi 1.1 is the first version of the API. For details about the API, see the [Office.context.ui](/javascript/api/office/office.context#office-office-context-ui-member) reference topic.

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
