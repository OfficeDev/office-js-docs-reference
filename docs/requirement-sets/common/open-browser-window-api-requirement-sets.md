---
title: Open Browser Window requirement sets
description: Specifies which Office platforms and builds support the openBrowserWindow API.
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Open Browser Window API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

The OpenBrowserWindow API set enables add-ins to open a browser to accomplish tasks that cannot always be done in the sandboxed webview control within the add-in itself; for example, downloading a PDF file when the webview control is provided by Microsoft Edge.

Office Add-ins run across multiple versions of Office. The following table lists the OpenBrowserWindow API requirement sets, the Office host applications that support that requirement set, and the minimum build or version numbers for the Office application. For Windows, new requirement sets usually get deployed with feature updates to Office (subscription) and Office 2016 (retail perpetual) or later and so are available to users who adopt updated builds; typically, new requirement sets don't get deployed to Office 2016 (volume-licensed perpetual) or later, nor to Office 2013.

| Requirement set | Office on Windows | Office on iPad | Office on Mac | Office on the web | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | - Microsoft 365 subscription; Office 2016 (retail perpetual) or later: Version 1810 (Build 16.0.11001.20074)<br><br>- Office 2021 (volume-licensed perpetual) or later: Build 16.0.14326.20454 | 16.0.0.0 | 16.0.0.0 | Not supported | Not supported |

> [!IMPORTANT]
> The OpenBrowserWindowApi requirement set is only available as follows:
>
> - Excel, PowerPoint, Word: Windows, Mac, iPad
> - Outlook: Windows, Mac

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Microsoft 365 Apps](/officeupdates/update-history-microsoft365-apps-by-date)
- [What version of Office am I using?](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office client application](/officeupdates/update-history-microsoft365-apps-by-date)
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## OpenBrowserWindowApi 1.1

The OpenBrowserWindowApi 1.1 is the first version of the API. For details about the API, see the [Office.context.ui](/javascript/api/office/office.context#office-office-context-ui-member) reference topic.

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
