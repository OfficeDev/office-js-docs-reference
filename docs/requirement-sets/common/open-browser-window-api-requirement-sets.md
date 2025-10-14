---
title: Open Browser Window requirement sets
description: Specifies which Office platforms and builds support the openBrowserWindow API.
ms.date: 10/16/2025
ms.topic: overview
ms.localizationpriority: medium
---

# Open Browser Window API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

The OpenBrowserWindow API set enables add-ins to open a browser to accomplish tasks that cannot always be done in the sandboxed webview control within the add-in itself; for example, downloading a PDF file when the webview control is provided by Microsoft Edge.

Office Add-ins run across multiple versions of Office. The following table lists the OpenBrowserWindow API requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual</li></ul> | Office on Windows<ul><li>volume-licensed perpetual/[LTSC](/office/dev/add-ins/resources/resources-glossary#long-term-service-channel-ltsc)</li></ul> | Office on Mac | Office on iOS | Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1 | Not supported | <ul><li>Excel, Outlook (classic), PowerPoint, Word: Version 1810 (Build 11001.20074)</li><li>Outlook ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)): Not supported</li></ul> | Office 2021: Version 2108 (Build 14326.20454) | Version 16.0 | Excel, PowerPoint, Word (iPad only): Version 16.0 | Not supported |

> [!IMPORTANT]
> The OpenBrowserWindowApi 1.1 requirement set is only available as follows:
>
> - Excel, PowerPoint, Word: Windows, Mac, iPad
> - Outlook: Windows (classic), Mac

To find out more about versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## OpenBrowserWindowApi 1.1

The OpenBrowserWindowApi 1.1 is the first version of the API. For details about the API, see the [Office.context.ui](/javascript/api/office/office.context#office-office-context-ui-member) reference topic.

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
