---
title: Device Permission Service requirement sets
description: Learn more about the Device Permission Service API requirement sets and the platforms it supports.
ms.date: 10/22/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Device Permission Service requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

With the Device Permissions Service requirement set, your add-in can request access to a user's device capabilities. A user's device capabilities include their camera, geolocation, or microphone.

Office Add-ins run across multiple versions of Office. The following table lists the Device Permission Service requirement sets, its supported Office client applications, and the minimum builds or versions for those applications, where applicable.

| Requirement set | Office on the web | Office on Windows<ul><li>Microsoft 365 subscription</li><li>retail perpetual</li><li>volume-licensed perpetual</li></ul> | Office on Mac | Office on iOS | Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|
| DevicePermissionService 1.1 | Chromium-based browsers* | [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) | Not supported | Not supported | Not supported |

> [!NOTE]
> \* DevicePermissionService 1.1 is supported in Office on the web running in Chromium-based browsers, such as Microsoft Edge and Google Chrome.

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## DevicePermissionService 1.1

For details about the API, see [Office.DevicePermission](/javascript/api/office/office.devicepermission).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
