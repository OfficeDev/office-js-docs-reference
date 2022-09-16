---
title: PowerPoint JavaScript API requirement sets
description: Learn more about the PowerPoint JavaScript API requirement sets.
ms.date: 06/13/2022
ms.prod: powerpoint
ms.localizationpriority: high
---

# PowerPoint JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the **minimum** builds or versions for those applications.

| Requirement set | Office on Windows<br>(subscription) | Office on Windows<br>(retail perpetual Office 2016 or later) | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iPad | Office on the web |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| [PowerPointApi 1.4](powerpoint-api-1-4-requirement-set.md) | Version 2207 (Build 15330.20122) | Version 2207 (Build 15330.20122) | Not available | 16.62 | Not available | Supported |
| [PowerPointApi 1.3](powerpoint-api-1-3-requirement-set.md) | Version 2111 (Build 14701.20060) | Version 2111 (Build 14701.20060) | Not available | 16.55 | Not available | Supported |
| [PowerPointApi 1.2](powerpoint-api-1-2-requirement-set.md) | Version 2011 (Build 13426.20184) | Version 2011 (Build 13426.20184) | Office 2021: Version 2011 (Build 13426.20184) | 16.43 | Not available | Supported |
| [PowerPointApi 1.1](powerpoint-api-1-1-requirement-set.md) | Version 1810 (Build 11001.20074) | Version 1810 (Build 11001.20074) | Office 2021: Version 1810 (Build 11001.20074) | 16.19 | 2.17 | Supported |

## Office versions and build numbers

For more information about Office versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## PowerPoint JavaScript API 1.1

PowerPoint JavaScript API 1.1 contains a [single API to create a new presentation](/javascript/api/powerpoint#PowerPoint_createPresentation_base64File_). For details about the API, see [Create a presentation](/office/dev/add-ins/powerpoint/powerpoint-add-ins#create-a-presentation).

## PowerPoint JavaScript API 1.2

PowerPoint JavaScript API 1.2 adds support for inserting slides from another PowerPoint presentation into the current presentation and for deleting slides. For details about the APIs, see [Insert and delete slides in a PowerPoint presentation](/office/dev/add-ins/powerpoint/insert-slides-into-presentation).

## PowerPoint JavaScript API 1.3

PowerPoint JavaScript API 1.3 adds additional support for adding and deleting slides. It also lets add-ins apply custom metadata tags. For details about the APIs, see [Add and delete slides in PowerPoint](/office/dev/add-ins/powerpoint/add-slides) and [Use custom tags for presentations, slides, and shapes in PowerPoint](/office/dev/add-ins/powerpoint/tagging-presentations-slides-shapes).

## PowerPoint JavaScript API 1.4

PowerPoint JavaScript API 1.4 adds additional support for adding, moving, sizing, formatting, and deleting shapes. For more information about using these APIs, see [Working with shapes](/office/dev/add-ins/powerpoint/shapes).

## How to use PowerPoint requirement sets at runtime and in the manifest

> [!NOTE]
> This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets) and [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.

### Checking for requirement set support at runtime

The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### Defining requirement set support in the manifest

You can use the [Requirements element](/javascript/api/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins**. If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that don't support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.

The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support PowerPointApi requirement set version 1.1 or greater.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## Office Common API requirement sets

Most of the PowerPoint add-in functionality comes from the Common API set. For information about Common API requirement sets, see [Office Common API requirement sets](../common/office-add-in-requirement-sets.md).

## See also

- [PowerPoint JavaScript API reference documentation](/javascript/api/powerpoint)
- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
