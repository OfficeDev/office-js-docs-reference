---
title: OneNote JavaScript API requirement sets
description: Learn more about the OneNote JavaScript API requirement sets.
ms.date: 09/28/2022
ms.topic: overview
ms.localizationpriority: high
---

# OneNote JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

The following table lists the OneNote requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Office on the web |
|:-----|:-----|
| [OneNoteApi 1.1](/javascript/api/onenote?view=onenote-js-1.1&preserve-view=true) | Supported |

## OneNote JavaScript API 1.1

OneNote JavaScript API 1.1 is the first version of the API. For details about the API, see the [OneNote JavaScript API programming overview](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).

## Runtime requirement support check

At runtime, add-ins can check if a particular Office application supports an API requirement set by doing the following:

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## Manifest-based requirement support check

Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in My Add-ins.

The following code example shows an add-in that loads in all Office client applications that support the OneNoteApi requirement set, version 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](../common/office-add-in-requirement-sets.md).

## See also

- [OneNote JavaScript API reference documentation](/javascript/api/onenote)
- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
