---
title: Set element in the manifest file
description: The Set element specifies an Office JavaScript API requirement set your Office Add-in requires in order to be activated by Office or to override base manifest settings.
ms.date: 06/29/2022
ms.localizationpriority: medium
---

# Set element

The meaning of this element depends on where it's used in the manifest.

## In the base manifest

When used in the base manifest (that is, the grandparent **Requirements** element is a direct child of [OfficeApp](officeapp.md)), the **Set** element specifies a [requirement set](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets) from the Office JavaScript API that your Office Add-in needs in order to be activated by Office.

**Add-in type:** Content, Task pane, Mail

## As a great-grandchild of a VersionOverrides element

Specifies a [requirement set](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets) from the Office JavaScript API that must be supported by the Office version and platform (such as Windows, Mac, web, and iOS or iPad) in order for the [VersionOverrides](versionoverrides.md) to take effect.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Same as the grandparent [Requirements](requirements.md) element.

**Associated with these requirement sets**:

- Same as the grandparent [Requirements](requirements.md) element.

## Syntax

```XML
<Set Name="string" MinVersion="n .n">
```

## Contained in

- [Sets](sets.md)

## Attributes

|Attribute|Type|Required|Description|
|:-----|:-----:|:-----:|:-----|
|Name|string|Yes|The name of a [requirement set](/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|string|No|Specifies the minimum version of the API set required by your add-in. Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.|

## Remarks

Certain requirement sets can't be declared in this element of the manifest; they're listed in the following table. In those cases, you should do a runtime check to determine if the user's version of Office supports your target requirement set. To learn more about how to do so, see [Runtime checks for method and requirement set support](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#runtime-checks-for-method-and-requirement-set-support).

|Requirement set|Affected hosts|
|:---|:---|
|[ExcelApiOnline 1.1](../requirement-sets/excel/excel-api-online-requirement-set.md#recommended-usage)|Excel|
|[IdentityApi 1.3](../requirement-sets/common/identity-api-requirement-sets.md#outlook-and-identity-api-requirement-sets)|Outlook|
|[WordApiOnline 1.1](../requirement-sets/word/word-api-online-requirement-set.md#recommended-usage)|Word|
|[Preview APIs](/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn#preview-apis)|All hosts|

For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Specify which Office versions and platforms can host your add-in](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#specify-which-office-versions-and-platforms-can-host-your-add-in).
