---
title: SourceLocation element (version overrides) in the manifest file
description: Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel, or needed by the DetectedEntity or LaunchEvent extension points in Outlook.
ms.date: 06/14/2022
ms.localizationpriority: medium
---

# SourceLocation element (version overrides)

Defines the location of a resource needed by the **Script** or **Page** elements used by custom functions in Excel, or needed by the **DetectedEntity** or **LaunchEvent** extension points in Outlook.

> [!IMPORTANT]
> This article refers only to the **SourceLocation** that is a child of the **Page** or **Script** elements, or of the **DetectedEntity** or **LaunchEvent** extension points. See [SourceLocation](sourcelocation.md) for information about the **SourceLocation** element of the base manifest.

**Add-in type:** Custom function, Mail

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)
- [Mailbox 1.6](../requirement-sets/outlook/requirement-set-1.6/outlook-requirement-set-1.6.md)
- [Mailbox 1.10](../requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10.md)

## Contained in

- [ExtensionPoint](extensionpoint.md) ([Contextual](extensionpoint.md#detectedentity) and [LaunchEvent](extensionpoint.md#launchevent) mail add-ins)
- [Page](page.md)
- [Script](script.md)

## Attributes

| Attribute | Required | Description                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Yes      | The name of a URL resource defined in the **Resources** section of the manifest. Can be no more than 32 characters. |

## Child elements

None

## Example

```xml
<SourceLocation resid="pageURL"/>
```
