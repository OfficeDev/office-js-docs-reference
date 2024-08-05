---
title: SourceLocation element (version overrides) in the manifest file
description: Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel, or needed by the ReportPhishingCustomization element, DetectedEntity extension point, or LaunchEvent extension point in Outlook.
ms.date: 05/20/2024
ms.localizationpriority: medium
---

# SourceLocation element (version overrides)

Defines the location of a resource needed by the **\<Script\>** or **\<Page\>** elements used by custom functions in Excel, or needed by the **\<ReportPhishingCustomization\>** element, **\<DetectedEntity\>** extension point, or **\<LaunchEvent\>** extension point in Outlook.

> [!IMPORTANT]
> This article only refers to the **\<SourceLocation\>** that is a child of the following:
>
> - **\<Page\>**, **\<Script\>**, or **\<ReportPhishingCustomization\>** elements
> - **\<DetectedEntity\>** or **\<LaunchEvent\>** extension points
>
> For information about the **\<SourceLocation\>** element of the base manifest, see [SourceLocation](sourcelocation.md).

**Add-in type:** Custom function, Mail

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)
- [Mailbox 1.6](../requirement-sets/outlook/requirement-set-1.6/outlook-requirement-set-1.6.md)
- [Mailbox 1.10](../requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10.md)
- [Mailbox 1.14](../requirement-sets/outlook/requirement-set-1.14/outlook-requirement-set-1.14.md)

## Contained in

- [ExtensionPoint](extensionpoint.md) ([Contextual](extensionpoint.md#detectedentity) and [LaunchEvent](extensionpoint.md#launchevent) mail add-ins)
- [Page](page.md)
- [ReportPhishingCustomization](reportphishingcustomization.md) (Mail add-ins)
- [Script](script.md)

## Attributes

| Attribute | Required | Description |
|:----------|:--------:|:------------|
| **resid** | Yes | The name of a URL resource defined in the **\<Resources\>** section of the manifest. Can be no more than 32 characters. |

## Child elements

None.

## Example

```xml
<SourceLocation resid="pageURL"/>
```
