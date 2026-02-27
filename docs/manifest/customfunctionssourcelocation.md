---
title: SourceLocation element (version overrides) in the manifest file
description: Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel, or needed by the ReportPhishingCustomization element, DetectedEntity extension point, or LaunchEvent extension point in Outlook.
ms.date: 02/27/2026
ms.localizationpriority: medium
---

# SourceLocation element (version overrides)

Defines the location of a resource needed by the following elements.

- `LaunchEvent` extension point in Excel, Outlook, PowerPoint, and Word.
- **\<Script\>** or **\<Page\>** elements used by custom functions in Excel.
- **\<ReportPhishingCustomization\>** element, `DetectedEntity` extension point, or `Module` extension point in Outlook.

> [!IMPORTANT]
> This article only refers to the **\<SourceLocation\>** that's a child of the following:
>
> - **\<Page\>**, **\<Script\>**, or **\<ReportPhishingCustomization\>** elements
> - `DetectedEntity`, `LaunchEvent`, or `Module` extension points
>
> For information about the **\<SourceLocation\>** element of the base manifest, see [SourceLocation](sourcelocation.md).

**Add-in type:** Custom function, Mail

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)
- [Mailbox 1.5](../requirement-sets/outlook/outlook-requirement-set-1-5.md)
- [Mailbox 1.6](../requirement-sets/outlook/outlook-requirement-set-1-6.md)
- [Mailbox 1.10](../requirement-sets/outlook/outlook-requirement-set-1-10.md)
- [Mailbox 1.14](../requirement-sets/outlook/outlook-requirement-set-1-14.md)

## Contained in

- [ExtensionPoint](extensionpoint.md) ([DetectedEntity](extensionpoint.md#detectedentity), [LaunchEvent](extensionpoint.md#launchevent), and [Module](extensionpoint.md#module) add-ins)
- [Page](page.md)
- [ReportPhishingCustomization](reportphishingcustomization.md) (Mail add-ins)
- [Script](script.md)

## Attributes

| Attribute | Required | Description |
|:----------|:--------:|:------------|
| **resid** | Yes | The name of a URL resource defined in the **\<Resources\>** section of the manifest. Can be no more than 32 characters.<br><br>**Important**: The `resid` value of the **\<SourceLocation\>** element in a `LaunchEvent` extension point or **\<ReportPhishingCustomization\>** element must match the `resid` value of the **\<Runtime\>** element that represents the [browser runtime](/office/dev/add-ins/testing/runtimes#types-of-runtimes). For example, if your runtime is defined as `<Runtime resid="WebViewRuntime.Url">`, specify `<SourceLocation resid="WebViewRuntime.Url"/>`. |

## Child elements

None.

## Example

```xml
<SourceLocation resid="pageURL"/>
```
