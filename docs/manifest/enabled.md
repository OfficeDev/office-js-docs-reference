---
title: Enabled element in the manifest file
description: Learn how to specify that an Add-in Command is disabled when the add-in launches.
ms.date: 03/12/2022
ms.localizationpriority: medium
---

# Enabled element

Specifies whether a [Button control](control-button.md) or [Menu control](control-menu.md) is enabled when the add-in launches. The **\<Enabled\>** element is a child element of [Control](control.md). If it is omitted, the default is `true`.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [RibbonApi 1.0](../requirement-sets/common/ribbon-api-requirement-sets.md)

This element is only valid in Excel, PowerPoint, and Word; that is, when the `Name` attribute of the [Host](host.md) element is "Workbook", "Presentation", or "Document".

The parent control can also be programmatically enabled and disabled. For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).

## Example

```xml
<Enabled>false</Enabled>
```
