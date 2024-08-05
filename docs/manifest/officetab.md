---
title: OfficeTab element in the manifest file
description: The OfficeTab element defines the ribbon tab where your add-in command appears.
ms.date: 07/31/2024
ms.localizationpriority: medium
---

# OfficeTab element

Defines the built-in Office ribbon tab on which your add-in command appears. If you want the add-in command to appear on a custom tab of your own, use the [CustomTab](customtab.md) element.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md) when the parent **\<VersionOverrides\>** is type Taskpane 1.0.
- [Mailbox 1.3](../requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3.md) when the parent **\<VersionOverrides\>** is type Mail 1.0.
- [Mailbox 1.5](../requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5.md) when the parent **\<VersionOverrides\>** is type Mail 1.1.

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----:|:-----|
|  [id](#id)  |  Yes  | The ID of a built-in Office tab. |

### id

This attribute specifies the ID of the built-in Office tab. A list of valid `id` values is at [Find the IDs of built-in Office ribbon tabs](/office/dev/add-ins/develop/built-in-ui-ids).

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----:|:-----|
|  Group      | Yes |  Defines a group of controls. |

## Group

A group of controls in a tab. A group can have up to six controls. You can add only one group per add-in to the default tab. See [Group element](group.md).

## OfficeTab example

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="Contoso.msgreadTabMessage.group1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
