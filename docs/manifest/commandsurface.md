---
title: CommandSurface element in the manifest file
description: Defines the custom tab and ribbon buttons of a module extension add-in in classic Outlook on Windows.
ms.date: 02/27/2025
ms.localizationpriority: medium
---

# CommandSurface element

Defines the custom tab and ribbon buttons of a module extension add-in in classic Outlook on Windows. For more information about module extensions, see [Module extension Outlook add-ins](/office/dev/add-ins/outlook/extension-module-outlook-add-ins).

**Add-in type**: Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.5](../requirement-sets/outlook/outlook-requirement-set-1-5.md)

## Contained in

- [ExtensionPoint](extensionpoint.md) ([Module](extensionpoint.md#module) mail add-ins)

## Attributes

None.

## Child elements

| Element | Required | Description |
|:-----|:-----:|:-----|
| [CustomTab](customtab.md) | Yes | Defines a custom tab on the ribbon for the module extension add-in. The custom tab hosts buttons that run add-in operations. |

## Example

```xml
<ExtensionPoint xsi:type="Module">
  <SourceLocation resid="residExtensionPointUrl"/>
  <Label resid="residExtensionPointLabel"/>
  <CommandSurface>
    <CustomTab id="idTab">
      <Group id="idGroup">
        <Label resid="residGroupLabel"/>
        <Control xsi:type="Button" id="group.changeToAssociate">
          <Label resid="residChangeToAssociateLabel"/>
          <Supertip>
            <Title resid="residChangeToAssociateLabel"/>
            <Description resid="residChangeToAssociateDesc"/>
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="residAssociateIcon16"/>
            <bt:Image size="32" resid="residAssociateIcon32"/>
            <bt:Image size="80" resid="residAssociateIcon80"/>
          </Icon>
          <Action xsi:type="ExecuteFunction">
            <FunctionName>changeToAssociateRate</FunctionName>
          </Action>
        </Control>
      </Group>
      <Label resid="residCustomTabLabel"/>
    </CustomTab>
  </CommandSurface>
</ExtensionPoint>
```
