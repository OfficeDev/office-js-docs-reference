---
title: OverriddenByRibbonApi element in the manifest file
description: Learn how to specify that a group, control, or menu item shouldn't appear when it is also part of a custom contextual tab.
ms.date: 05/25/2022
ms.localizationpriority: medium
---

# OverriddenByRibbonApi element

Specifies whether a [Group](group.md), [Button control](control-button.md), [Menu control](control-menu.md), or menu item will be hidden on application and platform combinations that support the API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1))) that installs custom contextual tabs on the ribbon.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Ribbon 1.2](../requirement-sets/common/add-in-commands-requirement-sets.md) (Required for Excel, PowerPoint, and Word.)

If this element is omitted, the default is `false`. If it's used, it must be the *first* child element of its parent element.

> [!NOTE]
> For a full understanding of this element, please read [Implement an alternate UI experience when custom contextual tabs are not supported](/office/dev/add-ins/design/contextual-tabs#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

The purpose of this element is to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs. The essential strategy is that you duplicate some or all of the groups and controls from your custom contextual tab onto a custom core tab (that is, *noncontextual* custom tab). Then, to ensure that these groups and controls appear when custom contextual tabs are *not* supported, but do not appear when custom contextual tabs *are* supported, you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the **\<Group\>**, **\<Control\>**, or menu **\<Item\>** elements. The effect of doing so is the following:

- If the add-in runs on an application and platform that support custom contextual tabs, then the duplicated groups and controls won't appear on the ribbon. Instead, the custom contextual tab will be installed when the add-in calls the `requestCreateControls` method.
- If the add-in runs on an application or platform that *doesn't* support custom contextual tabs, then the duplicated groups and controls will appear on the ribbon.

## Examples

### Overriding a group

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.TabCustom">
    <Group id="Contoso.CustomTab.group1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="Contoso.MyButton1">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel"/>
  </CustomTab>
</ExtensionPoint>
```

### Overriding a control

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.TabCustom">
    <Group id="Contoso.CustomTab.group2">
      <Control  xsi:type="Button" id="Contoso.MyButton2">
        <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
        <!-- Other child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel"/>
  </CustomTab>
</ExtensionPoint>
```

### Overriding a menu item

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.TabCustom">
    <Group id="Contoso.CustomTab.group3">
      <Control  xsi:type="Menu" id="Contoso.MyMenu">
        <!-- Other child elements omitted. -->
        <Items>
          <Item id="showGallery">
            <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
            <!-- Other child elements omitted. -->
          </Item>
        </Items>
      </Control>
    </Group>
    <Label resid="customTabLabel"/>
  </CustomTab>
</ExtensionPoint>
```
