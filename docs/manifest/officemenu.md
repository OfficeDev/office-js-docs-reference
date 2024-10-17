---
title: OfficeMenu element in the manifest file
description: The OfficeMenu element defines a collection of controls to be added to the Office context menu.
ms.date: 08/13/2024
ms.localizationpriority: medium
---

# OfficeMenu element

Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md)

## Attributes

| Attribute            | Required | Description                          |
|:---------------------|:--------:|:-------------------------------------|
| [id](#id) | Yes      | The type of OfficeMenu being defined.|

### id

Although its official data type is string, this attribute effectively functions as a type attribute, and there are only two possible values. The attribute specifies the type of built-in Office menu to add this Office Add-in to.

- `ContextMenuText` - Displays the item on the context menu when text is selected and the user opens that menu (e.g., right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.
- `ContextMenuCell` - Displays the item on the context menu when the user opens that menu (e.g., right-clicks) on a cell on the spreadsheet. Applies to Excel.


## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----:|:-----|
|  [Control of type Button](control-button.md)    | Yes |  A single **Button** control object.  |

> [!NOTE]
> There can be only one child control and it must be type **Button**.

## Example

```xml
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuText">
    <Control xsi:type="Button" id="ContextMenuButton">
      <Label resid="TaskpaneButton.Label"/>
      <Supertip>
        <!-- ToolTip title. resid must point to a ShortString resource. -->
        <Title resid="TaskpaneButton.Label" />
        <!-- ToolTip description. resid must point to a LongString resource. -->
        <Description resid="TaskpaneButton.Tooltip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="tpicon_16x16" />
        <bt:Image size="32" resid="tpicon_32x32" />
        <bt:Image size="80" resid="tpicon_80x80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>action</FunctionName>
      </Action>
    </Control>
  </OfficeMenu>
</ExtensionPoint>
```
