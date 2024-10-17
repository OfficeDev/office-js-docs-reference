---
title: Action element in the manifest file
description: This element specifies the action to perform when the user selects a button or menu control.
ms.date: 02/29/2024
ms.localizationpriority: medium
---

# Action element

Specifies the action to perform when the user selects a [Button](control-button.md) or [Menu](control-menu.md) control.

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
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
|  [xsi:type](#xsitype)  |  Yes  | Action type to take|

### xsi:type

This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:

- `ExecuteFunction`
- `ShowTaskpane`

Once the user selects a button that kicks off the `ExecuteFunction` action, the add-in times out after 5 minutes if it hasn't completed by then.

> [!IMPORTANT]
> Outlook: Registering [Mailbox](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.md#events) and [Item](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item.md#events) events is not available when **xsi:type** is `ExecuteFunction`.

## Child elements

The valid child elements very depending on the value of the `xsi:type` parameter.

### xsi:type is ExecuteFunction

|  Element |  Description  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Specifies the name of the function to execute. |

#### FunctionName

Required element when **xsi:type** is `ExecuteFunction`. Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

### xsi:type is ShowTaskpane

|  Element |  Description  |
|:-----|:-----|
|  [SourceLocation](#sourcelocation) |    Specifies the source file location for this action. |
|  [TaskpaneId](#taskpaneid) | Specifies the ID of the task pane container. Not supported in Outlook add-ins.|
|  [Title](#title) | Specifies the custom title for the task pane. Not supported in Outlook add-ins.|
|  [SupportsPinning](#supportspinning) | Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection. Supported in Outlook only. |
|  [SupportsMultiselect](#supportsmultiselect) | Specifies that an Outlook add-in can activate on multiple selected messages. Supported in Outlook only. |
|  [SupportsNoItemContext](#supportsnoitemcontext) | Specifies that an Outlook add-in can activate without the Reading Pane enabled or a message selected. Supported in Outlook desktop clients only. |

#### SourceLocation

Required element when **xsi:type** is `ShowTaskpane`. Specifies the source file location for this action. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **\<Url\>** element in the **\<Urls\>** element in the [Resources](resources.md) element.

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

#### TaskpaneId

Optional element when  **xsi:type** is `ShowTaskpane`. Specifies the ID of the task pane container. When you have multiple `ShowTaskpane` actions, use a different **\<TaskpaneId\>** if you want an independent pane for each. Use the same **\<TaskpaneId\>** for  different actions that share the same pane. When users choose commands that share the same **\<TaskpaneId\>**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action `SourceLocation`.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md)

> [!NOTE]
> This element is not supported in Outlook.

The following example shows two actions that share the same **\<TaskpaneId\>**.

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

The following examples show two actions that use a different **\<TaskpaneId\>**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

#### Title

Optional element when  **xsi:type** is `ShowTaskpane`. Specifies the custom title for the task pane for this action.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md)

> [!NOTE]
> This child element is not supported in Outlook add-ins.

The following example shows an action that uses the **\<Title\>** element. Note that you don't assign the **\<Title\>** to a string directly. Instead, you assign it a resource ID (resid), that is defined in the **\<Resources\>** section of the manifest and can be no more than 32 characters.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="PG.Code.Url" />
    <Title resid="PG.CodeCommand.Title" />
</Action>

 ... Other markup omitted ...
<Resources>
    <bt:Images> ...
    </bt:Images>
    <bt:Urls>
        <bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
    </bt:ShortStrings>
 ... Other markup omitted ...
</Resources>
```

#### SupportsPinning

Optional element when **xsi:type** is `ShowTaskpane`. The containing [VersionOverrides](versionoverrides.md) elements must have an **xsi:type** attribute value of `VersionOverridesV1_1`. Include this element with a value of `true` to support task pane pinning. The user will be able to "pin" the task pane, causing it to stay open when changing the selection. For more information, see [Implement a pinnable task pane in Outlook](/office/dev/add-ins/outlook/pinnable-taskpane).

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.5](../requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5.md)

> [!IMPORTANT]
> Although the **SupportsPinning** element was introduced in [requirement set 1.5](../requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following:
>
> - Modern Outlook on the web
> - [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)
> - Outlook 2016 or later on Windows (Version 1612 (Build 7628.1000) or later)
> - Outlook on Mac (Version 16.13 (18050300) or later)

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```

#### SupportsMultiselect

Optional element in Outlook add-ins when **xsi:type** is `ShowTaskpane`. Include a value of `true` to allow an add-in to activate and perform specific operations on multiple selected messages. Because item multi-select only applies to messages, the [ExtensionPoint element's xsi:type attribute value](extensionpoint.md#extension-points-for-outlook) must be set to `MessageReadCommandSurface` or `MessageComposeCommandSurface`. To learn more about item multi-select, see [Activate your Outlook add-in on multiple messages](/office/dev/add-ins/outlook/item-multi-select).

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.13](../requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13.md)

```xml
<Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskpaneUrl" />
    <SupportsMultiSelect>true</SupportsMultiSelect>
</Action>
```

#### SupportsNoItemContext

Optional element in Outlook add-ins when **xsi:type** is `ShowTaskpane`. Include a value of `true` to allow an add-in to activate without the Reading Pane enabled or a message selected. If **\<SupportsNoItemContext\>** is set to `true`, the [ExtensionPoint element's xsi:type attribute value](extensionpoint.md#extension-points-for-outlook) must be set to `MessageReadCommandSurface`. To learn more, see [Activate your Outlook add-in without the Reading Pane enabled or a message selected](/office/dev/add-ins/outlook/contextless).

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.13](../requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13.md)

> [!NOTE]
> Although Outlook on the web and new Outlook on Windows support Mailbox requirement set 1.13, an add-in won't activate if the Reading Pane is hidden or a message isn't first selected. To learn more, see [Feature support in Outlook on the web and new Outlook on Windows](/office/dev/add-ins/outlook/contextless#feature-support-in-outlook-on-the-web-and-new-outlook-on-windows).

```xml
<Action xsi:type="ShowTaskpane">
    <SourceLocation resid="Taskpane.Url"/>
    <SupportsNoItemContext>true</SupportsNoItemContext>
</Action>
```
