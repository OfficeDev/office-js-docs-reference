---
title: ExtensionPoint element in the manifest file
description: Defines where an add-in exposes functionality in the Office UI.
ms.date: 07/11/2025
ms.localizationpriority: medium
---

# ExtensionPoint element

 Defines where an add-in exposes functionality in the Office UI. The **\<ExtensionPoint\>** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).

**Add-in type**: Document, Mail, Presentation, Task pane, Workbook

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----:|:-----|
|  **xsi:type**  |  Yes  | The type of extension point being defined. Possible values depend on the Office host application defined in the grandparent **\<Host\>** element value.|

## Extension points for Excel, Outlook, PowerPoint, and Word

- [LaunchEvent](#launchevent) - Activates tasks based on application events, such as opening.

### LaunchEvent

This extension point enables an add-in to activate based on supported events in both the desktop and mobile form factors. To learn more about event-based activation and for the full list of supported events, see [Activate add-ins with events](/office/dev/add-ins/develop/event-based-activation).

> [!IMPORTANT]
> Registering [Mailbox](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.md#events) and [Item](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  List of [LaunchEvent](launchevent.md) for event-based activation.  |
| [SourceLocation](customfunctionssourcelocation.md) |  The location of the source JavaScript file.  |

#### Example

```xml
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## Extension points for Excel, OneNote, PowerPoint, and Word add-in commands

There are three types of extension points available in some or all of these hosts.

- [PrimaryCommandSurface](#primarycommandsurface) (Valid for Word, Excel, PowerPoint, and OneNote) - The ribbon in Office.
- [ContextMenu](#contextmenu) (Valid for Word, Excel, PowerPoint, and OneNote) - The shortcut menu that appears when you right-click (or select and hold) in the Office UI.
- [CustomFunctions](#customfunctions) (Valid only for Excel) - A custom function written in JavaScript for Excel.

See the following subsections for the child elements and examples of these types of extension points.

### PrimaryCommandSurface

The primary command surface in Word, Excel, PowerPoint, and OneNote is the ribbon.

#### Child elements

|Element|Description|
|:-----|:-----|
|[CustomTab](customtab.md)|Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **\<CustomTab\>** element, you can't use the **\<OfficeTab\>** element. The **id** attribute is required. There can be no more than one **\<CustomTab\>** child element.|
|[OfficeTab](officetab.md)|Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**). If you use the **\<OfficeTab\>** element, you can't use the **\<CustomTab\>** element.|

> [!IMPORTANT]
> There can be no more than one **\<ExtensionPoint\>** element in the add-in that has a child **\<CustomTab\>** element; and that one **\<ExtensionPoint\>** element can have only one **\<CustomTab\>**, so there is only one **\<CustomTab\>** element across all **\<ExtensionPoint\>** elements.

#### Example

The following example shows how to use the **\<ExtensionPoint\>** element with **PrimaryCommandSurface**. It adds a custom tab to the ribbon.

> [!IMPORTANT]
> For elements that contain an ID attribute, make sure you provide a unique ID.

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.MyTab1">
    <Label resid="residLabel4" />
    <Group id="Contoso.Group1">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Control xsi:type="Button" id="Contoso.Button1">
          <!-- information about the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
  </CustomTab>
</ExtensionPoint>
```

### ContextMenu

A context menu is a shortcut menu that appears when you right-click (or select and hold) in the Office UI.

#### Child elements

|Element|Description|
|:-----|:-----|
|[OfficeMenu](officemenu.md)|Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to one of the following strings.<ul><li>**ContextMenuText** if the context menu should open when a user right-clicks (or selects and holds) on the selected text.</li><li>**ContextMenuCell** if the context menu should open when a user right-clicks (or selects and holds) on a cell in an Excel spreadsheet.</li></ul>|

#### Example

The following customizes the context menu opened on the selected text in a supported Office application. The context menu control used is of type **Button**.

```xml
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuText"> <!-- OR, for Excel only: <OfficeMenu id="ContextMenuCell"> -->
    <Control xsi:type="Button" id="ContextMenuButton">
      <Label resid="TaskpaneButton.Label"/>
      <Supertip>
        <Title resid="TaskpaneButton.Label" />
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

### CustomFunctions

A custom function written in JavaScript or TypeScript for Excel.

#### Child elements

|Element|Description|
|:-----|:-----|
|[Script](script.md)|Required. Links to the JavaScript file with the custom function's definition and registration code.|
|[Page](page.md)|Required. Links to the HTML page for your custom functions.|
|[MetaData](metadata.md)|Required. Defines the metadata settings used by a custom function in Excel.|
|[Namespace](namespace.md)|Optional. Defines the namespace used by a custom function in Excel.|

#### Example

```xml
<ExtensionPoint xsi:type="CustomFunctions">
  <Script>
    <SourceLocation resid="Functions.Script.Url"/>
  </Script>
  <Page>
    <SourceLocation resid="Shared.Url"/>
  </Page>
  <Metadata>
    <SourceLocation resid="Functions.Metadata.Url"/>
  </Metadata>
  <Namespace resid="Functions.Namespace"/>
</ExtensionPoint>
```

## Extension points for Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (can only be used in the [DesktopFormFactor](desktopformfactor.md))
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface)
- [MobileLogEventAppointmentAttendee](#mobilelogeventappointmentattendee)
- [Events](#events)
- [DetectedEntity](#detectedentity)
- [ReportPhishingCommandSurface](#reportphishingcommandsurface)

### MessageReadCommandSurface

This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

#### OfficeTab example

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab example

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="Contoso.TabCustom2">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### MessageComposeCommandSurface

This extension point puts buttons on the ribbon for add-ins using mail compose form.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

#### OfficeTab example

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab example

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="Contoso.TabCustom3">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentOrganizerCommandSurface

This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

#### OfficeTab example

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab example

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="Contoso.TabCustom4">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentAttendeeCommandSurface

This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

#### OfficeTab example

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab example

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="Contoso.TabCustom5">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### Module

This extension point adds a module extension add-in to the Outlook navigation bar. It also adds buttons to a custom tab on the ribbon for the module extension. To learn how to create module extensions, see [Module extension Outlook add-ins](/office/dev/add-ins/outlook/extension-module-outlook-add-ins).

> [!IMPORTANT]
> Registering [Mailbox](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.md#events) and [Item](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
| [SourceLocation](customfunctionssourcelocation.md) | Specifies the location of the HTML file that sets up the main user interface of the add-in. |
| Label | Specifies the label of the module extension. Its **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **\<String\>** element in the [ShortStrings](shortstrings.md) element. |
| [CommandSurface](commandsurface.md) | Adds a group of add-in buttons to a custom tab on the ribbon. |

#### Example

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

### MobileMessageReadCommandSurface

This extension point puts buttons in the command surface for the mail read view in the mobile form factor.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [Group](group.md) |  Adds a group of buttons to the command surface.  |

**\<ExtensionPoint\>** elements of this type can only have one child element: a **\<Group\>** element.

**\<Control\>** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.

#### Example

```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="Contoso.mobileGroup1">
    <Label resid="residAppName"/>
    <Control xsi:type="MobileButton" id="Contoso.mobileButton1">
      <!-- Control definition -->
    </Control>
  </Group>
</ExtensionPoint>
```

### MobileOnlineMeetingCommandSurface

This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor. A meeting organizer can create an online meeting. An attendee can subsequently join the online meeting. To learn more about this scenario, see [Create an Outlook mobile add-in for an online-meeting provider](/office/dev/add-ins/outlook/online-meeting).

> [!NOTE]
> This extension point is only supported on Android and iOS with a Microsoft 365 subscription.
>
> Registering [Mailbox](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.md#events) and [Item](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [Control](control.md) |  Adds a button to the command surface.  |

**\<ExtensionPoint\>** elements of this type can only have one child element: a **\<Control\>** element.

The **\<Control\>** element contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.

The images specified in the **\<Icon\>** element should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).

#### Example

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="Contoso.onlineMeetingFunctionButton1">
    <Label resid="residUILessButton0Name" />
    <Icon>
      <bt:Image resid="UiLessIcon" size="25" scale="1" />
      <bt:Image resid="UiLessIcon" size="25" scale="2" />
      <bt:Image resid="UiLessIcon" size="25" scale="3" />
      <bt:Image resid="UiLessIcon" size="32" scale="1" />
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
      <bt:Image resid="UiLessIcon" size="32" scale="3" />
      <bt:Image resid="UiLessIcon" size="48" scale="1" />
      <bt:Image resid="UiLessIcon" size="48" scale="2" />
      <bt:Image resid="UiLessIcon" size="48" scale="3" />
    </Icon>
    <Action xsi:type="ExecuteFunction">
      <FunctionName>insertContosoMeeting</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### MobileLogEventAppointmentAttendee

This extension point puts a **Log** action button contextually in the command surface for an appointment in the mobile form factor. Appointment attendees who have the add-in installed can save their appointment notes to an external app in one click. This extension point supports functionality for task pane and function commands. To learn more about this scenario, see [Log appointment notes to an external application in Outlook mobile add-ins](/office/dev/add-ins/outlook/mobile-log-appointments).

> [!NOTE]
> This extension point is only supported on Android and iOS with a Microsoft 365 subscription.
>
> Registering [Mailbox](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.md#events) and [Item](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.

#### Child elements

|  Element |  Description  |
|:-----|:-----|
|  [Control](control.md) |  Adds a button to the command surface.  |

**\<ExtensionPoint\>** elements of this type can only have one child element: a **\<Control\>** element.

The **\<Control\>** element contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.

The images specified in the **\<Icon\>** element should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).

#### Example

```xml
<ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
  <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
    <Label resid="LogButtonLabel" />
    <Icon>
      <bt:Image resid="Icon.16x16" size="25" scale="1" />
      <bt:Image resid="Icon.16x16" size="25" scale="2" />
      <bt:Image resid="Icon.16x16" size="25" scale="3" />
      <bt:Image resid="Icon.32x32" size="32" scale="1" />
      <bt:Image resid="Icon.32x32" size="32" scale="2" />
      <bt:Image resid="Icon.32x32" size="32" scale="3" />
      <bt:Image resid="Icon.80x80" size="48" scale="1" />
      <bt:Image resid="Icon.80x80" size="48" scale="2" />
      <bt:Image resid="Icon.80x80" size="48" scale="3" />
    </Icon>
    <Action xsi:type="ExecuteFunction">
      <FunctionName>logToCRM</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### Events

This extension point adds an event handler for a specified event. For more information about using this extension point, see [On-send feature for Outlook add-ins](/office/dev/add-ins/outlook/outlook-on-send-addins).

> [!IMPORTANT]
> Registering [Mailbox](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.md#events) and [Item](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.

> [!NOTE]
> [Smart Alerts](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events), which is a newer version of the on-send feature, uses the [LaunchEvent extension point](#launchevent) to enable event activation in an add-in. To learn more about the key differences between Smart Alerts and the on-send feature, see [Differences between Smart Alerts and the on-send feature](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events#differences-between-smart-alerts-and-the-on-send-feature). We invite you to [try out Smart Alerts by completing the walkthrough](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough).

| Element | Description  |
|:-----|:-----|
|  [Event](event.md) |  Specifies the event and event handler function.  |

#### ItemSend event example

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### DetectedEntity

This extension point adds a contextual add-in activation on a specified entity type. For more information about using this extension point, see [Contextual Outlook add-ins](/office/dev/add-ins/outlook/contextual-outlook-add-ins).

[!INCLUDE [outlook-contextual-add-ins-retirement](../includes/outlook-contextual-add-ins-retirement.md)]

The containing [VersionOverrides](versionoverrides.md) element must have an **xsi:type** attribute value of `VersionOverridesV1_1`.

> [!NOTE]
>
> - This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).
> - Registering [Mailbox](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.md#events) and [Item](../requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item.md#events) events isn't available with this extension point.

|  Element |  Description  |
|:-----|:-----|
|  [Label](#label) |  Specifies the label for the add-in in the contextual window.  |
|  [SourceLocation](customfunctionssourcelocation.md) |  Specifies the URL for the contextual window.  |
|  [Rule](rule.md) |  Specifies the rule or rules that determine when an add-in activates.  |

#### Label

Required. The label of the group. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **\<String\>** element in the **\<ShortStrings\>** element in the [Resources](resources.md) element.

#### Highlight requirements

The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the **Highlight** attribute of the **\<Rule\>** element for the `ItemHasRegularExpressionMatch` rule type.

However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.

- If using a single rule, the **Highlight** attribute MUST be set to `all`.
- If using a `RuleCollection` rule type with `Mode="And"` to combine multiple rules, at least one of the rules MUST have the **Highlight** attribute set to `all`.
- If using a `RuleCollection` rule type with `Mode="Or"` to combine multiple rules, all of the rules MUST have the **Highlight** attribute set to `all`.

#### DetectedEntity event example

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="Context.Label"/>
  <SourceLocation resid="DetectedEntity.URL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
  </Rule>
</ExtensionPoint>
```

### ReportPhishingCommandSurface

This extension point activates your spam-reporting add-in in the Outlook ribbon and prevents it from appearing at the end of the ribbon or in the overflow menu.

To learn more about how to implement the spam reporting feature in your add-in, see [Implement an integrated spam-reporting add-in](/office/dev/add-ins/outlook/spam-reporting).

#### Child elements

| Element | Description |
| ------- | ------- |
| [ReportPhishingCustomization element](reportphishingcustomization.md)| Configures the ribbon button and preprocessing dialog of a spam-reporting add-in. |

#### Example

```xml
<ExtensionPoint xsi:type="ReportPhishingCommandSurface">
  <ReportPhishingCustomization>
    <!-- Configures the ribbon button. -->
    <Control xsi:type="Button" id="ReportingButton">
      <Label resid="ReportingButton.Label"/>
      <Supertip>
        <Title resid="ReportingButton.Label"/>
        <Description resid="ReportingButton.Description"/>
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="Icon.16x16"/>
        <bt:Image size="32" resid="Icon.32x32"/>
        <bt:Image size="64" resid="Icon.64x64"/>
        <bt:Image size="80" resid="Icon.80x80"/>
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>onMessageReport</FunctionName>
      </Action>
    </Control>
    <!-- Configures the preprocessing dialog. -->
    <PreProcessingDialog>
      <Title resid="PreProcessingDialog.Label"/>
      <Description resid="PreProcessingDialog.Description"/>
      <ReportingOptions>
        <Title resid="OptionsTitle.Label"/>
        <Option resid="Option1.Label"/>
        <Option resid="Option2.Label"/>
        <Option resid="Option3.Label"/>
        <Option resid="Option4.Label"/>
      </ReportingOptions>
      <FreeTextLabel resid="FreeText.Label"/>
      <MoreInfo>
        <MoreInfoText resid="MoreInfo.Label"/>
        <MoreInfoUrl resid="MoreInfo.Url"/>
      </MoreInfo>
    </PreProcessingDialog>
    <SourceLocation resid="Commands.Url"/>
  </ReportPhishingCustomization>
</ExtensionPoint>
```
