---
title: LaunchEvents in the manifest file
description: The LaunchEvents element configures your add-in to activate based on supported events.
ms.date: 07/02/2025
ms.localizationpriority: medium
---

# LaunchEvents element

Configures your add-in to activate based on supported events. Child of the [ExtensionPoint](extensionpoint.md) element. For more information, see [Activate add-ins with events](/office/dev/add-ins/develop/event-based-activation).

**Add-in type:** Document, Mail, Presentation, Workbook

**Valid only in these VersionOverrides schemas**:

- Mail 1.1
- Task pane 1.0

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

## Syntax

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## Contained in

- [ExtensionPoint](extensionpoint.md) (**\<LaunchEvent\>** mail add-in)

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----:|:-----|
| [LaunchEvent](launchevent.md) | Yes |  Map supported event to its function in the JavaScript file for add-in activation. |

## See also

- [LaunchEvent](launchevent.md)
- [Activate add-ins with events](/office/dev/add-ins/develop/event-based-activation)
- [Automatically run an add-in when the document opens](/office/dev/add-ins/develop/launch-add-in-on-open)
- [Use Smart Alerts and the OnMessageSend event in your Outlook add-in](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough)
- [Automatically update your signature when switching between Exchange accounts](/office/dev/add-ins/outlook/onmessagefromchanged-onappointmentfromchanged-events)
