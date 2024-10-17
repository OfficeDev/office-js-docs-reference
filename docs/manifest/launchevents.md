---
title: LaunchEvents in the manifest file
description: The LaunchEvents element configures your add-in to activate based on supported events.
ms.date: 01/18/2024
ms.localizationpriority: medium
---

# LaunchEvents element

Configures your add-in to activate based on supported events. Child of the [ExtensionPoint](extensionpoint.md) element. For more information, see [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch).

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

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
- [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events)
- [Use Smart Alerts and the OnMessageSend event in your Outlook add-in](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough)
- [Automatically update your signature when switching between Exchange accounts](/office/dev/add-ins/outlook/onmessagefromchanged-onappointmentfromchanged-events)
