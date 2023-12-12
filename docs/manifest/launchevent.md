---
title: LaunchEvent in the manifest file
description: The LaunchEvent element configures your add-in to activate based on supported events.
ms.date: 12/12/2023
ms.localizationpriority: medium
---

# LaunchEvent element

Configures your add-in to activate based on supported events. Child of the [LaunchEvents](launchevents.md) element. For more information, see [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch).

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the manifest](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

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

- [LaunchEvents](launchevents.md)

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----:|:-----|
|  **\<Type\>**  |  Yes  | Specifies a supported event type. For the set of supported types, see [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events). |
|  **\<FunctionName\>**  |  Yes  | Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute. |
|  **SendMode** |  No  | Used by `OnMessageSend` and `OnAppointmentSend` events. Specifies the options available to the user if your add-in stops an item from being sent or if the add-in is unavailable. If the **SendMode** property isn't included, the `SoftBlock` option is set by default. For a list of available send mode options, see [Available send mode options](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events#available-send-mode-options). |

## See also

- [LaunchEvents](launchevents.md)
- [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events)
- [Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events)
- [Automatically check for an attachment before a message is sent](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough)
- [Implement event-based activation in Outlook mobile add-ins](/office/dev/add-ins/outlook/mobile-event-based)
