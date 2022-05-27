---
title: LaunchEvent in the manifest file
description: The LaunchEvent element configures your add-in to activate based on supported events.
ms.date: 05/26/2022
ms.localizationpriority: medium
---

# LaunchEvent element

Configures your add-in to activate based on supported events. Child of the [`<LaunchEvents>`](launchevents.md) element. For more information, see [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch).

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
|:-----|:-----|:-----|
|  **Type**  |  Yes  | Specifies a supported event type. For the set of supported types, see [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events). |
|  **FunctionName**  |  Yes  | Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute. |
|  **SendMode** (preview) |  No  | Used by `OnMessageSend` and `OnAppointmentSend` events. Specifies the options available to the user if your add-in stops an item from being sent or if the add-in is unavailable. If the **SendMode** property isn't included, the `SoftBlock` option is set by default. For available options, refer to [Available SendMode options](#available-sendmode-options-preview). |

## Available SendMode options (preview)

When you include the `OnMessageSend` or `OnAppointmentSend` event in the manifest, you should also set the **SendMode** property. If the **SendMode** property isn't included, the `SoftBlock` option is set by default. The following are the available options. Based on the conditions your add-in is looking for, the user is alerted if your add-in finds an issue in the item being sent.

### `PromptUser`

If the item doesn't meet the add-in's conditions, the user can choose **Send Anyway** in the alert, or address the issue then try to send the item again. If the add-in is taking a long time to process the item, the user will be prompted with the option to stop running the add-in and choose **Send Anyway**. In the event the add-in is unavailable (for example, there's an error loading the add-in), the item will be sent.

![PromptUser dialog with the Send Anyway and Don't Send options.](/javascript/api/images/outlook-launchevent-promptUser.png)

Use the `PromptUser` option in your add-in if one of the following applies.

- The condition checked by the add-in isn't mandatory, but is nice to have in the message or appointment being sent.
- You'd like to recommend an action and allow the user to decide whether they want to apply it to the message or appointment being sent.

Some scenarios where the `PromptUser` option is applied include suggesting to tag the message or appointment as low or high importance and recommending to apply a color category to the item.

### `SoftBlock`

Default option if the **SendMode** property isn't included. The user is alerted that the item they're sending doesn't meet the add-in's conditions and they must address the issue before trying to send the item again. However, if the add-in is unavailable (for example, there's an error loading the add-in), the item will be sent.

![SoftBlock dialog with the Don't Send option.](/javascript/api/images/outlook-launchevent-soft-block.png)

Use the `SoftBlock` option in your add-in when you want a condition to be met before a message or appointment can be sent, but you don't want the user to be blocked from sending the item if the add-in is unavailable. Sample scenarios where the `SoftBlock` option is used include prompting the user to set a message or appointment's importance level and checking that the appropriate signature is applied before the item is sent.

### `Block`

The item isn't sent if any of the following situations occur.

- The item doesn't meet the add-in's conditions.
- The add-in is unable to connect to the server.
- There's an error loading the add-in.

![Block dialog with the Don't Send option.](/javascript/api/images/outlook-launchevent-block.png)

Use the `Block` option if the add-in's conditions are mandatory, even if the add-in is unavailable. For example, the `Block` option is ideal when users are required to apply a sensitivity label to a message or appointment before it can be sent.

For additional information on each **SendMode** option's behavior in particular scenarios, see [Smart Alerts feature behavior and scenarios](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough#smart-alerts-feature-behavior-and-scenarios).

## See also

- [LaunchEvents](launchevents.md)
- [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events)
- [Use Smart Alerts and the OnMessageSend event in your Outlook add-in](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough)
