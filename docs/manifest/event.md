---
title: Event element in the manifest file
description: Defines an event handler in an add-in.
ms.date: 10/17/2022
ms.localizationpriority: medium
---

# Event element

Defines an event handler in an add-in. For information about support and usage, see [On-send feature for Outlook add-ins](/office/dev/add-ins/outlook/outlook-on-send-addins).

> [!NOTE]
> [Smart Alerts](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events), which is a newer version of the on-send feature, uses the [LaunchEvents element](launchevents.md) to configure an add-in for event-based activation. To learn more about the key differences between Smart Alerts and the on-send feature, see [Differences between Smart Alerts and the on-send feature](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events#differences-between-smart-alerts-and-the-on-send-feature). We invite you to [try out Smart Alerts by completing the walkthrough](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough).

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----:|:-----|
|  [Type](#type-attribute)  |  Yes  | Specifies the event to handle. |
|  [FunctionExecution](#functionexecution-attribute)  |  Yes  | Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported. |
|  [FunctionName](#functionname-attribute)  |  Yes  | Specifies the function name for the event handler. |

### Type attribute

Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.

|  Event type  |  Description  |
|:-----|:-----|
|  `ItemSend`  |  The event handler will be invoked when the user sends a message or meeting invitation.  |

### FunctionExecution attribute

Required. MUST be set to `synchronous`.

### FunctionName attribute

Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
