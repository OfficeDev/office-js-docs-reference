---
title: Outlook add-in API requirement set 1.12
description: Requirement set 1.12 for Outlook add-in API.
ms.date: 08/28/2023
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.12

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.12?

Requirement set 1.12 includes all of the features of [requirement set 1.11](../requirement-set-1.11/outlook-requirement-set-1.11.md). It added the following features.

- Added new events for [event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events).
- Added [send mode options](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough#available-send-mode-options) for add-ins that use the `OnMessageSend` or `OnAppointmentSend` event.
- Added support to display an error message to the user in [event-based activation](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough) add-ins.

### Change log

- Added new events for [event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events): Adds support for the following events.

  - `OnMessageSend`
  - `OnAppointmentSend`
  - `OnMessageCompose`
  - `OnAppointmentOrganizer`

- Modified the [LaunchEvent manifest element](/javascript/api/manifest/launchevent): Adds the `SendMode` attribute used by the `OnMessageSend` and `OnAppointmentSend` events. This attribute specifies options available to the user if an add-in stops an item from being sent or if the add-in is unavailable.
- Created [Office.SmartAlertsEventCompletedOptions](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.12&preserve-view=true): Adds the `allowEvent` and `errorMessage` properties. The `allowEvent` property indicates if the handled event should continue execution or be canceled. If `allowEvent` is set to `false` when the add-in's condition isn't met, the `errorMessage` property can be used to display a message to the user.

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
