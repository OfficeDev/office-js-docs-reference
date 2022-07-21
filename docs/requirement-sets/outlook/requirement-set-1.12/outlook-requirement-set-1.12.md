---
title: Outlook add-in API requirement set 1.12
description: Requirement set 1.12 for Outlook add-in API.
ms.date: 07/20/2022
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.12

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

## What's new in 1.12?

Requirement set 1.12 includes all of the features of [requirement set 1.11](../requirement-set-1.11/outlook-requirement-set-1.11.md). It added the following features.

- Added the `OnMessageSend` and `OnAppointmentSend` events for [event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events).
- Added the `errorMessage` value to the `options` parameter of [Event.completed](/javascript/api/office/office.addincommands.event).

### Change log

- Added new events for [event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events): Adds support for the `OnMessageSend` and `OnAppointmentSend` events.
- Modified [Office.AddinCommands.EventCompletedOptions](/javascript/api/office/office.addincommands.eventcompletedoptions?view=outlook-js-1.12&preserve-view=true): Adds the `errorMessage` property to display a message to the user if `allowEvent` is set to `false` when the add-in's event handler condition isn't met.

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
