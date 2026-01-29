---
title: Outlook add-in API requirement set 1.7
description: Lists the APIs introduced in Mailbox requirement set 1.7 for Outlook add-ins.
ms.date: 01/30/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.7

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.7

Mailbox requirement set 1.7 includes all of the features of [requirement set 1.6](outlook-requirement-set-1.6.md). It added the following features.

- Added a property to get or set the [recurrence pattern](/office/dev/add-ins/outlook/get-and-set-recurrence) of an appointment and get the recurrence pattern of messages that are meeting requests.
- Added a property to get the ID of the series that an appointment instance belongs to.
- Added methods to manage the start and end dates and times of a recurring appointment series.
- Added a property to get the organizer value of an appointment in compose mode.
- Added a property to get the sender (From value) of a message in compose mode.
- Added methods to add and remove event handlers for supported events on items.
- Added events that occur when the recurrence pattern, recipients, or appointment time is changed.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.7. To view API reference documentation for all APIs supported by Mailbox requirement set 1.7 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.7&preserve-view=true).

[!INCLUDE [outlook-1_7](../../includes/outlook-1_7.md)]

## Events

The following table lists the [events](/javascript/api/office/office.eventtype?view=outlook-js-1.7&preserve-view=true) introduced in requirement set 1.7. For a list of all supported events that can be handled using the `addHandlerAsync` and `removeHandlerAsync` methods, see [Office.EventType](/javascript/api/office/office.eventtype?view=outlook-js-1.7&preserve-view=true).

| Event | Description | Object |
| --- | --- | --- |
| [AppointmentTimeChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.7&preserve-view=true#fields) | The date or time of the selected appointment or series was changed. | Item |
| [RecipientsChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.7&preserve-view=true#fields) | The recipient list of the selected item or appointment location was changed. | Item |
| [RecurrenceChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.7&preserve-view=true#fields) | The recurrence pattern of the selected meeting series was changed. | Item |

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
