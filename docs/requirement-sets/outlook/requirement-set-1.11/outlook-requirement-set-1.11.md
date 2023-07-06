---
title: Outlook add-in API requirement set 1.11
description: Requirement set 1.11 for Outlook add-in API.
ms.date: 09/09/2022
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.11

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.11?

Requirement set 1.11 includes all of the features of [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). It added the following features.

- Added new events for [event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events).
- Added SessionData APIs.

### Change log

- Added [Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties): Adds a new property to manage the session data of an item in Compose mode.
- Added [Office.SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-1.11&preserve-view=true): Adds a new object that represents the session data of a compose item.
- Added new events for [event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events): Adds support for the following events.

  - `OnAppointmentAttachmentsChanged`
  - `OnAppointmentAttendeesChanged`
  - `OnAppointmentRecurrenceChanged`
  - `OnAppointmentTimeChanged`
  - `OnInfoBarDismissClicked`
  - `OnMessageAttachmentsChanged`
  - `OnMessageRecipientsChanged`

- Added [Office.AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true): Adds an object that supports the `OnAppointmentTimeChanged` event.
- Added [Office.AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true): Adds an object that supports the `OnAppointmentAttachmentsChanged` and `OnMessageAttachmentsChanged` events.
- Added [Office.InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true): Adds an object that supports the `OnInfoBarDismissClicked` event.
- Added [Office.RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true): Adds an object that supports the `OnAppointmentAttendeesChanged` and `OnMessageRecipientsChanged` events.
- Added [Office.RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true): Adds an object that supports the `OnAppointmentRecurrenceChanged` event.

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
