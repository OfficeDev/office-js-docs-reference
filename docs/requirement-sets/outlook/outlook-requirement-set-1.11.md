---
title: Outlook add-in API requirement set 1.11
description: Lists the APIs introduced in Mailbox requirement set 1.11 for Outlook add-ins.
ms.date: 01/27/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.11

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.11

Mailbox requirement set 1.11 includes all of the features of [requirement set 1.10](outlook-requirement-set-1.10.md). It added the following features.

- Added the following new events for [event-based activation](/office/dev/add-ins/develop/event-based-activation#supported-events).
  - `OnAppointmentAttachmentsChanged`
  - `OnAppointmentAttendeesChanged`
  - `OnAppointmentRecurrenceChanged`
  - `OnAppointmentTimeChanged`
  - `OnInfoBarDismissClicked`
  - `OnMessageAttachmentsChanged`
  - `OnMessageRecipientsChanged`
- Added SessionData APIs.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.11. To view API reference documentation for all APIs supported by Mailbox requirement set 1.11 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.11&preserve-view=true).

[!INCLUDE [outlook-1_11](../../includes/outlook-1_11.md)]

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
