---
title: Outlook add-in API requirement set 1.1
description: Lists the APIs introduced in Mailbox requirement set 1.1 for Outlook add-ins.
ms.date: 01/30/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.1

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in. Outlook JavaScript API 1.1 (Mailbox 1.1) is the first version of the API.

> [!NOTE]
> This documentation is for a [requirement set](outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.1

Mailbox requirement set 1.1 includes all of the [Common API requirement sets](../common/office-add-in-requirement-sets.md) supported in Outlook. It added the following features.

- Added methods to get or set the body of messages and appointments. For more information, see [Get or set the body of a message or appointment in Outlook](/office/dev/add-ins/outlook/insert-data-in-the-body).
- Added methods to add, get, and remove attachments from messages and appointments being composed. For more information, see [Manage an item's attachments in a compose form in Outlook](/office/dev/add-ins/outlook/add-and-remove-attachments-to-an-item-in-a-compose-form).
- Added methods to get or set the subject of messages and appointments being composed. For more information, see [Get or set the subject when composing an appointment or message in Outlook](/office/dev/add-ins/outlook/get-or-set-the-subject).
- Added methods to get or set the location of appointments being composed. For more information, see [Get or set the location when composing an appointment in Outlook](/office/dev/add-ins/outlook/get-or-set-the-location-of-an-appointment).
- Added methods to get or set the start time and end time of appointments being composed. For more information, see [Get or set the time when composing an appointment in Outlook](/office/dev/add-ins/outlook/get-or-set-the-time-of-an-appointment).
- Added methods to get or set the recipients (To, Cc, Bcc) of messages being composed. For more information, see [Get, set, or add recipients to an appointment or message in Outlook](/office/dev/add-ins/outlook/get-set-or-add-recipients).
- Added methods to display reply forms and new appointment forms.
- Added methods to display existing messages and appointments.
- Added the ability to make Exchange Web Services (EWS) requests directly from an add-in.

    [!INCLUDE [legacy-exchange-online](../../includes/legacy-exchange-online.md)]
- Added support for custom properties and roaming settings to store add-in data. For more information, see [Get and set add-in metadata for an Outlook add-in](/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in).
- Added methods to detect and extract entities (such as addresses, phone numbers, and URLs) from item bodies.

    [!INCLUDE [outlook-contextual-add-ins-retirement](../../includes/outlook-contextual-add-ins-retirement.md)]

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.1. To view API reference documentation for all APIs supported by Mailbox requirement set 1.1, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.1&preserve-view=true).

[!INCLUDE [outlook-1_1](../../includes/outlook-1_1.md)]

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
