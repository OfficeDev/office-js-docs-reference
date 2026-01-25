---
title: Outlook add-in API requirement set 1.12
description: Lists the APIs introduced in Mailbox requirement set 1.12 for Outlook add-ins.
ms.date: 01/27/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.12

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.12

Mailbox requirement set 1.12 includes all of the features of [requirement set 1.11](outlook-requirement-set-1.11.md). It added the following features.

- Added the following new events for [event-based activation](/office/dev/add-ins/develop/event-based-activation#supported-events).
  - `OnMessageSend`
  - `OnAppointmentSend`
  - `OnMessageCompose`
  - `OnAppointmentOrganizer`
- Introduced [Smart Alerts add-ins](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events) to handle the `OnMessageSend` and `OnAppointmentSend` events.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.12. To view API reference documentation for all APIs supported by Mailbox requirement set 1.12 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.12&preserve-view=true).

[!INCLUDE [outlook-1_12](../../../includes/outlook-1_12.md)]

## Manifest updates

The following table lists manifest updates introduced in Mailbox requirement set 1.12. To learn more about the types of Office Add-in manifests, see [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests).

| Feature | Unified manifest for Microsoft 365 | Add-in only manifest | Description |
| --- | --- | --- | --- |
| Smart Alerts | [`"extensions.autoRunEvents.events.options.sendMode"`](/microsoft-365/extensibility/schema/extension-auto-run-events-array-events-options#sendmode) | [LaunchEvent SendMode attribute](/javascript/api/manifest/launchevent) | Specifies options available to the user if an add-in stops an item from being sent or if the add-in is unavailable. Used by the `OnMessageSend` and `OnAppointmentSend` events. |

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
