---
title: Outlook add-in API requirement set 1.8
description: Lists the APIs introduced in Mailbox requirement set 1.8 for Outlook add-ins.
ms.date: 01/27/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.8

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.8

Mailbox requirement set 1.8 includes all of the features of [requirement set 1.7](outlook-requirement-set-1.7.md). It added the following features.

- Added methods to get the content of an attachment or get all attachments of an item in compose mode. For more information, see [Manage an item's attachments in a compose form in Outlook](/office/dev/add-ins/outlook/add-and-remove-attachments-to-an-item-in-a-compose-form).
- Added a method to add a file as an attachment using a Base64-encoded string.
- Added properties and methods to manage [categories](/office/dev/add-ins/outlook/categories) on items and on the mailbox's master category list.
- Added support for [delegate access](/office/dev/add-ins/outlook/delegate-access) scenarios, including a method to get shared properties and a manifest element to enable add-ins in shared folders.
- Added an object to manage the set of locations on an appointment. For more information, see [Get or set the location when composing an appointment in Outlook](/office/dev/add-ins/outlook/get-or-set-the-location-of-an-appointment).
- Added an object to get and set [custom internet headers](/office/dev/add-ins/outlook/internet-headers) on a message item in compose mode.
- Added a method to get all internet headers on a message item in read mode.
- Added a method to get initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in).
- Added a method to get the ID of a saved appointment or message item.
- Added optional `options` parameter to the [`event.completed`](/javascript/api/office/office.addincommands.event?view=outlook-js-1.8&preserve-view=true#office-office-addincommands-event-completed-member(1)) call to cancel execution of an event.
- Added events that occur when an attachment is added or removed and when the appointment location is changed.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.8. To view API reference documentation for all APIs supported by Mailbox requirement set 1.8 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.8&preserve-view=true).

[!INCLUDE [outlook-1_8](../../includes/outlook-1_8.md)]

## Events

The following table lists the [events](/javascript/api/office/office.eventtype?view=outlook-js-1.8&preserve-view=true) introduced in requirement set 1.8. For a list of all supported events that can be handled using the `addHandlerAsync` and `removeHandlerAsync` methods, see [Office.EventType](/javascript/api/office/office.eventtype?view=outlook-js-1.8&preserve-view=true).

| Event | Description | Object |
| --- | --- | --- |
| [AttachmentsChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.8&preserve-view=true#fields) | An attachment was added to or removed from an item. | Item |
| [EnhancedLocationsChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.8&preserve-view=true#fields) | The appointment location was changed. | Item |

## Manifest updates

The following table lists manifest updates introduced in Mailbox requirement set 1.8. To learn more about the types of Office Add-in manifests, see [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests).

| Feature | Unified manifest for Microsoft 365 | Add-in only manifest | Description |
| --- | --- | --- | --- |
| Shared folder support | [`"authorization.permissions.resourceSpecific.name"`](/microsoft-365/extensibility/schema/root-authorization-permissions-resource-specific#name) set to `"Mailbox.SharedFolder"` | [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) | Defines whether the add-in is available in shared folder scenarios. |

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
