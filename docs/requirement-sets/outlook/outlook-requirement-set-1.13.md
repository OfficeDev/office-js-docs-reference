---
title: Outlook add-in API requirement set 1.13
description: Lists the APIs introduced in Mailbox requirement set 1.13 for Outlook add-ins.
ms.date: 01/27/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.13

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.13

Mailbox requirement set 1.13 includes all of the features of [requirement set 1.12](outlook-requirement-set-1.12.md). It added the following features.

- Added support to activate an add-in without the Reading Pane enabled or a message selected.
- Added support to manage the delivery data and time of a message.
- Added new events for [event-based activation](/office/dev/add-ins/develop/event-based-activation#supported-events).
- Added the item multi-select feature.
- Added the prepend-on-send feature.
- Added the sensitivity label feature.
- Added support for shared mailbox scenarios.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.13. To view API reference documentation for all APIs supported by Mailbox requirement set 1.13 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.13&preserve-view=true).

[!INCLUDE [outlook-1_13](../../../includes/outlook-1_13.md)]

## Events

The following table lists the [events](/javascript/api/office/office.eventtype?view=outlook-js-1.13&preserve-view=true) introduced in requirement set 1.13. For a list of all supported events that can be handled using the `addHandlerAsync` and `removeHandlerAsync` methods, see [Office.EventType](/javascript/api/office/office.eventtype?view=outlook-js-1.13&preserve-view=true).

| Event | Description | Object |
| --- | --- | --- |
| [SelectedItemsChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.13&preserve-view=true#fields) | One or more messages are selected or deselected. | Mailbox |
| [SensitivityLabelChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.13&preserve-view=true#fields) | The sensitivity label of a message or appointment is changed. | Item |

> [!NOTE]
> For events supported in an event-based activation add-in, see [Activate add-ins with events](/office/dev/add-ins/develop/event-based-activation#outlook-events).

## Manifest updates

The following table lists manifest updates introduced in Mailbox requirement set 1.13. To learn more about the types of Office Add-in manifests, see [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests).

| Feature | Unified manifest for Microsoft 365 | Add-in only manifest | Description |
| --- | --- | --- | --- |
| No item context | [`"extensions.runtimes.actions.supportsNoItemContext"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item#supportsnoitemcontext) | [SupportsNoItemContext](/javascript/api/manifest/action?view=outlook-js-1.13&preserve-view=true#supportsnoitemcontext) | Allows task pane add-ins to activate without the Reading Pane enabled or a message selected. |
| Shared mailbox support | [`"authorization.permissions.resourceSpecific.name"`](/microsoft-365/extensibility/schema/root-authorization-permissions-resource-specific#name) set to `"Mailbox.SharedFolder"` | [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders?view=outlook-js-1.13&preserve-view=true) | Defines whether the add-in is available in shared mailbox scenarios. |

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
