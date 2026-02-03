---
title: Outlook add-in API requirement set 1.9
description: Lists the APIs introduced in Mailbox requirement set 1.9 for Outlook add-ins.
ms.date: 02/03/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.9

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.9

Mailbox requirement set 1.9 includes all of the features of [requirement set 1.8](outlook-requirement-set-1-8.md). It added the following features.

- Added support to [append content to a mail item on send](/office/dev/add-ins/outlook/append-on-send).
- Added a method to get all [custom properties](/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in?tabs=custom-properties) of a mail item in a mailbox.
- Added methods to display a new message or appointment form.
- Added methods to display an existing message or appointment.
- Added support for [Dialog.messageChild](/office/dev/add-ins/develop/dialog-api-in-office-add-ins#pass-information-to-the-dialog-box) to deliver a message from the host page, such as a task pane or a function command, to a dialog that was opened from the page.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.9. To view API reference documentation for all APIs supported by Mailbox requirement set 1.9 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.9&preserve-view=true).

[!INCLUDE [outlook-1_9](../../includes/outlook-1_9.md)]

## Manifest updates

The following table lists manifest updates introduced in Mailbox requirement set 1.9. To learn more about the types of Office Add-in manifests, see [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests).

| Feature | Unified manifest for Microsoft 365 | Add-in only manifest | Description |
| --- | --- | --- | --- |
| Append-on-send | [`"authorization.permissions.resourceSpecific.name"`](/microsoft-365/extensibility/schema/root-authorization-permissions-resource-specific#name) set to `"Mailbox.AppendOnSend.User"` | [ExtendedPermissions element](/javascript/api/manifest/extendedpermissions) | Enables support for the append-on-send feature. |

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
