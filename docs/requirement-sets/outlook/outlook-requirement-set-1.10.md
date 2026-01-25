---
title: Outlook add-in API requirement set 1.10
description: Lists the APIs introduced in Mailbox requirement set 1.10 for Outlook add-ins.
ms.date: 01/27/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.10

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.10

Mailbox requirement set 1.10 includes all of the features of [requirement set 1.9](outlook-requirement-set-1.9.md). It added the following features.

- Added support for [event-based activation](/office/dev/add-ins/develop/event-based-activation).
- Added support for mail signature features.
- Added support for the [OfficeRuntime.Storage](/javascript/api/office-runtime/officeruntime.storage?view=outlook-js-1.10&preserve-view=true) object with the event-based activation feature.
- Added ability to include a [custom action on a notification message](/office/dev/add-ins/outlook/notifications#insightmessage).

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.10. To view API reference documentation for all APIs supported by Mailbox requirement set 1.10 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.10&preserve-view=true).

[!INCLUDE [outlook-1_10](../../../includes/outlook-1_10.md)]

## Manifest updates

The following table lists manifest updates introduced in Mailbox requirement set 1.10. To learn more about the types of Office Add-in manifests, see [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests).

| Feature | Unified manifest for Microsoft 365 | Add-in only manifest | Description |
| --- | --- | --- | --- |
| Event-based activation | [`"extensions.autoRunEvents"`](/microsoft-365/extensibility/schema/extension-auto-run-events-array) | [LaunchEvent extension point](/javascript/api/manifest/extensionpoint#launchevent) | Configures event-based activation functionality. |

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
