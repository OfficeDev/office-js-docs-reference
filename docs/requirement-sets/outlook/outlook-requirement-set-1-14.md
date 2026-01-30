---
title: Outlook add-in API requirement set 1.14
description: Lists the APIs introduced in Mailbox requirement set 1.14 for Outlook add-ins.
ms.date: 01/30/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.14

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.14

Mailbox requirement set 1.14 includes all of the features of [requirement set 1.13](outlook-requirement-set-1-13.md). It added the following features.

- Added the [integrated spam-reporting](/office/dev/add-ins/outlook/spam-reporting) feature.
- Added a method to get the Base64 encoding of a message.
- Added a method to get the URL of the JavaScript runtime of an add-in.
- Added support to customize the button text and configure a task pane for the **Don't Send** option of a [Smart Alerts](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events) dialog.
- Added support to override the send mode option of a [Smart Alerts](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events) add-in at runtime.
- Added members to get additional properties of a message in compose mode.
- Added a method to close a current message being composed with the option to discard unsaved changes.
- Added additional mail item properties for the [item multi-select](/office/dev/add-ins/outlook/item-multi-select) feature.
- Added support to identify the current Office theme of an Outlook client.
- Added support to get and set the [sensitivity level](https://support.microsoft.com/office/4a76d05b-6c29-4a0d-9096-71784a6b12c1) of an appointment.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.14. To view API reference documentation for all APIs supported by Mailbox requirement set 1.14 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.14&preserve-view=true).

[!INCLUDE [outlook-1_14](../../includes/outlook-1_14.md)]

## Events

The following table lists the [events](/javascript/api/office/office.eventtype?view=outlook-js-1.14&preserve-view=true) introduced in requirement set 1.14. For a list of all supported events that can be handled using the `addHandlerAsync` and `removeHandlerAsync` methods, see [Office.EventType](/javascript/api/office/office.eventtype?view=outlook-js-1.14&preserve-view=true).

| Event | Description | Object |
| --- | --- | --- |
| [OfficeThemeChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.14&preserve-view=true#fields) | The Office theme has changed in Outlook. | Mailbox |
| [SpamReporting](/javascript/api/office/office.eventtype?view=outlook-js-1.14&preserve-view=true#fields) | An unsolicited message is reported. | Item |

> [!NOTE]
> For events supported in an event-based activation add-in, see [Activate add-ins with events](/office/dev/add-ins/develop/event-based-activation#outlook-events).

## Manifest updates

The following table lists manifest updates introduced in Mailbox requirement set 1.14. To learn more about the types of Office Add-in manifests, see [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests).

| Feature | Unified manifest for Microsoft 365 | Add-in only manifest | Description |
| --- | --- | --- | --- |
| Spam reporting | [`"extensions.ribbons.contexts"`](/microsoft-365/extensibility/schema/extension-ribbons-array#contexts) array that contains `"spamReportingOverride"` | [ReportPhishingCommandSurface](/javascript/api/manifest/extensionpoint?view=outlook-js-1.14&preserve-view=true#reportphishingcommandsurface) | Activates a spam-reporting add-in in a prominent section of the Outlook ribbon. |
| Spam reporting | [`"extensions.ribbons.spamPreProcessingDialog"`](/microsoft-365/extensibility/schema/extension-ribbons-array#spampreprocessingdialog) | [ReportPhishingCustomization](/javascript/api/manifest/reportphishingcustomization?view=outlook-js-1.14&preserve-view=true) | Configures the ribbon button and preprocessing dialog of a spam-reporting add-in. |

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
