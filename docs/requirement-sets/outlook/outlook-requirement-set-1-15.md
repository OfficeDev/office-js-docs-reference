---
title: Outlook add-in API requirement set 1.15
description: Lists the APIs introduced in Mailbox requirement set 1.15 for Outlook add-ins.
ms.date: 02/03/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.15

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

## What's new in 1.15

Mailbox requirement set 1.15 includes all of the features of [requirement set 1.14](outlook-requirement-set-1-14.md). It added the following features.

- Added support for radio buttons to format reporting options in the [integrated spam-reporting](/office/dev/add-ins/outlook/spam-reporting) dialog.
- Added support to include a ["Don't show this message again" checkbox](/office/dev/add-ins/outlook/spam-reporting#suppress-the-preprocessing-dialog) in a spam-reporting dialog that doesn't require input from a user.
- Added support to [open a task pane from a spam-reporting dialog](/office/dev/add-ins/outlook/spam-reporting#open-a-task-pane-after-reporting-a-message).
- Added support for Markdown to format error messages in a [Smart Alerts](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events) dialog.
- Added support to [run a function from the Smart Alerts dialog](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough#customize-the-text-and-functionality-of-a-button-in-the-dialog-optional).
- Added a method to programmatically send a mail item.
- Added a method to load a single message from multiple selected messages to get its properties or perform operations on it.
- Added an event that occurs when an add-in's task pane is opened from an [actionable message](/outlook/actionable-messages), [InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.15&preserve-view=true#fields) notification, Smart Alerts dialog, or an integrated spam-reporting dialog.
- Added support for Base64-encoded attachments in reply forms.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.15. To view API reference documentation for all APIs supported by Mailbox requirement set 1.15 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.15&preserve-view=true).

[!INCLUDE [outlook-1_15](../../includes/outlook-1_15.md)]

## Events

The following table lists the [events](/javascript/api/office/office.eventtype?view=outlook-js-1.15&preserve-view=true) introduced in requirement set 1.15. For a list of all supported events that can be handled using the `addHandlerAsync` and `removeHandlerAsync` methods, see [Office.EventType](/javascript/api/office/office.eventtype?view=outlook-js-1.15&preserve-view=true).

| Event | Description | Object |
| --- | --- | --- |
| [InitializationContextChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.15&preserve-view=true#fields) | The task pane of an add-in has been opened from an actionable message, `InsightMessage` notification, Smart Alerts dialog, or integrated spam-reporting dialog. | Item |

> [!NOTE]
> For events supported in an event-based activation add-in, see [Activate add-ins with events](/office/dev/add-ins/develop/event-based-activation#outlook-events).

## Manifest updates

The following table lists manifest updates introduced in Mailbox requirement set 1.15. To learn more about the types of Office Add-in manifests, see [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests).

| Feature | Unified manifest for Microsoft 365 | Add-in only manifest | Description |
| --- | --- | --- | --- |
| Spam reporting | [`"extensions.ribbons.spamPreProcessingDialog.spamReportingOptions.type"`](/microsoft-365/extensibility/schema/extension-ribbons-spam-pre-processing-dialog-spam-reporting-options#type) | [ReportingOptions](/javascript/api/manifest/reportingoptions) | Added support for radio buttons in the preprocessing dialog of an integrated spam-reporting add-in. |
| Spam reporting | [`"extensions.ribbons.spamPreProcessingDialog.spamNeverShowAgainOption"`](/microsoft-365/extensibility/schema/extension-ribbons-spam-pre-processing-dialog#spamnevershowagainoption) | [PreProcessingDialog](/javascript/api/manifest/preprocessingdialog#child-elements) | Added support for the "Don't show this message again" checkbox in the preprocessing dialog of an integrated spam-reporting add-in. |

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
