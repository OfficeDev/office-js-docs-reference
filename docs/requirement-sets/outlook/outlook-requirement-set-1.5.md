---
title: Outlook add-in API requirement set 1.5
description: Lists the APIs introduced in Mailbox requirement set 1.5 for Outlook add-ins.
ms.date: 01/27/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.5

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.5

Mailbox requirement set 1.5 includes all of the features of [requirement set 1.4](outlook-requirement-set-1.4.md). It added the following features.

- Added support for [pinnable task panes](/office/dev/add-ins/outlook/pinnable-taskpane).
- Added support for calling [REST APIs](/office/dev/add-ins/outlook/use-rest-api).
    > [!IMPORTANT]
    > Outlook REST v2.0 and beta endpoints are deprecated. Use the [Microsoft Graph REST API](/office/dev/add-ins/outlook/microsoft-graph) instead.
- Added support to mark an attachment as inline.
- Added support to programmatically close a task pane or dialog.
- Added support for the [Office.context.diagnostics](/javascript/api/office/office.context#office-office-context-diagnostics-member) property and its related objects.
- Added an event to determine when an Outlook mail item is selected for viewing while the add-in's task pane is pinned.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.5. To view API reference documentation for all APIs supported by Mailbox requirement set 1.5 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.5&preserve-view=true).

[!INCLUDE [outlook-1_5](../../includes/outlook-1_5.md)]

## Events

The following table lists the [events](/javascript/api/office/office.eventtype?view=outlook-js-1.5&preserve-view=true) introduced in requirement set 1.5. For a list of all supported events that can be handled using the `addHandlerAsync` and `removeHandlerAsync` methods, see [Office.EventType](/javascript/api/office/office.eventtype?view=outlook-js-1.5&preserve-view=true).

| Event | Description | Object |
| --- | --- | --- |
| [ItemChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.5&preserve-view=true#fields) | A different Outlook item is selected for viewing while the add-in's task pane is pinned. | Mailbox |

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
