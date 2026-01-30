---
title: Outlook add-in API requirement set 1.3
description: Lists the APIs introduced in Mailbox requirement set 1.3 for Outlook add-ins.
ms.date: 01/30/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.3

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.3

Mailbox requirement set 1.3 includes all of the features of [requirement set 1.2](outlook-requirement-set-1-2.md). It added the following features.

- Added support for [function commands](/javascript/api/office/office.addincommands.event?view=outlook-js-1.3&preserve-view=true). For more information, see [Add-in commands](/office/dev/add-ins/design/add-in-commands).
- Added a method to save an item being composed as a draft.
- Added a method to close an item being composed.
- Added methods to get or set the entire body of a message or appointment. For more information, see [Get or set the body of a message or appointment in Outlook](/office/dev/add-ins/outlook/insert-data-in-the-body).
- Added methods to convert item IDs between EWS and REST formats.

    [!INCLUDE [legacy-exchange-online](../../includes/legacy-exchange-online.md)]
- Added a property and methods to add, get, replace, and remove [notification messages](/office/dev/add-ins/outlook/notifications) from an item.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.3. To view API reference documentation for all APIs supported by Mailbox requirement set 1.3 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.3&preserve-view=true).

[!INCLUDE [outlook-1_3](../../includes/outlook-1_3.md)]

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
