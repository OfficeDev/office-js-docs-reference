---
title: Outlook Item object model
description: Learn more about the Outlook Item object model and its APIs.
ms.date: 03/18/2026
ms.topic: overview
ms.localizationpriority: high
---

# Outlook Item object model

Use the ``Office.context.mailbox.item`` object to access and perform operations on messages and appointments in read or compose mode.

> [!IMPORTANT]
> Android and iOS: There are limitations on when add-ins activate and which APIs are available. To learn more, refer to [Add mobile support to an Outlook add-in](/office/dev/add-ins/outlook/add-mobile-support#compose-mode-and-appointments).

## Properties

[!INCLUDE [Outlook item properties](../../includes/outlook-item-object-model-properties.md)]

## Methods

[!INCLUDE [Outlook item methods](../../includes/outlook-item-object-model-methods.md)]

## Events

You can subscribe to and unsubscribe from the following events using `addHandlerAsync` and `removeHandlerAsync` respectively. For more information, see [Office.EventType](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true).

[!INCLUDE [Outlook item events](../../includes/outlook-item-object-model-events.md)]

> [!NOTE]
> For events supported in an event-based activation add-in, see [Activate add-ins with events](/office/dev/add-ins/develop/event-based-activation#outlook-events).
