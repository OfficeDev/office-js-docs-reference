---
title: Outlook add-in API requirement set 1.1
description: Features and APIs that were introduced for Outlook add-ins and the Office JavaScript APIs as part of Mailbox API 1.1.
ms.date: 11/19/2024
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.1

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in. Outlook JavaScript API 1.1 (Mailbox 1.1) is the first version of the API.

> [!NOTE]
> This documentation is for a [requirement set](../outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.1?

Requirement set 1.1 includes all of the [Common API requirement sets](../../common/office-add-in-requirement-sets.md) supported in Outlook. It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.

[!INCLUDE [outlook-contextual-add-ins-retirement](../../../includes/outlook-contextual-add-ins-retirement.md)]

### Change log

- Added [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true) object: Provides methods for adding and updating the content of an item in an Outlook add-in.
- Added [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1&preserve-view=true) object: Provides methods to get and set the location of a meeting in an Outlook add-in.
- Added [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.
- Added [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1&preserve-view=true) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.
- Added [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1&preserve-view=true) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.
- Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): Adds a file as an attachment to a message or appointment. This method is available in Outlook on Windows (classic) and on Mac.
- Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods): Adds an Exchange item, such as a message, as an attachment to the message or appointment.
- Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods): Removes an attachment from a message or appointment.
- Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties): Gets an object that provides methods for manipulating the body of an item.
- Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#properties) line of a message.
- Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1&preserve-view=true): Specifies the type of recipient for an appointment.

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
