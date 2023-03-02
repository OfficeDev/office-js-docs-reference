---
title: Outlook add-in API requirement set 1.5
description: Features and APIs that were introduced for Outlook add-ins and the Office JavaScript APIs as part of Mailbox API 1.5.
ms.date: 04/06/2022
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.5

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.5?

Requirement set 1.5 includes all of the features of [requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.

- Added support for [pinnable task panes](/office/dev/add-ins/outlook/pinnable-taskpane).
- Added support for calling [REST APIs](/office/dev/add-ins/outlook/use-rest-api).
- Added ability to mark an attachment as inline.
- Added ability to close a task pane or dialog.
- Added support for the [Office.context.diagnostics](office.context.md#diagnostics-contextinformation) property and its related objects.

### Change log

- Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): Adds an event handler for a supported event.
- Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): Removes the event handlers for a supported event type.
- Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.
- Added [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): Gets the URL of the REST endpoint for this email account.
- Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.
- Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui?view=outlook-js-1.5&preserve-view=true#office-office-ui-closecontainer-member(1)).
- Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Added [Office.ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.5&preserve-view=true): Provides information about the environment in which the add-in is running.
- Added [Office.context.diagnostics](office.context.md#diagnostics-contextinformation): Gets information about the environment in which the add-in is running, including host, platform, and version information.
- Added [Office.context.host](office.context.md#host-hosttype): Gets the Office application that is hosting the add-in.
- Added [Office.context.platform](office.context.md#platform-platformtype): Gets the platform on which the add-in is running.
- Added [Office.HostType](office.md#hosttype-string): Specifies the host Office application in which the add-in is running.
- Added [Office.PlatformType](office.md#platformtype-string): Specifies the OS or other platform on which the Office host application is running.

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
