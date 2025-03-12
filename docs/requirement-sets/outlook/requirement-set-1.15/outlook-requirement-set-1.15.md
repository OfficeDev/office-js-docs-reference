---
title: Outlook add-in API requirement set 1.15
description: Requirement set 1.15 for Outlook add-in API.
ms.date: 03/11/2025
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.15

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

## What's new in 1.15?

Requirement set 1.15 includes all of the features of [requirement set 1.14](../requirement-set-1.14/outlook-requirement-set-1.14.md). It added the following features.

- Added support for radio buttons to format reporting options in the [integrated spam-reporting](/office/dev/add-ins/outlook/spam-reporting) dialog.
- Added support to include a "Don't show this message again" checkbox in a spam-reporting dialog that doesn't require input from a user.
- Added support to open a task pane from a spam-reporting dialog.
- Added support for Markdown to format error messages in a [Smart Alerts](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events) dialog.
- Added support to run a function from the Smart Alerts dialog.
- Added a method to programmatically send a mail item.
- Added a method to load a single message from multiple selected messages to get its properties or perform operations on it.
- Added an event that occurs when an add-in's task pane is opened from an [actionable message](/outlook/actionable-messages), [InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.15&preserve-view=true#fields) notification, Smart Alerts dialog, or an integrated spam-reporting dialog.
- Added support for Base64-encoded attachments in reply forms.

### Change log

- Added the `inputType` attribute to the [ReportingOptions](/javascript/api/manifest/reportingoptions?view=outlook-js-1.15&preserve-view=true#attributes) element of the add-in only manifest: When implementing the integrated spam-reporting feature, uses radio buttons to format reporting options in the preprocessing dialog.
- Added the [NeverShowAgainOption](/javascript/api/manifest/preprocessingdialog?view=outlook-js-1.15&preserve-view=true#child-elements) add-in only manifest element: When implementing the integrated spam-reporting feature, adds a "Don't show this message again" checkbox to the preprocessing dialog. This checkbox is supported in preprocessing dialogs that don't require input from a user.
- Added new properties to [Office.SpamReportingEventCompletedOptions](/javascript/api/outlook/office.spamreportingeventcompletedoptions?view=outlook-js-1.15&preserve-view=true): The following properties configure a task pane to open from the **Report** button of the preprocessing dialog.
  - [commandId](/javascript/api/outlook/office.spamreportingeventcompletedoptions?view=outlook-js-1.15&preserve-view=true#outlook-office-spamreportingeventcompletedoptions-commandid-member) property: When the **Report** option is selected from a preprocessing dialog, specifies the ID of the task pane that opens.
  - [contextData](/javascript/api/outlook/office.spamreportingeventcompletedoptions?view=outlook-js-1.15&preserve-view=true#outlook-office-spamreportingeventcompletedoptions-contextdata-member) property: When the **Report** option is selected from a preprocessing dialog, specifies any JSON data passed to the add-in's task pane for processing.
- Added the [errorMessageMarkdown](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.15&preserve-view=true#outlook-office-smartalertseventcompletedoptions-errormessagemarkdown-member) property to [Office.SmartAlertsEventCompletedOptions](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.15&preserve-view=true): Supports Markdown to format the error message shown in a Smart Alerts dialog.
- Updated the [commandId](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.15&preserve-view=true#outlook-office-smartalertseventcompletedoptions-commandid-member) and [contextData](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.15&preserve-view=true#outlook-office-smartalertseventcompletedoptions-contextdata-member) properties of [Office.SmartAlertsEventCompletedOptions](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.15&preserve-view=true): Now supports running a function from the Smart Alerts dialog.
- Added the [sendAsync](office.context.mailbox.item.md#methods) method: Programmatically sends a mail item.
- Added the [Office.context.mailbox.loadItemByIdAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.15&preserve-view=true#outlook-office-mailbox-loaditembyidasync-member(1)) method: When the [item multi-select](/office/dev/add-ins/outlook/item-multi-select) feature is implemented, loads a single item to perform operations on it or get properties that aren't provided by the [getSelectedItemsAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.15&preserve-view=true#outlook-office-mailbox-getselecteditemsasync-member(1)) method.
- Added the [Office.LoadedMessageCompose](/javascript/api/outlook/office.loadedmessagecompose?view=outlook-js-1.15&preserve-view=true) and [Office.LoadedMessageRead](/javascript/api/outlook/office.loadedmessageread?view=outlook-js-1.15&preserve-view=true) objects: Represents the properties and methods of a mail item that's currently loaded using the `loadItemByIdAsync` method.
- Added the [Office.EventType.InitializationContextChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.15&preserve-view=true#fields) event: Occurs when an add-in's task pane is opened from an actionable message, `InsightMessage` notification, Smart Alerts dialog, or integrated spam-reporting dialog.
- Added [Office.InitializationContextChangedEventArgs](/javascript/api/outlook/office.initializationcontextchangedeventargs?view=outlook-js-1.15&preserve-view=true): When the `Office.EventType.InitializationContextChanged` event occurs, provides data from an actionable message, `InsightMessage` notification, Smart Alerts dialog, or integrated spam-reporting dialog to an add-in's task pane.
- Updated the [type](/javascript/api/outlook/office.replyformattachment?view=outlook-js-1.15&preserve-view=true#outlook-office-replyformattachment-type-member) property of [Office.ReplyFormAttachment](/javascript/api/outlook/office.replyformattachment?view=outlook-js-1.15&preserve-view=true): Now supports Base64-encoded attachments in reply forms.
- Added the [base64File](/javascript/api/outlook/office.replyformattachment?view=outlook-js-1.15&preserve-view=true#outlook-office-replyformattachment-base64file-member) property to `Office.ReplyFormAttachment`: Specifies the Base64-encoded string of the file to be attached to a reply form.
- Added the [Office.MailboxEnums.AttachmentType.Base64](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.15&preserve-view=true#fields) enum: Specifies that the attachment is a Base64-encoded file. This attachment type is only supported by the [displayReplyAllForm](office.context.mailbox.item.md#methods), [displayReplyAllFormAsync](office.context.mailbox.item.md#methods), [displayReplyForm](office.context.mailbox.item.md#methods), and [displayReplyFormAsync](office.context.mailbox.item.md#methods) methods.

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
