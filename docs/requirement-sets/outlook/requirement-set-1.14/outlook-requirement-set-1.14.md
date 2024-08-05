---
title: Outlook add-in API requirement set 1.14
description: Requirement set 1.14 for Outlook add-in API.
ms.date: 05/20/2024
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.14

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

## What's new in 1.14?

Requirement set 1.14 includes all of the features of [requirement set 1.13](../requirement-set-1.13/outlook-requirement-set-1.13.md). It added the following features.

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

### Change log

- Added the [ReportPhishingCommandSurface](/javascript/api/manifest/extensionpoint?view=outlook-js-1.14&preserve-view=true#reportphishingcommandsurface) add-in only manifest extension point: Activates a spam-reporting add-in in the Outlook ribbon and prevents it from appearing at the end of the ribbon or in the overflow section.
- Added the [ReportPhishingCustomization](/javascript/api/manifest/reportphishingcustomization?view=outlook-js-1.14&preserve-view=true) add-in only manifest element: Configures the ribbon button and preprocessing dialog of a spam-reporting add-in.
- Added the [Office.EventType.SpamReporting](/javascript/api/office/office.eventtype?view=outlook-js-1.14&preserve-view=true#fields) event: Occurs in Outlook when an unsolicited message is reported.
- Added [Office.SpamReportingEventArgs](/javascript/api/outlook/office.spamreportingeventargs?view=outlook-js-1.14&preserve-view=true): Provides information about the `Office.EventType.SpamReporting` event that occurs when an unsolicited message is reported.
- Added [Office.SpamReportingEventCompletedOptions](/javascript/api/outlook/office.spamreportingeventcompletedoptions?view=outlook-js-1.14&preserve-view=true): Provides options to customize the post-processing dialog of a spam-reporting add-in and run additional operations on a reported message.
- Added the [Office.MailboxEnums.MoveSpamItemTo](/javascript/api/outlook/office.mailboxenums.movespamitemto?view=outlook-js-1.14&preserve-view=true) enum: Specifies the folder to which a reported message is moved once it's processed by a spam-reporting add-in.
- Added the [Office.context.mailbox.item.getAsFileAsync](/javascript/api/outlook/office.messageread?view=outlook-js-1.14&preserve-view=true#outlook-office-messageread-getasfileasync-member(1)) method: Gets the Base64 encoding of a message.
- Added [Office.Urls](/javascript/api/office/office.urls): Provides an object to get the URLs of the runtime environments used by an add-in.
- Added the [Office.context.urls.javascriptRuntimeUrl](/javascript/api/office/office.urls?view=outlook-js-1.14&preserve-view=true#office-office-urls-javascriptruntimeurl-member) method: Gets the URL of the JavaScript runtime of an add-in.
- Added new properties to [Office.SmartAlertsEventCompletedOptions](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.14&preserve-view=true): Adds the following properties to customize the **Don't Send** option of a Smart Alerts dialog and override the send mode option at runtime.

  - [cancelLabel](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.14&preserve-view=true#outlook-office-smartalertseventcompletedoptions-cancellabel-member) property: Customizes the text of the **Don't Send** option of a Smart Alerts dialog.
  - [commandId](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.14&preserve-view=true#outlook-office-smartalertseventcompletedoptions-commandid-member) property: Specifies the ID of the task pane that opens when the **Don't Send** option is selected from a Smart Alerts dialog.
  - [contextData](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.14&preserve-view=true#outlook-office-smartalertseventcompletedoptions-contextdata-member) property: Specifies any JSON data passed to the add-in for processing when the **Don't Send** option is selected from a Smart Alerts dialog.
  - [sendModeOverride](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-1.14&preserve-view=true#outlook-office-smartalertseventcompletedoptions-sendmodeoverride-member) property: Overrides the send mode option specified in the manifest at runtime.

- Added the [Office.MailboxEnums.SendModeOverride](/javascript/api/outlook/office.mailboxenums.sendmodeoverride?view=outlook-js-1.14&preserve-view=true) enum: Specifies the send mode option that overrides the option set in the manifest at runtime.
- Added the [Office.context.mailbox.item.inReplyTo](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.14&preserve-view=true#outlook-office-messagecompose-inreplyto-member) property:
Gets the message ID of the original message being replied to by the current message.
- Added the [Office.context.mailbox.item.getConversationIndexAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.14&preserve-view=true#outlook-office-messagecompose-getconversationindexasync-member(1)) method: Gets the Base64-encoded position of the current message in a conversation thread.
- Added the [Office.context.mailbox.item.getItemClassAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.14&preserve-view=true#outlook-office-messagecompose-getitemclassasync-member(1)) method: Gets the Exchange Web Services (EWS) item class of a message in compose mode.
- Added the [Office.context.mailbox.item.closeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.14&preserve-view=true#outlook-office-messagecompose-closeasync-member(1)) method: Closes a current message being composed with the option to discard unsaved changes.
- Added new properties to [Office.SelectedItemDetails](/javascript/api/outlook/office.selecteditemdetails?view=outlook-js-1.14&preserve-view=true): Adds the following supported properties to the item multi-select feature.

  - [conversationId](/javascript/api/outlook/office.selecteditemdetails?view=outlook-js-1.14&preserve-view=true#outlook-office-selecteditemdetails-conversationid-member) property: Provides the  identifier of the message conversation that contains the message that's currently selected.
  - [hasAttachment](/javascript/api/outlook/office.selecteditemdetails?view=outlook-js-1.14&preserve-view=true#outlook-office-selecteditemdetails-hasattachment-member) property: Identifies whether a message that's currently selected contains an attachment.
  - [internetMessageId](/javascript/api/outlook/office.selecteditemdetails?view=outlook-js-1.14&preserve-view=true#outlook-office-selecteditemdetails-internetmessageid-member) property: Provides the internet message identifier of the message that's currently selected.

- Added the [Office.context.officeTheme](/javascript/api/office/office.context?view=outlook-js-1.14&preserve-view=true#office-office-context-officetheme-member) property: Gets the object to access the properties of the currently selected Office theme.
- Added the [Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.14&preserve-view=true) event: Occurs when the Office theme is changed in Outlook.
- Added [Office.OfficeThemeChangedEventArgs](/javascript/api/outlook/office.officethemechangedeventargs?view=outlook-js-1.14&preserve-view=true): Provides the updated Office theme when the `Office.EventType.OfficeThemeChanged` event occurs.
- Added the [Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties) property: Represents the sensitivity level of an appointment.
- Added [Office.Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-1.14&preserve-view=true): Provides methods to get or set the sensitivity level of an appointment in compose mode.
- Added the [Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-1.14&preserve-view=true) enum: Specifies the sensitivity level of an appointment.

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
