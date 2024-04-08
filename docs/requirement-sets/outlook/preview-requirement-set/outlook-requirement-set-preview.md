---
title: Outlook add-in API preview requirement set
description: Features and APIs that are currently in preview for Outlook add-ins.
ms.date: 04/08/2024
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API preview requirement set

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

Preview APIs are subject to change and aren't intended for use in a production environment. We recommend that you try them out in test and development environments only. Don't use preview APIs in a production environment or within business-critical documents.

To use preview APIs:

- You must use the preview version of the Office JavaScript API library from the [Office.js content delivery network (CDN)](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). You can install these types with `npm install --save-dev @types/office-js-preview` (be sure to remove the types for `@types/office-js` if you've previously installed them).

- You may need to join the [Microsoft 365 Insider program](https://insider.microsoft365.com/join) for access to more recent Office builds in Outlook on Windows and on Mac.

- You may need to configure the **Targeted release** option on your Microsoft 365 tenant to preview features in Outlook on the web. For more information, see the "Targeted release" section of [Set up the Standard or Targeted release options](/microsoft-365/admin/manage/release-options-in-office-365#targeted-release).

The preview requirement set includes all of the features of [requirement set 1.13](../requirement-set-1.13/outlook-requirement-set-1.13.md).

> [!IMPORTANT]
> This documentation is for a **preview** [requirement set](../outlook-api-requirement-sets.md). This requirement set isn't fully implemented yet, and clients won't accurately report support for it. You shouldn't specify this requirement set in your add-in manifest.

## Features in preview

The following features are in preview.

### Activate an event-based add-in on a message in read mode

#### [OnMessageReadWithCustomAttachment and OnMessageReadWithCustomHeader events](/office/dev/add-ins/outlook/autolaunch#supported-events)

Added events to activate an event-based add-in on a message in read mode when it contains certain attachment types or custom internet headers.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [HeaderName attribute in the LaunchEvent element](../../../manifest/launchevent.md#attributes)

Added an attribute to the **\<LaunchEvent\>** XML element to specify the internet header name on which the `OnMessageReadWithCustomHeader` event occurs.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [MessageAttachments element](../../../manifest/messageattachments.md)

Added an XML element to specify the file extension of the attachment included in a message on which the `OnMessageReadWithCustomAttachment` event occurs.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Additional calendar properties

#### [Office.IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Added a new object that represents the all-day event property of an appointment in compose mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Added a new object that represents the sensitivity level of an appointment in compose mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on Mac (Microsoft 365 subscription), Outlook on the web (modern), new Outlook on Windows (preview)

#### [Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Added a new property that represents if an appointment is an all-day event.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Added a new property that represents the sensitivity of an appointment.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on Mac (Microsoft 365 subscription), Outlook on the web (modern), new Outlook on Windows (preview)

#### [Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on Mac (Microsoft 365 subscription), Outlook on the web (modern), new Outlook on Windows (preview)

<br>

---

---

### Additional message members

#### [Office.context.mailbox.item.inReplyTo](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#outlook-office-messagecompose-inreplyto-member)

Added a property to get the message ID of the original message being replied to by the current message.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.context.mailbox.item.getConversationIndexAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#outlook-office-messagecompose-getconversationindexasync-member(1))

Added a method to get the Base64-encoded position of the current message in a conversation thread.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.context.mailbox.item.getItemClassAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#outlook-office-messagecompose-getitemclassasync-member(1))

Added a method to get the Exchange Web Services (EWS) item class of a message in compose mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Close and discard a message in compose

#### [Office.context.mailbox.item.closeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#outlook-office-messagecompose-closeasync-member(1))

Added method to close a current message being composed with the option to discard unsaved changes.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Get the URL of the JavaScript runtime of an add-in

#### [Office.context.urls.javascriptRuntimeUrl](/javascript/api/office/office.urls?view=common-js-preview&preserve-view=true#office-office-urls-javascriptruntimeurl-member)

Added property to get the URL of the JavaScript runtime of an add-in.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Integrated spam reporting

#### [ReportPhishingCommandSurface extension point](../../../manifest/extensionpoint.md#reportphishingcommandsurface-preview)

Added an extension point to activate your spam-reporting add-in in the Outlook ribbon and prevent it from appearing at the end of the ribbon or in the overflow section.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on Mac (Microsoft 365 subscription), Outlook on the web (modern), new Outlook on Windows (preview)

#### [ReportPhishingCustomization element](../../../manifest/reportphishingcustomization.md)

Added a manifest element to configure the ribbon button and preprocessing dialog of a spam-reporting add-in.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on Mac (Microsoft 365 subscription), Outlook on the web (modern), new Outlook on Windows (preview)

#### [Office.context.mailbox.item.getAsFileAsync](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#outlook-office-messageread-getasfileasync-member(1))

Added a method to get the Base64 encoding of a message.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on Mac (Microsoft 365 subscription), Outlook on the web (modern), new Outlook on Windows (preview)

#### [Office.SpamReportingEventCompletedOptions](/javascript/api/outlook/office.spamreportingeventcompletedoptions?view=outlook-js-preview&preserve-view=true)

Created options to customize a post-processing dialog or configure a spam-reporting add-in to perform additional operations on a reported message, such as deleting it from the inbox.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on Mac (Microsoft 365 subscription), Outlook on the web (modern), new Outlook on Windows (preview)

#### [Office.MailboxEnums.MoveSpamItemTo](/javascript/api/outlook/office.mailboxenums.movespamitemto?view=outlook-js-preview&preserve-view=true)

Added a new enum to specify the folder to which a reported message is moved once it's processed by a spam-reporting add-in.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on Mac (Microsoft 365 subscription), Outlook on the web (modern), new Outlook on Windows (preview)

To learn more about how to implement the integrated spam-reporting feature in your add-in, see [Implement an integrated spam-reporting add-in (preview)](/office/dev/add-ins/outlook/spam-reporting).

<br>

---

---

### Item multi-select: Additional message properties

#### [Office.SelectedItemDetails](/javascript/api/outlook/office.selecteditemdetails?view=outlook-js-preview&preserve-view=true)

The ability to get the properties of selected messages in Outlook using [Office.context.mailbox.getSelectedItemsAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-getselecteditemsasync-member(1)) was introduced in [requirement set 1.13](../requirement-set-1.13/outlook-requirement-set-1.13.md). Additional properties, such as `conversationId`, `internetMessageId`, and `hasAttachment`, are now available in preview.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Office theme

#### [Office.context.officeTheme](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#office-office-context-officetheme-member)

Added ability to get Office theme.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true)

Added `OfficeThemeChanged` event to `Mailbox`.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Smart Alerts: Customize the Don't Send option and override the send mode option at runtime

#### [Office.SmartAlertsEventCompletedOptions](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-preview&preserve-view=true): Additional options

Added additional `event.completed` options to customize the **Don't Send** button of the Smart Alerts dialog and override the send mode option at runtime.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.MailboxEnums.SendModeOverride](/javascript/api/outlook/office.mailboxenums.sendmodeoverride?view=outlook-js-preview&preserve-view=true)

Added an enum to specify the send mode option that overrides the option set in the manifest at runtime.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

To learn how to implement these features, see the [Smart Alerts walkthrough](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough).

<br>

---

---

### Temporarily set the body or subject displayed in read mode

#### [Office.context.mailbox.item.display](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#outlook-office-messageread-display-member)

Added a property that gets an object to temporarily set the content displayed in the body or subject of a message in read mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.Display](/javascript/api/outlook/office.display?view=outlook-js-preview&preserve-view=true)

Added an object that provides properties to temporarily set the content displayed in the body or subject of a message in read mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.DisplayedBody](/javascript/api/outlook/office.displayedbody?view=outlook-js-preview&preserve-view=true)

Added an object that provides a method to temporarily set the content displayed in the body of a message in read mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.DisplayedSubject](/javascript/api/outlook/office.displayedsubject?view=outlook-js-preview&preserve-view=true)

Added an object that provides a method to temporarily set the content displayed in the subject of a message in read mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
