---
title: Outlook add-in API preview requirement set
description: Features and APIs that are currently in preview for Outlook add-ins.
ms.date: 11/20/2025
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

The preview requirement set includes all of the features of [requirement set 1.15](../requirement-set-1.15/outlook-requirement-set-1.15.md).

> [!IMPORTANT]
> This documentation is for a **preview** [requirement set](../outlook-api-requirement-sets.md). This requirement set isn't fully implemented yet, and clients won't accurately report support for it. You shouldn't specify this requirement set in your add-in manifest.

## Features in preview

The following features are in preview.

### Activate an event-based add-in on a message in read mode

#### [OnMessageReadWithCustomAttachment and OnMessageReadWithCustomHeader events](/office/dev/add-ins/develop/event-based-activation#supported-events)

Added events to activate an event-based add-in on a message in read mode when it contains certain attachment types or custom internet headers.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [HeaderName attribute in the LaunchEvent element](../../../manifest/launchevent.md#attributes)

Added an attribute to the **\<LaunchEvent\>** XML element to specify the internet header name on which the `OnMessageReadWithCustomHeader` event occurs.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [MessageAttachments element](../../../manifest/messageattachments.md)

Added an XML element to specify the file extension of the attachment included in a message on which the `OnMessageReadWithCustomAttachment` event occurs.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Additional calendar properties

#### [Office.IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Added a new object that represents the all-day event property of an appointment in compose mode.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Added a new property that represents if an appointment is an all-day event.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Check if Exchange Web Services (EWS) tokens are supported in an organization

#### [Office.MailboxEnums.TokenStatus](/javascript/api/outlook/office.mailboxenums.tokenstatus?view=outlook-js-preview&preserve-view=true)

Added an enum to specify the status of tokens in an organization.

**Available in**: Outlook on the web and on Windows (new and classic)

#### [Office.context.mailbox.diagnostics.ews.getTokenStatusAsync](/javascript/api/outlook/office.ews?view=outlook-js-preview&preserve-view=true#outlook-office-ews-gettokenstatusasync-member(1))

Added a method to get the status of EWS callback tokens in an organization.

**Available in**: Outlook on the web and on Windows (new and classic)

<br>

---

---

### Customize the Smart Alerts dialog message using Markdown in Outlook on Mac

#### [errorMessageMarkdown](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-preview&preserve-view=true#outlook-office-smartalertseventcompletedoptions-errormessagemarkdown-member) property of the `event.completed` method

Updated the `errorMessageMarkdown` property of the `event.completed` method to include support in Outlook on Mac.

**Available in**: Outlook on Mac

<br>

---

---

### Decrypt a message and its attachments

#### [OnMessageRead event](/office/dev/add-ins/develop/event-based-activation#outlook-events)

Added a decryption event that occurs when the header of an encrypted message matches the header key of an installed encryption add-in.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [HeaderName attribute in the LaunchEvent element](../../../manifest/launchevent.md#attributes)

Updated the `HeaderName` attribute of the **\<LaunchEvent\>** XML element to specify the header key used for decryption.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [Event.completed method](/javascript/api/outlook/office.mailboxevent?view=outlook-js-preview&preserve-view=true#outlook-office-mailboxevent-completed-member(1))

Updated the `event.completed` method to indicate when an encryption add-in has completed processing the `OnMessageRead` event.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [Office.MessageDecryptEventCompletedOptions](/javascript/api/outlook/office.messagedecrypteventcompletedoptions?view=outlook-js-preview&preserve-view=true)

Added an object to specify the behavior of an encryption add-in after it completes processing an `OnMessageRead` event.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [Office.DecryptedMessageAttachment](/javascript/api/outlook/office.decryptedmessageattachment?view=outlook-js-preview&preserve-view=true)

Added an object that represents an attachment in a decrypted message.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [Office.DecryptedMessageBody](/javascript/api/outlook/office.decryptedmessagebody?view=outlook-js-preview&preserve-view=true)

Added an object that represents the body of a decrypted message.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Increase the number of recipients in a target field

### [Office.Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)

Updated the Recipients APIs to increase the maximum number of recipients in a target field to 1,000.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Store even more custom data for a mail item during an Outlook session

#### [Office.context.mailbox.item.sessionData.setAsync](/javascript/api/outlook/office.sessiondata?view=outlook-js-preview&preserve-view=true#outlook-office-sessiondata-setasync-member)

Increased the `SessionData` object limit to 2,621,440 characters.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Temporarily set the body or subject displayed in read mode

#### [Office.context.mailbox.item.display](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#outlook-office-messageread-display-member)

Added a property that gets an object to temporarily set the content displayed in the body or subject of a message in read mode.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [Office.Display](/javascript/api/outlook/office.display?view=outlook-js-preview&preserve-view=true)

Added an object that provides properties to temporarily set the content displayed in the body or subject of a message in read mode.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [Office.DisplayedBody](/javascript/api/outlook/office.displayedbody?view=outlook-js-preview&preserve-view=true)

Added an object that provides a method to temporarily set the content displayed in the body of a message in read mode.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [Office.DisplayedSubject](/javascript/api/outlook/office.displayedsubject?view=outlook-js-preview&preserve-view=true)

Added an object that provides a method to temporarily set the content displayed in the subject of a message in read mode.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
