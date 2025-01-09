---
title: Outlook add-in API preview requirement set
description: Features and APIs that are currently in preview for Outlook add-ins.
ms.date: 11/14/2024
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

The preview requirement set includes all of the features of [requirement set 1.14](../requirement-set-1.14/outlook-requirement-set-1.14.md).

> [!IMPORTANT]
> This documentation is for a **preview** [requirement set](../outlook-api-requirement-sets.md). This requirement set isn't fully implemented yet, and clients won't accurately report support for it. You shouldn't specify this requirement set in your add-in manifest.

## Features in preview

The following features are in preview.

### Activate an event-based add-in on a message in read mode

#### [OnMessageReadWithCustomAttachment and OnMessageReadWithCustomHeader events](/office/dev/add-ins/outlook/autolaunch#supported-events)

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

### Item multi-select: Get additional message properties and run operations on multiple selected messages

#### [Office.context.mailbox.loadItemByIdAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-loaditembyidasync-member(1))

Added a new method to get additional properties and run operations on selected messages.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [Office.LoadedMessageCompose](/javascript/api/outlook/office.loadedmessagecompose?view=outlook-js-preview&preserve-view=true)

Added a new object that represents the properties and methods of a selected message in compose mode.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

#### [Office.LoadedMessageRead](/javascript/api/outlook/office.loadedmessageread?view=outlook-js-preview&preserve-view=true)

Added a new object that represents the properties and methods of a selected message in read mode.

**Available in**: Classic Outlook on Windows (Microsoft 365 subscription)

<br>

---

---

### Smart Alerts: Format the dialog message using Markdown

#### [Office.SmartAlertsEventCompletedOptions.errorMessageMarkdown](/javascript/api/outlook/office.smartalertseventcompletedoptions?view=outlook-js-preview&preserve-view=true#outlook-office-smartalertseventcompletedoptions-errormessagemarkdown-member)

Added an `event.completed` option to format a message in a Smart Alerts dialog using Markdown. To learn more, see the [Smart Alerts walkthrough](/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough) and [Limitations to formatting the dialog message using Markdown](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events#limitations-to-formatting-the-dialog-message-using-markdown).

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
