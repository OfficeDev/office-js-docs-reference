---
title: Outlook add-in API preview requirement set
description: Features and APIs that are currently in preview for Outlook add-ins.
ms.date: 06/21/2023
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API preview requirement set

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.

To use preview APIs:

- You must use the preview version of the Office JavaScript API library from the [Office.js content delivery network (CDN)](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). You can install these types with `npm install --save-dev @types/office-js-preview` (be sure to remove the types for `@types/office-js` if you've previously installed them).

- You may need to join the [Microsoft 365 Insider program](https://insider.microsoft365.com/join) for access to more recent Office builds in Outlook on Windows and on Mac.

- You may need to configure the **Targeted release** option on your Microsoft 365 tenant to preview features in Outlook on the web. For more information, see the "Targeted release" section of [Set up the Standard or Targeted release options](/microsoft-365/admin/manage/release-options-in-office-365#targeted-release).

The preview requirement set includes all of the features of [requirement set 1.13](../requirement-set-1.13/outlook-requirement-set-1.13.md).

> [!IMPORTANT]
> This documentation is for a **preview** [requirement set](../outlook-api-requirement-sets.md). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest.

## Features in preview

The following features are in preview.

### Additional calendar properties

#### [Office.IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Added a new object that represents the all-day event property of an appointment in compose mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Added a new object that represents the sensitivity level of an appointment in compose mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Added a new property that represents if an appointment is an all-day event.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Added a new property that represents the sensitivity of an appointment.

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.

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

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
