---
title: Outlook add-in API preview requirement set
description: Features and APIs that are currently in preview for Outlook add-ins.
ms.date: 04/18/2023
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API preview requirement set

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.

To use preview APIs:

- You must use the preview version of the Office JavaScript API library from the [Office.js content delivery network (CDN)](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). You can install these types with `npm install --save-dev @types/office-js-preview` (be sure to remove the types for `@types/office-js` if you've previously installed them).

- You may need to join the [Office Insider program](https://insider.office.com) for access to more recent Office builds in Outlook on Windows and on Mac.

- You may need to configure the **Targeted release** option on your Microsoft 365 tenant to preview features in Outlook on the web. For more information, see the "Targeted release" section of [Set up the Standard or Targeted release options](/microsoft-365/admin/manage/release-options-in-office-365#targeted-release).

The preview requirement set includes all of the features of [requirement set 1.12](../requirement-set-1.12/outlook-requirement-set-1.12.md).

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

### Delay delivery time

#### [Office.context.mailbox.item.delayDeliveryTime](office.context.mailbox.item.md#properties)

Added a new property that returns an object that allows you to manage the delivery date and time of a message in compose mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on the web (modern)

#### [Office.DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime?view=outlook-js-preview&preserve-view=true)

Added a new object that allows you to manage the delivery date and time of a message in compose mode.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on the web (modern)

<br>

---

---

### Event-based activation: OnMessageFromChanged and OnAppointmentFromChanged events

Added support for the `OnMessageFromChanged` and `OnAppointmentFromChanged` events in [event-based activation add-ins](/office/dev/add-ins/outlook/autolaunch). To learn more about these events, see [Automatically update your signature when switching between mail accounts (preview)](/office/dev/add-ins/outlook/onmessagefromchanged-onappointmentfromchanged-events).

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on Mac (Microsoft 365 subscription), Outlook on the web (modern)

<br>

---

---

### Item multi-select

#### [Office.context.mailbox.getSelectedItemsAsync](office.context.mailbox.md#methods)

Added a new method that retrieves currently selected messages. To learn more about item multi-select, see [Activate your Outlook add-in on multiple messages (preview)](/office/dev/add-ins/outlook/item-multi-select).

**Available in**: Outlook on Windows (Microsoft 365 subscription)

#### [Office.EventType.SelectedItemsChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true)

Added `SelectedItemsChanged` event to `Mailbox`. This event occurs when one or more messages are selected or deselected.

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

### Prepend content on send

#### [Office.context.mailbox.item.body.prependOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-prependonsendasync-member(1))

Added method to prepend content to the beginning of a message or appointment body when the mail item is sent.

**Available in**: Outlook on Windows (Microsoft 365 subscription), Outlook on Mac (Microsoft 365 subscription), Outlook on the web (modern)

<br>

---

---

### Manage the sensitivity label of a message or appointment

#### [Office.context.sensitivityLabelsCatalog](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#office-office-context-sensitivitylabelscatalog-member)

Added a property that provides the object to check the status of the catalog of sensitivity labels and retrieve all available sensitivity labels if the catalog is enabled.

**Available in**: Outlook on Windows (Microsoft 365 E5 subscription), Outlook on Mac (Microsoft 365 E5 subscription), Outlook on the web (modern)

#### [Office.context.mailbox.item.sensitivityLabel](office.context.mailbox.item.md#properties)

Added a property that provides the object to get or set the sensitivity label of a message or appointment in compose mode.

**Available in**: Outlook on Windows (Microsoft 365 E5 subscription), Outlook on Mac (Microsoft 365 E5 subscription), Outlook on the web (modern)

#### [Office.EventType.SensitivityLabelChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true)

Added the `SensitivityLabelChanged` event to `Item`. This event occurs when the sensitivity label of a message or appointment is changed.

**Available in**: Outlook on Windows (Microsoft 365 E5 subscription), Outlook on Mac (Microsoft 365 E5 subscription), Outlook on the web (modern)

#### [Office.SensitivityLabelChangedEventArgs](/javascript/api/outlook/office.sensitivitylabelchangedeventargs?view=outlook-js-preview&preserve-view=true)

Added an object that provides the change status of the sensitivity label applied to a message or appointment in compose mode.

**Available in**: Outlook on Windows (Microsoft 365 E5 subscription), Outlook on Mac (Microsoft 365 E5 subscription), Outlook on the web (modern)

#### [Office.SensitivityLabelsCatalog](/javascript/api/outlook/office.sensitivitylabelscatalog?view=outlook-js-preview&preserve-view=true)

Added an object that represents the catalog of sensitivity labels in Outlook.

**Available in**: Outlook on Windows (Microsoft 365 E5 subscription), Outlook on Mac (Microsoft 365 E5 subscription), Outlook on the web (modern)

#### [Office.SensitivityLabel](/javascript/api/outlook/office.sensitivitylabel?view=outlook-js-preview&preserve-view=true)

Added an object that represents the sensitivity label of a message or appointment in compose mode.

**Available in**: Outlook on Windows (Microsoft 365 E5 subscription), Outlook on Mac (Microsoft 365 E5 subscription), Outlook on the web (modern)

#### [Office.SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails?view=outlook-js-preview&preserve-view=true)

Added an object that represents the properties of a sensitivity label.

**Available in**: Outlook on Windows (Microsoft 365 E5 subscription), Outlook on Mac (Microsoft 365 E5 subscription), Outlook on the web (modern)

<br>

---

---

### Shared mailboxes

Feature support for shared folders (that is, delegate access) was released in [requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). However, support for shared mailboxes is now available in preview. To learn more, see [Enable shared folders and shared mailbox scenarios](/office/dev/add-ins/outlook/delegate-access).

**Available in**: Outlook on Windows (Exchange Online or on-premises Exchange environment), Outlook on Mac

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
