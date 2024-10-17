---
title: Office.context.mailbox.item - requirement set 1.7
description: Outlook Mailbox API requirement set 1.7 version of the Item object model.
ms.date: 07/18/2024
ms.localizationpriority: medium
---

# item (Mailbox requirement set 1.7)

### [Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` is used to access the currently selected message, meeting request, or appointment. You can determine the type of the item by using the `itemType` property.

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../outlook-api-requirement-sets.md)|1.1|
|[Minimum permission level](/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions)|**restricted**|
|[Applicable Outlook mode](/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points)|Appointment Organizer, Appointment Attendee,<br>Message Compose, or Message Read|

> [!IMPORTANT]
> Android and iOS: There are limitations on when add-ins activate and which APIs are available. To learn more, refer to [Add mobile support to an Outlook add-in](/office/dev/add-ins/outlook/add-mobile-support#compose-mode-and-appointments).

## Properties

| Property | Minimum<br>permission level | Details by mode | Return type | Minimum<br>requirement set |
|---|---|---|---|:---:|
| attachments | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-attachments-member) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-attachments-member) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | **read item** | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-bcc-member) | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | **read item** | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-cc-member) | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-cc-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | **read item** | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-conversationid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-conversationid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-datetimecreated-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-datetimecreated-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-datetimemodified-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-datetimemodified-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-end-member) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-end-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-end-member)<br>(Meeting Request) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | **read/write item** | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-from-member) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | **read item** | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-from-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | **read item** | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-internetmessageid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-itemclass-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-itemclass-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-itemid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-itemid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-location-member) | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-location-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-location-member)<br>(Meeting Request) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-normalizedsubject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-normalizedsubject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-optionalattendees-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | **read/write item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-organizer-member) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-organizer-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-recurrence-member) | [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-recurrence-member) | [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-recurrence-member)<br>(Meeting Request) | [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-requiredattendees-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | **read item** | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-sender-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-seriesid-member) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-seriesid-member) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-seriesid-member) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-seriesid-member) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| start | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-start-member) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-start-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-start-member)<br>(Meeting Request) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-subject-member) | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-subject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-subject-member) | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-subject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| to | **read item** | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-to-member) | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-to-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## Methods

| Method | Minimum<br>permission level | Details by mode | Minimum<br>requirement set |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | **read/write item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-addfileattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-addfileattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-addhandlerasync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-addhandlerasync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-addhandlerasync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-addhandlerasync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | **read/write item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-additemattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-additemattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | **restricted** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-close-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-close-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-displayreplyallform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-displayreplyallform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-displayreplyform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-displayreplyform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities() **(deprecated)** | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getentities-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getentities-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) **(deprecated)** | **restricted** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getentitiesbytype-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getentitiesbytype-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName(name) **(deprecated)** | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getfilteredentitiesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getfilteredentitiesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches() | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getregexmatches-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getregexmatches-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName(name) | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getregexmatchesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getregexmatchesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-getselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-getselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() **(deprecated)** | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getselectedentities-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getselectedentities-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | **read item** | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getselectedregexmatches-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getselectedregexmatches-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | **read/write item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-removeattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-removeattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | **read item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-removehandlerasync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-removehandlerasync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-removehandlerasync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-removehandlerasync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | **read/write item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-saveasync-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-saveasync-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | **read/write item** | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-setselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-setselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## Events

You can subscribe to and unsubscribe from the following events using `addHandlerAsync` and `removeHandlerAsync` respectively.

> [!IMPORTANT]
> Events are only available with task pane implementation.

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-1.7&preserve-view=true) | Description | Minimum<br>requirement set |
|---|---|:---:|
|`AppointmentTimeChanged`| The date or time of the selected appointment or series has changed. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecipientsChanged`| The recipient list of the selected item or appointment location has changed. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| The recurrence pattern of the selected series has changed. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

## Example

The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready method.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    const item = Office.context.mailbox.item;
    const subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```
