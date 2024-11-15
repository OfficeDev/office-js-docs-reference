---
title: Office.context.mailbox - preview requirement set
description: Outlook Mailbox API preview requirement set version of the Mailbox object model.
ms.date: 11/14/2024
ms.localizationpriority: medium
---

# mailbox (preview requirement set)

### [Office](office.md)[.context](office.context.md).mailbox

Provides access to the Outlook add-in object model for Microsoft Outlook.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../outlook-api-requirement-sets.md)| 1.1|
|[Minimum permission level](/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions)| **restricted**|
|[Applicable Outlook mode](/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points)| Compose or Read|

## Properties

| Property | Minimum<br>permission level | Modes | Return type | Minimum<br>requirement set |
|---|---|---|---|:---:|
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-diagnostics-member) | **read item** | Compose<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-ewsurl-member) | **read item** | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | **restricted** | Compose<br>Read | [Item](/javascript/api/outlook/office.item?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-mastercategories-member) | **read/write mailbox** | Compose<br>Read | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-preview&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-resturl-member) | **read item** | Compose<br>Read | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-userprofile-member) | **read item** | Compose<br>Read | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## Methods

| Method | Minimum<br>permission level | Modes | Minimum<br>requirement set |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) | **read item** | Compose<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId(id, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-converttoewsid-member(1)) | **restricted** | Compose<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-converttolocalclienttime-member(1)) | **read item** | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(id, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-converttorestid-member(1)) | **restricted** | Compose<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-converttoutcclienttime-member(1)) | **read item** | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-displayappointmentform-member(1)) | **read item** | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-displayappointmentformasync-member(1)) | **read item** | Compose<br>Read | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-displaymessageform-member(1)) | **read item** | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-displaymessageformasync-member(1)) | **read item** | Compose<br>Read | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-displaynewappointmentform-member(1)) | **read item** | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-displaynewappointmentformasync-member(1)) | **read item** | Read | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-displaynewmessageform-member(1)) | **read item** | Read | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-displaynewmessageformasync-member(1)) | **read item** | Read | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(1)) | **read item** | Compose<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(2)) | **read item** | Compose<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getSelectedItemsAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-getselecteditemsasync-member(1)) | **read/write mailbox** | Compose<br>Read | [1.13](../requirement-set-1.13/outlook-requirement-set-1.13.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-getuseridentitytokenasync-member(1)) | **read item** | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [loadItemByIdAsync(itemId, [options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-loaditembyidasync-member(1)) | **read/write item** | Compose<br>Read | [Preview](outlook-requirement-set-preview.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-makeewsrequestasync-member(1)) | **read/write mailbox** | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) | **read item** | Compose<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## Events

You can subscribe to and unsubscribe from the following events using [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) and [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) respectively.

> [!IMPORTANT]
> Events are only available with task pane implementation.

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true) | Description | Minimum<br>requirement set |
|---|---|:---:|
|`ItemChanged`| A different Outlook item is selected for viewing while the task pane is pinned. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
|`SelectedItemsChanged`| One or more messages are selected or deselected. | [1.13](../requirement-set-1.13/outlook-requirement-set-1.13.md) |
