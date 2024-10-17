---
title: Outlook add-in API requirement set 1.13
description: Requirement set 1.13 for Outlook add-in API.
ms.date: 05/19/2023
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.13

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.13?

Requirement set 1.13 includes all of the features of [requirement set 1.12](../requirement-set-1.12/outlook-requirement-set-1.12.md). It added the following features.

- Added support to activate an add-in without the Reading Pane enabled or a message selected.
- Added support to manage the delivery data and time of a message.
- Added new events for [event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events).
- Added the item multi-select feature.
- Added the prepend-on-send feature.
- Added the sensitivity label feature.
- Added support for shared mailbox scenarios.

### Change log

- Added the [SupportsNoItemContext](/javascript/api/manifest/action?view=outlook-js-1.13&preserve-view=true#supportsnoitemcontext) add-in only manifest element: Allows task pane add-ins to activate without the Reading Pane enabled or a message selected.
- Added [Office.context.mailbox.item.delayDeliveryTime](office.context.mailbox.item.md#properties): Adds a property that provides the object to manage the delivery date and time of a message in compose mode.
- Added [Office.DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime?view=outlook-js-1.13&preserve-view=true): Adds an object to manage the delivery date and time of a message in compose mode.
- Added new events for [event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events): Adds support for the following events.
  - `OnMessageFromChanged`
  - `OnAppointmentFromChanged`
  - `OnSensitivityLabelChanged`
- Added [Office.context.mailbox.getSelectedItemsAsync](office.context.mailbox.md#methods): Adds a method to retrieve currently selected messages.
- Added [Office.EventType.SelectedItemsChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.13&preserve-view=true): Adds a new event to `Mailbox`. This event occurs when one or more messages are selected or deselected.
- Added [Office.context.mailbox.item.body.prependOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.13&preserve-view=true#outlook-office-body-prependonsendasync-member(1)): Adds a method to prepend content to the beginning of a message or appointment body when the mail item is sent.
- Added [Office.context.sensitivityLabelsCatalog](/javascript/api/office/office.context?view=outlook-js-1.13&preserve-view=true#office-office-context-sensitivitylabelscatalog-member): Adds a property that provides the object to check the status of the catalog of sensitivity labels and retrieve all available sensitivity labels if the catalog is enabled.
- Added [Office.context.mailbox.item.sensitivityLabel](office.context.mailbox.item.md#properties): Adds a property that provides the object to get or set the sensitivity label of a message or appointment in compose mode.
- Added [Office.EventType.SensitivityLabelChanged](/javascript/api/office/office.eventtype?view=outlook-js-1.13&preserve-view=true): Adds a new event to `Item`. This event occurs when the sensitivity label of a message or appointment is changed.
- Added [Office.SensitivityLabelChangedEventArgs](/javascript/api/outlook/office.sensitivitylabelchangedeventargs?view=outlook-js-1.13&preserve-view=true): Adds an object that provides the change status of the sensitivity label applied to a message or appointment in compose mode.
- Added [Office.SensitivityLabelsCatalog](/javascript/api/outlook/office.sensitivitylabelscatalog?view=outlook-js-1.13&preserve-view=true): Adds an object that represents the catalog of sensitivity labels in Outlook.
- Added [Office.SensitivityLabel](/javascript/api/outlook/office.sensitivitylabel?view=outlook-js-1.13&preserve-view=true): Adds an object that represents the sensitivity label of a message or appointment in compose mode.
- Added [Office.SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails?view=outlook-js-1.13&preserve-view=true): Adds an object that represents the properties of a sensitivity label.
- Modified [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#methods): Adds support for shared mailbox scenarios. This method gets an object that represents the shared properties of a message or appointment.
- Modified [Office.SharedProperties](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.13&preserve-view=true): Adds support for shared mailbox scenarios. This object represents the properties of a message or appointment in a shared folder or shared mailbox.
- Modified the [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders?view=outlook-js-1.13&preserve-view=true) add-in only manifest element: Adds support for shared mailbox scenarios. This element defines whether the add-in is available in shared folder and shared mailbox scenarios.

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
