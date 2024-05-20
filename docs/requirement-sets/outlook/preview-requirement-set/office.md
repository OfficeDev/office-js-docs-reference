---
title: Office namespace - preview requirement set
description: Office namespace members available for Outlook add-ins using Mailbox API preview requirement set.
ms.date: 05/20/2024
ms.localizationpriority: medium
---

# Office (Mailbox preview requirement set)

The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office?view=outlook-js-preview&preserve-view=true).

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points)| Compose or Read|

## Properties

| Property | Modes | Return type | Minimum<br>requirement set |
|---|---|---|:---:|
| [context](office.context.md) | Compose<br>Read | [Context](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## Enumerations

| Enumeration | Modes | Return type | Minimum<br>requirement set |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Compose<br>Read | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [HostType](#hosttype-string) | Compose<br>Read | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [PlatformType](#platformtype-string) | Compose<br>Read | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## Namespaces

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.

## Enumeration details

#### AsyncResultStatus: String

Specifies the result of an asynchronous call.

##### Type

*   String

##### Properties

|Name| Type| Description|
|---|---|---|
|`Succeeded`| String|The call succeeded.|
|`Failed`| String|The call failed.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points)| Compose or Read|

<br>

---
---

#### CoercionType: String

Specifies how to coerce data returned or set by the invoked method.

##### Type

*   String

##### Properties

|Name| Type| Description|
|---|---|---|
|`Html`| String|Requests the data be returned in HTML format.|
|`Text`| String|Requests the data be returned in text format.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points)| Compose or Read|

<br>

---
---

#### EventType: String

Specifies the event associated with an event handler.

##### Type

*   String

##### Properties

| Name | Type | Description | Minimum requirement set |
|---|---|---|:---:|
|`AppointmentTimeChanged`| String | The date or time of the selected appointment or series has changed. | 1.7 |
|`AttachmentsChanged`| String | An attachment has been added to or removed from the item. | 1.8 |
|`EnhancedLocationsChanged`| String | The location of the selected appointment has changed. | 1.8 |
|`InfobarClicked`| String | An action on a notification message is selected. | 1.10 |
|`ItemChanged`| String | A different Outlook item is selected for viewing while the task pane is pinned. | 1.5 |
|`OfficeThemeChanged`| String | The Office theme on the mailbox has changed. | 1.14 |
|`RecipientsChanged`| String | The recipient list of the selected item or appointment location has changed. | 1.7 |
|`RecurrenceChanged`| String | The recurrence pattern of the selected series has changed. | 1.7 |
|`SelectedItemsChanged`| String | One or more messages are selected or deselected. | 1.13 |
|`SensitivityLabelChanged`| String | The sensitivity label of a message or appointment has changed. | 1.13 |
|`SpamReporting`| String | An unsolicited message has been reported in Outlook. | 1.14 |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../outlook-api-requirement-sets.md)| 1.5 |
|[Applicable Outlook mode](/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points)| Compose or Read|

<br>

---
---

#### HostType: String

Specifies the host Office application in which the add-in is running.

##### Type

*   String

##### Properties

| Name | Type | Description | Minimum requirement set |
|---|---|---|:---:|
|`Outlook`| String | The Office host is Microsoft Outlook. | 1.5 |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../outlook-api-requirement-sets.md)| 1.5 |
|[Applicable Outlook mode](/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points)| Compose or Read |

<br>

---
---

#### PlatformType: String

Specifies the OS or other platform on which the Office host application is running.

##### Type

*   String

##### Properties

| Name | Type | Description | Minimum requirement set |
|---|---|---|:---:|
|`Android`| String | The platform is an Android device. | 1.5 |
|`iOS`| String | The platform is an iOS device. | 1.5 |
|`Mac`| String | The platform is Mac. | 1.5 |
|`OfficeOnline`| String | The platform is Office on the web (in a browser). | 1.5 |
|`PC`| String | The platform is PC (Windows). | 1.5 |
|`Universal`| String | The platform is WinRT. | 1.5 |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../outlook-api-requirement-sets.md)| 1.5 |
|[Applicable Outlook mode](/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points)| Compose or Read |

<br>

---
---

#### SourceProperty: String

Specifies the source of the data returned by the invoked method.

##### Type

*   String

##### Properties

|Name| Type| Description|
|---|---|---|
|`Body`| String|The source of the data is from the body of a message.|
|`Subject`| String|The source of the data is from the subject of a message.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points)| Compose or Read|
