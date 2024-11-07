---
title: Dialog API requirement sets
description: Learn more about the Dialog API requirement sets.
ms.date: 11/07/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Dialog API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Office on the web | Office on Windows<br>(Microsoft 365 subscription) | Office on Windows<sup>1</sup><br>(retail perpetual) | Office on Windows<sup>1</sup><br>(volume-licensed perpetual) | Office on Mac | Office on iPad | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.2 | Supported | Version 2005 (Build 12827.20268) | Version 2005 (Build 12827.20268) | Office 2021: Version 2005 (Build 12827.20268) | Version 16.37 (20051002) | Version 16.37 | Not supported |
| DialogApi 1.1 | Supported | Version 1602 (Build 6741.0000) | Version 1602 (Build 6741.0000) | Office 2016<sup>2</sup> | Version 15.20 | Version 1.22 | Version 1608 (Build 7601.6800) |

> [!NOTE]
> <sup>1</sup>Users of perpetual versions of Office may not have accepted all patches and updates. If so, the DLL that Office uses to report its version in the UI may be greater than the versions listed here even if the updated DLLs needed to support DialogApi haven't been installed on the user's computer. To ensure that the needed patch is installed, the user must go to the [Office 2016 update list](/officeupdates/msp-files-office-2016), search for **osfclient-x-none**, and install the listed patch.
>
> <sup>2</sup> DialogApi 1.1 isn't supported in Outlook 2016 (volume-licensed perpetual).

## Outlook support

In Outlook, dialog features and their APIs were introduced in the following Mailbox requirement sets.

| Feature | API | Mailbox requirement set |
| :----- | :----- | :----- |
| Display a dialog | [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) | [1.4](../outlook/requirement-set-1.4/outlook-requirement-set-1.4.md) |
| Deliver a message from the dialog to the its parent page | [Office.context.ui.messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) | [1.4](../outlook/requirement-set-1.4/outlook-requirement-set-1.4.md) |
| Close a dialog | [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) | [1.5](../outlook/requirement-set-1.5/outlook-requirement-set-1.5.md) |
| Deliver a message from the host page to a dialog | [Office.dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) | [1.9](../outlook/requirement-set-1.9/outlook-requirement-set-1.9.md)\* |

> [!NOTE]
> \* Although classic Outlook on Mac doesn't support Mailbox requirement set 1.9, it does support the DialogApi 1.2 requirement set.

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## Dialog API 1.1 and 1.2

The Dialog API 1.1 is the first version of the API (see [Office.UI](/javascript/api/office/office.ui)). Requirement set 1.2 adds support for sending data from the parent page to the dialog box with the [Office.dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) method. For details about these APIs, see the [Dialog API](/javascript/api/office/office.ui) reference topic.

## See also

- [Use the Office dialog API in Office Add-ins](/office/dev/add-ins/develop/dialog-api-in-office-add-ins)
- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
