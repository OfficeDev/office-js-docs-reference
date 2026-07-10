---
title: Outlook add-in API requirement set 1.16
description: Lists the APIs introduced in Mailbox requirement set 1.16 for Outlook add-ins.
ms.date: 06/23/2026
ms.topic: whats-new
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.16

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

## What's new in 1.16

Mailbox requirement set 1.16 includes all of the features of [requirement set 1.15](outlook-requirement-set-1-15.md). It added the following features.

- Added an event and objects to support [decrypting a message and its attachments](/office/dev/add-ins/outlook/encryption-decryption).
- Extended support for the `contentId` property to get the content identifier of an inline attachment.
- Added a method to check if Exchange Web Services (EWS) tokens are supported in an organization.
- Updated the Recipients APIs to increase the maximum number of recipients you can retrieve from a target field.
- Increased the [SessionData](/javascript/api/outlook/office.sessiondata) object limit to 2,621,440 characters.

## API list

The following table lists the APIs introduced in Mailbox requirement set 1.16. To view API reference documentation for all APIs supported by Mailbox requirement set 1.16 or earlier, see [Outlook APIs](/javascript/api/outlook?view=outlook-js-1.16&preserve-view=true).

[!INCLUDE [outlook-1_16](../../includes/outlook-1_16.md)]

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
