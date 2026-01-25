---
title: Outlook add-in API preview requirement set
description: Features and APIs that are currently in preview for Outlook add-ins.
ms.date: 12/18/2025
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

The preview requirement set includes all of the features of [requirement set 1.15](outlook-requirement-set-1.15.md).

> [!IMPORTANT]
> This documentation is for a **preview** [requirement set](../outlook-api-requirement-sets.md). This requirement set isn't fully implemented yet, and clients won't accurately report support for it. You shouldn't specify this requirement set in your add-in manifest.

## Features in preview

The following features are in preview.

- Added an event and objects to support [decrypting a message and its attachments](/office/dev/add-ins/outlook/encryption-decryption).
- Extended support for the `contentId` property to get the content identifier of an inline attachment in classic Outlook on Windows.
- Added a method to check if Exchange Web Services (EWS) tokens are supported in an organization.
- Updated the Recipients APIs to increase the maximum number of recipients in a target field to 1,000.
- Increased the [SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-preview&preserve-view=true) object limit to 2,621,440 characters.
- Extended support for the `errorMessageMarkdown` property of the `event.completed` method to [Smart Alerts](/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events) add-ins in Outlook on Mac.
- Added a property to get or set whether an appointment is an all-day event.
- Added events to activate an [event-based add-in](/office/dev/add-ins/develop/event-based-activation) on a message in read mode when it contains certain attachment types or custom internet headers.
- Added a property and objects to temporarily set the body or subject displayed in read mode.

## API list

The following table lists the Outlook JavaScript APIs currently in preview. To view API reference documentation for all Outlook JavaScript APIs (including preview APIs and previously released APIs), see [Outlook APIs](/javascript/api/outlook?view=outlook-js-preview&preserve-view=true).

[!INCLUDE [outlook-preview](../../../includes/outlook-preview.md)]

## See also

- [Outlook add-ins](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Requirement sets and supported clients](../outlook-api-requirement-sets.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Build your first Outlook add-in](/office/dev/add-ins/quickstarts/outlook-quickstart)
