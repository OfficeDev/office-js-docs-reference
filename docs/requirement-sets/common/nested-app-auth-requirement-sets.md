---
title: Nested app auth requirement sets
description: Nested app auth requirement set information for Office Add-ins.
ms.date: 11/22/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Nested app auth requirement set

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office Add-ins run across multiple versions of Office. The following table lists the nested app auth (NAA) requirement set, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Outlook on the web | Outlook on Windows<ul><li>Microsoft 365 subscription</li></ul> | Office on Windows<ul><li>retail perpetual</li><li>volume-licensed perpetual</li></ul> | Outlook on Mac | Outlook on iOS | Outlook on Android |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| NestedAppAuth 1.1 | Supported | Version 2409 (Build 18025.20000) | Not available | Version 16.89 (Build 24090815) | Build v4.2433.0 | Build v4.2433.0 |

> [!IMPORTANT]
>
> - Currently, the NestedAppAuth 1.1 requirement set is supported in Office on the web only for documents that are opened from Microsoft SharePoint Online and OneDrive.
> - In Outlook, the NestedAppAuth 1.1 requirement set isn't supported if the add-in is loaded in an Outlook.com or Gmail mailbox.

## Supported accounts and hosts

NAA supports both Microsoft Accounts and Microsoft Entra ID (work/school) identities. It doesn't support Azure Active Directory B2C for business-to-consumer identity management scenarios. The following table explains the current support by platform. Platforms listed as generally available (GA) are ready for production usage in your add-in.

| Application | Web        | Windows                                              | Mac        | iOS/iPad           | Android        |
|-------------|------------|------------------------------------------------------|------------|--------------------|----------------|
| Excel       | In preview | In preview                                           | In preview | In preview on iPad | Not applicable |
| Outlook     | GA         | GA in Current Channel and Monthly Enterprise Channel, Preview in Semi-Annual Channels | GA         | GA (iOS)           | GA             |
| PowerPoint  | In preview | In preview                                           | In preview | In preview on iPad | Not applicable |
| Word        | In preview | In preview                                           | In preview | In preview on iPad | Not applicable |

> [!IMPORTANT]
> To use NAA on platforms that are still in preview (Word, Excel, and PowerPoint), join the Microsoft 365 Insider Program (https://insider.microsoft365.com/join) and choose **Current Channel (Preview)**. Don't use NAA in production add-ins for any preview platforms. We invite you to try out NAA in test or development environments and welcome feedback on your experience through GitHub (see the **Feedback** section at the end of this page).

## Outlook and NestedAppAuth requirement set

To require the NestedAppAuth requirement set 1.1 in your Outlook add-in code, check if it's supported by calling `isSetSupported('NestedAppAuth', '1.1')`.
Declaring it in the Outlook add-in's manifest isn't supported. You can also determine if the API is supported by checking that it's not `undefined`.
For further details, see [Using APIs from later requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#using-apis-from-later-requirement-sets).

## Office versions and build numbers

To find out more about versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
- [Enable SSO in an Office Add-in using nested app authentication](/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in)
