---
title: Outlook JavaScript API requirement sets
description: Learn more about the Outlook JavaScript API requirement sets.
ms.date: 10/22/2024
ms.topic: overview
ms.localizationpriority: high
---

# Outlook JavaScript API requirement sets

Outlook add-ins declare what API versions they require in their manifest. The markup varies depending on whether you're using the [add-in only manifest format](/office/dev/add-ins/develop/xml-manifest-overview) or the [unified manifest for Microsoft 365](/office/dev/add-ins/develop/unified-manifest-overview).

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

The API version is specified by the "extensions.requirements.capabilities" property. Set the "capabilities.name" property to "Mailbox" and the "capabilities.minVersion" property to the minimum API requirement set that supports the add-in's scenarios.

For example, the following manifest snippet indicates a minimum requirement set of 1.1.

```json
"extensions": [
{
  "requirements": {
    "capabilities": [
      {
        "name": "Mailbox", "minVersion": "1.1"
      }
    ]
  },
  ...
}
```

# [XML Manifest](#tab/xmlmanifest)

The API version is specified by the [Requirements](/javascript/api/manifest/requirements) element. Outlook add-ins always include a [Set](/javascript/api/manifest/set) element with a `Name` attribute set to `Mailbox` and a `MinVersion` attribute set to the minimum API requirement set that supports the add-in's scenarios.

For example, the following manifest snippet indicates a minimum requirement set of 1.1.

```xml
<Requirements>
  <Sets>
    <Set Name="Mailbox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

---

All Outlook APIs belong to the `Mailbox` [requirement set](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements). The `Mailbox` requirement set has versions, and each new set of APIs that we release belongs to a higher version of the set. Not all Outlook clients support the newest set of APIs, but if an Outlook client declares support for a requirement set, generally it supports all of the APIs in that requirement set (check the documentation on a specific API or feature for any exceptions).

Setting a minimum requirement set version in the manifest controls in which Outlook client the add-in will appear. If a client doesn't support the minimum requirement set, it doesn't load the add-in. For example, if requirement set version 1.3 is specified, this means the add-in will not show up in any Outlook client that doesn't support at least 1.3.

> [!NOTE]
> Although Outlook on Android and on iOS support up to requirement set 1.5, your mobile add-in can now implement some APIs from later requirement sets. For more information on which APIs are supported in Outlook mobile, see [Outlook JavaScript APIs supported in Outlook on mobile devices](/office/dev/add-ins/outlook/outlook-mobile-apis).

## Use APIs from later requirement sets

Setting a requirement set doesn't limit the available APIs that the add-in can use. For example, if the add-in specifies requirement set "Mailbox 1.1", but it's running in an Outlook client which supports "Mailbox 1.3", the add-in can use APIs from requirement set "Mailbox 1.3".

To use a newer API, developers can check if a particular application supports the requirement set by doing the following:

```js
if (Office.context.requirements.isSetSupported('Mailbox', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

Alternatively, developers can check for the existence of a newer API by using standard JavaScript technique.

```js
if (item.somePropertyOrMethod !== undefined) {
  // Use item.somePropertyOrMethod.
  item.somePropertyOrMethod;
}
```

No such checks are necessary for any APIs which are present in the requirement set version specified in the manifest.

## Choose a minimum requirement set

Developers should use the earliest requirement set that contains the critical set of APIs for their scenario, without which the add-in won't work.

## Requirement sets supported by Exchange servers and Outlook clients

In this section, we note the range of requirement sets supported by Exchange server and Outlook clients. For details about server and client requirements for running Outlook add-ins, see [Outlook add-ins requirements](/office/dev/add-ins/outlook/add-in-requirements).

> [!IMPORTANT]
> If your target Exchange server and Outlook client support different requirement sets, then you may be restricted to the lower requirement set range. For example, if an add-in is running in Outlook 2019 on Windows (highest requirement set: 1.6) against Exchange 2016 (highest requirement set: 1.5), your add-in may be limited to requirement set 1.5.

### Exchange server support

The following servers support Outlook add-ins.

| Product | Major Exchange version | Supported API requirement sets |
|---|---|---|
| Exchange Online | Latest build | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md), [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md), [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md), [1.13](requirement-set-1.13/outlook-requirement-set-1.13.md), [1.14](requirement-set-1.14/outlook-requirement-set-1.14.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>1</sup> |
| Exchange on-premises<sup>2</sup> | 2019 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2016 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md) |

> [!NOTE]
> <sup>1</sup> [!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]
>
> <sup>2</sup> Even if an add-in implements features from requirement sets not supported in an Exchange on-premises environment, it can still be added to an Outlook client as long as the requirement set specified in its manifest aligns with those supported by Exchange on-premises. However, an implemented feature will only work if the Outlook client in which the add-in is installed supports the minimum requirement set needed by a feature. To determine the requirement sets supported by varying Outlook clients, see [Outlook client support](#outlook-client-support). We recommend supplementing this with the documentation on the specific feature for any exceptions.

### Outlook client support

Add-ins are supported in Outlook on the following platforms.

| Platform | Major Office/Outlook version | Supported API requirement sets |
|---|---|---|
| Web browser<sup>1 2</sup> | modern Outlook UI when connected to<br>Exchange Online: subscription, Outlook.com | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md), [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md), [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md), [1.13](requirement-set-1.13/outlook-requirement-set-1.13.md), [1.14](requirement-set-1.14/outlook-requirement-set-1.14.md)<br>[DevicePermissionService 1.1](../common/device-permission-service-requirement-sets.md)<br>[DialogAPI 1.1](../common/dialog-api-requirement-sets.md)<br>[DialogAPI 1.2](../common/dialog-api-requirement-sets.md)<br>[DialogOrigin 1.1](../common/dialog-origin-requirement-sets.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>3</sup><br>[NestedAppAuth 1.1](../common/nested-app-auth-requirement-sets.md) |
|| classic Outlook UI when connected to<br>Exchange on-premises | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md) |
| Windows | [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md), [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md), [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md), [1.13](requirement-set-1.13/outlook-requirement-set-1.13.md), [1.14](requirement-set-1.14/outlook-requirement-set-1.14.md)<br>[DevicePermissionService 1.1](../common/device-permission-service-requirement-sets.md)<br>[DialogAPI 1.1](../common/dialog-api-requirement-sets.md)<br>[DialogAPI 1.2](../common/dialog-api-requirement-sets.md)<br>[DialogOrigin 1.1](../common/dialog-origin-requirement-sets.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>3</sup><br>[NestedAppAuth 1.1](../common/nested-app-auth-requirement-sets.md) |
|| Microsoft 365 subscription | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>4</sup>, [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md)<sup>4</sup>, [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md)<sup>4</sup>, [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md)<sup>4</sup>, [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md)<sup>4</sup>, [1.13](requirement-set-1.13/outlook-requirement-set-1.13.md)<sup>4</sup>, [1.14](requirement-set-1.14/outlook-requirement-set-1.14.md)<sup>4</sup><br>[DialogAPI 1.1](../common/dialog-api-requirement-sets.md)<br>[DialogAPI 1.2](../common/dialog-api-requirement-sets.md)<br>[DialogOrigin 1.1](../common/dialog-origin-requirement-sets.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>3</sup><br>[NestedAppAuth 1.1](../common/nested-app-auth-requirement-sets.md)<br>[OpenBrowserWindowApi 1.1](../common/open-browser-window-api-requirement-sets.md) |
|| retail perpetual Outlook 2016 and later | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>4</sup>, [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md)<sup>4</sup>, [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md)<sup>4</sup>, [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md)<sup>4</sup>, [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md)<sup>4</sup>, [1.13](requirement-set-1.13/outlook-requirement-set-1.13.md)<sup>4</sup>, [1.14](requirement-set-1.14/outlook-requirement-set-1.14.md)<sup>4</sup><br>[DialogAPI 1.1](../common/dialog-api-requirement-sets.md)<br>[DialogAPI 1.2](../common/dialog-api-requirement-sets.md)<br>[DialogOrigin 1.1](../common/dialog-origin-requirement-sets.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>3</sup><br>[OpenBrowserWindowApi 1.1](../common/open-browser-window-api-requirement-sets.md) |
|| volume-licensed perpetual Outlook 2024 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md), [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md), [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md), [1.13](requirement-set-1.13/outlook-requirement-set-1.13.md), [1.14](requirement-set-1.14/outlook-requirement-set-1.14.md)<br>[DialogAPI 1.1](../common/dialog-api-requirement-sets.md)<br>[DialogAPI 1.2](../common/dialog-api-requirement-sets.md)<br>[DialogOrigin 1.1](../common/dialog-origin-requirement-sets.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>3</sup><br>[OpenBrowserWindowApi 1.1](../common/open-browser-window-api-requirement-sets.md) |
|| volume-licensed perpetual Outlook 2021 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md)<br>[DialogAPI 1.1](../common/dialog-api-requirement-sets.md)<br>[DialogAPI 1.2](../common/dialog-api-requirement-sets.md)<br>[DialogOrigin 1.1](../common/dialog-origin-requirement-sets.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>3</sup><br>[OpenBrowserWindowApi 1.1](../common/open-browser-window-api-requirement-sets.md) |
|| volume-licensed perpetual Outlook 2019 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| volume-licensed perpetual Outlook 2016 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>5</sup> |
| Mac | new UI<sup>6</sup> | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md), [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md), [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md), [1.13](requirement-set-1.13/outlook-requirement-set-1.13.md)<br>[DialogAPI 1.1](../common/dialog-api-requirement-sets.md)<br>[DialogAPI 1.2](../common/dialog-api-requirement-sets.md)<br>[DialogOrigin 1.1](../common/dialog-origin-requirement-sets.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>3</sup><br>[NestedAppAuth 1.1](../common/nested-app-auth-requirement-sets.md)<br>[OpenBrowserWindowApi 1.1](../common/open-browser-window-api-requirement-sets.md) |
|| classic UI | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md)<br>[DialogAPI 1.1](../common/dialog-api-requirement-sets.md)<br>[DialogAPI 1.2](../common/dialog-api-requirement-sets.md)<sup>7</sup><br>[DialogOrigin 1.1](../common/dialog-origin-requirement-sets.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>3</sup><br>[NestedAppAuth 1.1](../common/nested-app-auth-requirement-sets.md)<br>[OpenBrowserWindowApi 1.1](../common/open-browser-window-api-requirement-sets.md) |
| Android<sup>1 8</sup> | subscription | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md)<br>[NestedAppAuth 1.1](../common/nested-app-auth-requirement-sets.md) |
| iOS<sup>1 8</sup> | subscription | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md)<br>[NestedAppAuth 1.1](../common/nested-app-auth-requirement-sets.md) |

> [!NOTE]
> <sup>1</sup> Add-ins aren't supported in Outlook on Android, on iOS, and modern mobile web with on-premises Exchange accounts. Certain iOS devices still support add-ins when using on-premises Exchange accounts with classic Outlook on the web. For information about supported devices, see [Requirements for running Office Add-ins](/office/dev/add-ins/concepts/requirements-for-running-office-add-ins#client-requirements-non-windows-smartphone-and-tablet).
>
> <sup>2</sup> Add-ins don't work in modern Outlook on the web on iPhone and Android smartphones. For information about supported devices, see [Requirements for running Office Add-ins](/office/dev/add-ins/concepts/requirements-for-running-office-add-ins#client-requirements-non-windows-smartphone-and-tablet).
>
> <sup>3</sup> [!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]
>
> <sup>4</sup> To learn about the minimum supported versions for recent requirement sets in classic Outlook on Windows with a Microsoft 365 subscription or a retail perpetual license, see [Version support for requirement sets in classic Outlook on Windows](#version-support-for-requirement-sets-in-classic-outlook-on-windows).
>
> <sup>5</sup> Support for 1.4 in volume-licensed perpetual Outlook 2016 was added as part of the [July 3, 2018, update for Office 2016 (KB4022223)](https://support.microsoft.com/help/4022223).
>
> <sup>6</sup> Support for the new Mac UI is available from Outlook Version 16.38.506. For more information, see the [Add-in support in Outlook on new Mac UI](/office/dev/add-ins/outlook/compare-outlook-add-in-support-in-outlook-for-mac#add-in-support-in-outlook-on-new-mac-ui) section.
>
> <sup>7</sup> Although classic Outlook on Mac doesn't support Mailbox requirement set 1.9, it does support the DialogApi 1.2 requirement set. For information on the minimum supported version and build, see [Dialog API requirement sets](../common/dialog-api-requirement-sets.md).
>
> <sup>8</sup> Currently, there are additional considerations when designing and implementing add-ins for mobile clients. For more details, see [code considerations when adding support for add-in commands in Outlook on mobile devices](/office/dev/add-ins/outlook/add-mobile-support#code-considerations). Although Outlook on Android and on iOS support up to requirement set 1.5, your mobile add-in can now implement some APIs from later requirement sets. For more information on which APIs are supported in Outlook mobile, see [Outlook JavaScript APIs supported in Outlook on mobile devices](/office/dev/add-ins/outlook/outlook-mobile-apis).

> [!TIP]
> You can distinguish between classic and modern Outlook in a web browser by checking your mailbox toolbar.
>
> **modern**
>
> ![The modern Outlook toolbar.](/office/dev/add-ins/images/outlook-on-the-web-new-toolbar.png)
>
> **classic**
>
> ![The classic Outlook toolbar.](/office/dev/add-ins/images/outlook-on-the-web-classic-toolbar.png)

#### Version support for requirement sets in classic Outlook on Windows

The following table lists version support for more recent Mailbox requirement sets in classic Outlook on Windows with a Microsoft 365 subscription or a retail perpetual license.

| Requirement set | Version |
| ---- | ---- |
| 1.8 | Version 1910 (Build 12130.20272) |
| 1.9 | Version 2008 (Build 13127.20296) |
| 1.10 | Version 2104 (Build 13929.20296) |
| 1.11 | Version 2110 (Build 14527.20226) |
| 1.12 | Version 2206 (Build 15330.20196) |
| 1.13 | Version 2304 (Build 16327.20248) |
| 1.14 | Version 2404 (Build 17530.15000) |

For more details about your client version, see the update history page for [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) or [Office 2024](/officeupdates/update-history-office-2024) and how to [find your Office client version and update channel](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

## Reference the Office JavaScript API production library

To use APIs in any of the numbered requirement sets, you should reference the **production** library on the [Office.js content delivery network (CDN)](https://appsforoffice.microsoft.com/lib/1/hosted/office.js). For information on how to use preview APIs, see [Test preview APIs](#test-preview-apis).

## Test preview APIs

New Outlook JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired. To provide feedback about a preview API, please use the feedback mechanism at the end of the web page where the API is documented.

> [!NOTE]
> Preview APIs are subject to change and aren't intended for use in a production environment.

For more details about the preview APIs, see [Outlook API preview requirement set](preview-requirement-set/outlook-requirement-set-preview.md).
