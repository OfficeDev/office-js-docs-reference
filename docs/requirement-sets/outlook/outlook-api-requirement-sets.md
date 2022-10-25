---
title: Outlook JavaScript API requirement sets
description: Learn more about the Outlook JavaScript API requirement sets.
ms.date: 10/20/2022
ms.prod: outlook
ms.localizationpriority: high
---

# Outlook JavaScript API requirement sets

Outlook add-ins declare what API versions they require in their manifest. The markup varies depending on whether you are using the [Teams manifest format (preview)](/office/dev/add-ins/develop/json-manifest-overview) or the [XML manifest format](/office/dev/add-ins/develop/add-in-manifests).

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

# [Teams Manifest (developer preview)](#tab/jsonmanifest)

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

---

All Outlook APIs belong to the `Mailbox` [requirement set](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements). The `Mailbox` requirement set has versions, and each new set of APIs that we release belongs to a higher version of the set. Not all Outlook clients support the newest set of APIs, but if an Outlook client declares support for a requirement set, generally it supports all of the APIs in that requirement set (check the documentation on a specific API or feature for any exceptions).

Setting a minimum requirement set version in the manifest controls which Outlook client the add-in will appear in. If a client doesn't support the minimum requirement set, it doesn't load the add-in. For example, if requirement set version 1.3 is specified, this means the add-in will not show up in any Outlook client that doesn't support at least 1.3.

> [!NOTE]
> To use APIs in any of the numbered requirement sets, you should reference the **production** library on the [Office.js content delivery network (CDN)](https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> For information about using preview APIs, see the [Using preview APIs](#using-preview-apis) section later in this article.

## Using APIs from later requirement sets

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

> [!IMPORTANT]
> There is currently a bug where `isSetSupported('Mailbox', '1.3')` erroneously returns `true` in Outlook on the web against Exchange 2013. To learn more about the supported combinations of requirement sets, Exchange servers, and Outlook clients, see [Requirement sets supported by Exchange servers and Outlook clients](#requirement-sets-supported-by-exchange-servers-and-outlook-clients).

Alternatively, developers can check for the existence of a newer API by using standard JavaScript technique.

```js
if (item.somePropertyOrMethod !== undefined) {
  // Use item.somePropertyOrMethod.
  item.somePropertyOrMethod;
}
```

No such checks are necessary for any APIs which are present in the requirement set version specified in the manifest.

## Choosing a minimum requirement set

Developers should use the earliest requirement set that contains the critical set of APIs for their scenario, without which the add-in won't work.

## Requirement sets supported by Exchange servers and Outlook clients

In this section, we note the range of requirement sets supported by Exchange server and Outlook clients. For details about server and client requirements for running Outlook add-ins, see [Outlook add-ins requirements](/office/dev/add-ins/outlook/add-in-requirements).

> [!IMPORTANT]
> If your target Exchange server and Outlook client support different requirement sets, then you may be restricted to the lower requirement set range. For example, if an add-in is running in Outlook 2016 on Mac (highest requirement set: 1.6) against Exchange 2013 (highest requirement set: 1.1), your add-in may be limited to requirement set 1.1.

### Exchange server support

The following servers support Outlook add-ins.

| Product | Major Exchange version | Supported API requirement sets |
|---|---|---|
| Exchange Online | Latest build | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md), [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md), [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)\* |
| Exchange on-premises | 2019 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2016 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2013 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md) |

> [!NOTE]
> \* [!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]

### Outlook client support

Add-ins are supported in Outlook on the following platforms.

| Platform | Major Office/Outlook version | Supported API requirement sets |
|---|---|---|
| Windows | - Microsoft 365 subscription<br>- retail perpetual Outlook 2016 and later | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup>, [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md)<sup>1</sup>, [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md)<sup>1</sup>, [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md)<sup>1</sup>, [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md)<sup>1</sup><br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>2</sup> |
|| volume-licensed perpetual Outlook 2021 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup>, [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md)<sup>1</sup><br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>2</sup> |
|| volume-licensed perpetual Outlook 2019 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| volume-licensed perpetual Outlook 2016 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
|| perpetual Outlook 2013 | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md)<sup>3</sup>, [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
| Mac | classic UI | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>2</sup> |
|| new UI<sup>4</sup> | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md), [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md), [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>2</sup> |
| iOS<sup>5 6</sup> | subscription | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md) |
| Android<sup>5 6</sup> | subscription | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md) |
| Web browser<sup>5 7</sup> | modern Outlook UI when connected to<br>Exchange Online: subscription, Outlook.com | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](requirement-set-1.10/outlook-requirement-set-1.10.md), [1.11](requirement-set-1.11/outlook-requirement-set-1.11.md), [1.12](requirement-set-1.12/outlook-requirement-set-1.12.md)<br>[IdentityAPI 1.3](../common/identity-api-requirement-sets.md)<sup>2</sup> |
|| classic Outlook UI when connected to<br>Exchange on-premises | [1.1](requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](requirement-set-1.6/outlook-requirement-set-1.6.md) |

> [!NOTE]
> <sup>1</sup> Version support for more recent requirement sets in Outlook on Windows with a Microsoft 365 subscription or a retail perpetual license as follows:
>
> - Support for **1.8** is available from Version 1910 (Build 12130.20272).
> - Support for **1.9** is available from Version 2008 (Build 13127.20296).
> - Support for **1.10** is available from Version 2104 (Build 13929.20296).
> - Support for **1.11** is available from Version 2110 (Build 14527.20226).
> - Support for **1.12** is available from Version 2206 (Build 15330.20196).
>
> For more details according to your version, see the update history page for [Office 2021](/officeupdates/update-history-office-2021) or [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) and how to [find your Office client version and update channel](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).
>
> <sup>2</sup> [!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]
>
> <sup>3</sup> Support for 1.3 in Outlook 2013 was added as part of the [December 8, 2015, update for Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349). Support for 1.4 in Outlook 2013 was added as part of the [September 13, 2016, update for Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280). Support for 1.4 in volume-licensed perpetual Outlook 2016 was added as part of the [July 3, 2018, update for Office 2016 (KB4022223)](https://support.microsoft.com/help/4022223).
>
> <sup>4</sup> Support for the new Mac UI is available from Outlook Version 16.38.506. For more information, see the [Add-in support in Outlook on new Mac UI](/office/dev/add-ins/outlook/compare-outlook-add-in-support-in-outlook-for-mac#add-in-support-in-outlook-on-new-mac-ui) section.
>
> <sup>5</sup> Add-ins aren't supported in Outlook on Android, on iOS, and modern mobile web with on-premises Exchange accounts. Certain iOS devices still support add-ins when using on-premises Exchange accounts with classic Outlook on the web. For information about supported devices, see [Requirements for running Office Add-ins](/office/dev/add-ins/concepts/requirements-for-running-office-add-ins#client-requirements-non-windows-smartphone-and-tablet).
>
> <sup>6</sup> Currently, there are additional considerations when designing and implementing add-ins for mobile clients. For example, the only supported mode is Message Read. For more details, see [code considerations when adding support for add-in commands for Outlook Mobile](/office/dev/add-ins/outlook/add-mobile-support#code-considerations).
>
> <sup>7</sup> Add-ins don't work in modern Outlook on the web on iPhone and Android smartphones. For information about supported devices, see [Requirements for running Office Add-ins](/office/dev/add-ins/concepts/requirements-for-running-office-add-ins#client-requirements-non-windows-smartphone-and-tablet).

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

## Using preview APIs

New Outlook JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired. To provide feedback about a preview API, please use the feedback mechanism at the end of the web page where the API is documented.

> [!NOTE]
> Preview APIs are subject to change and aren't intended for use in a production environment.

For more details about the preview APIs, see [Outlook API preview requirement set](preview-requirement-set/outlook-requirement-set-preview.md).
