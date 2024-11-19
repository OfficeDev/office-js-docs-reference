---
title: Word JavaScript API online-only requirement set
description: Details about the WordApiOnline requirement set.
ms.date: 08/29/2024
ms.topic: whats-new
ms.localizationpriority: medium
---

# Word JavaScript API online-only requirement set

The `WordApiOnline` requirement set is a special requirement set that includes features that are only available for Word on the web. APIs in this requirement set are considered to be production APIs for the Word application on the web. They follow [Microsoft 365 developer support policies](/office/dev/add-ins/publish/maintain-breaking-changes). `WordApiOnline` APIs are considered to be "preview" APIs for other platforms (Windows, Mac, iPad) and may not be supported by any of those platforms.

When APIs in the `WordApiOnline` requirement set are supported across all platforms, they will be added to the next released requirement set (`WordApi 1.[NEXT]`). Once that new requirement set is public, those APIs will be removed from `WordApiOnline`. Think of this as a similar promotion process to an API moving from preview to release.

> [!IMPORTANT]
> `WordApiOnline` is a superset of the latest numbered requirement set.

> [!IMPORTANT]
> `WordApiOnline 1.1` is the only version of the online-only APIs. This is because Word on the web will always have a single version available to users that is the latest version.

The following table provides a concise summary of the APIs, while the subsequent [API list](#api-list) table gives a detailed list of the current `WordApiOnline` APIs.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| *None* |||

## Recommended usage

Because `WordApiOnline` APIs are only supported by Word on the web, your add-in should check if the requirement set is supported before calling these APIs. This avoids any attempt to use online-only APIs on an unsupported platform.

```js
if (Office.context.requirements.isSetSupported("WordApiOnline", "1.1")) {
   // Any API exclusive to the WordApiOnline requirement set.
}
```

Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check. This will enable your add-in's feature on other platforms. Be sure to test the feature on those platforms when making this change.

> [!IMPORTANT]
> Your manifest cannot specify `WordApiOnline 1.1` as an activation requirement. It isn't a valid value to use in the [Set element](/javascript/api/manifest/set).

## API list

The following table lists the Word JavaScript APIs currently included in the `WordApiOnline` requirement set. For a complete list of all Word JavaScript APIs (including `WordApiOnline` APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-online&preserve-view=true).

[!INCLUDE[API table](../../includes/word-online.md)]

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word?view=word-js-online&preserve-view=true)
- [Word JavaScript preview APIs](word-preview-apis.md)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
- [Understanding platform-specific requirement sets](https://aka.ms/PlatformSpecificReqtSets)
