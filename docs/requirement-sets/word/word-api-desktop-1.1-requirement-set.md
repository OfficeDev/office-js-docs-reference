---
title: Word JavaScript API desktop-only requirement set 1.1
description: Details about the WordApiDesktop 1.1 requirement set.
ms.date: 11/18/2024
ms.topic: whats-new
ms.localizationpriority: medium
---

# Word JavaScript API desktop-only requirement set 1.1

The `WordApiDesktop` requirement set is a special requirement set that includes features that are only available for Word on Windows, on Mac, and on iPad. APIs in this requirement set are considered to be production APIs for the Word application on Windows, on Mac, and on iPad. They follow [Microsoft 365 developer support policies](/office/dev/add-ins/publish/maintain-breaking-changes). `WordApiDesktop` APIs are considered to be "preview" APIs for other platforms (web) and may not be supported by any of those platforms.

When APIs in the `WordApiDesktop` requirement set are supported across all platforms, they'll be added to the next released requirement set (`WordApi 1.[NEXT]`). Once that new requirement set is public, those APIs will also continue to be tagged in this `WordApiDesktop` requirement set. To learn more about platform-specific requirements in general, see [Understanding platform-specific requirement sets](https://aka.ms/PlatformSpecificReqtSets).

> [!IMPORTANT]
> `WordApiDesktop 1.1` is a desktop-only requirement set. It's a superset of the WordApi 1.8.

## Recommended usage

Because the `WordApiDesktop 1.1` APIs are only supported by Word on Windows, on Mac, and on iPad, your add-in should check if the requirement set is supported before calling these APIs. This avoids any attempt to use desktop-only APIs on an unsupported platform.

```js
if (Office.context.requirements.isSetSupported("WordApiDesktop", "1.1")) {
   // Any API exclusive to this WordApiDesktop requirement set.
}
```

Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check. This will enable your add-in's feature on other platforms. Be sure to test the feature on those platforms when making this change.

> [!IMPORTANT]
> Your manifest cannot specify `WordApiDesktop 1.1` as an activation requirement. It isn't a valid value to use in the [Set element](/javascript/api/manifest/set).

## API list

The following table lists the Word JavaScript APIs currently included in the `WordApiDesktop 1.1` requirement set. For a complete list of all Word JavaScript APIs (including `WordApiDesktop 1.1` APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-desktop-1.1&preserve-view=true).

[!INCLUDE[API table](../../includes/word-desktop-1.1.md)]

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word?view=word-js-desktop-1.1&preserve-view=true)
- [Word JavaScript preview APIs](word-preview-apis.md)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
- [Understanding platform-specific requirement sets](https://aka.ms/PlatformSpecificReqtSets)
