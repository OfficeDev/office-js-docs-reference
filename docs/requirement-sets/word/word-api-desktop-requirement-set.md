---
title: Word JavaScript API desktop-only requirement set
description: Details about the WordApiHiddenDocument requirement set.
ms.date: 06/24/2022
ms.prod: word
ms.localizationpriority: medium
---

# Word JavaScript API desktop-only requirement set

The `WordApiHiddenDocument` requirement set is a special requirement set that includes features that are only available for Word on Windows and on Mac. APIs in this requirement set are considered to be production APIs for the Word application on Windows and on Mac. They follow [Microsoft 365 developer support policies](/office/dev/add-ins/publish/maintain-breaking-changes). `WordApiHiddenDocument` APIs are considered to be "preview" APIs for other platforms (web, iPad) and may not be supported by any of those platforms.

When APIs in the `WordApiHiddenDocument` requirement set are supported across all platforms, they will be added to a subsequent released requirement set (`WordApi 1.[FUTURE]`). Once that new requirement is public, those APIs will be removed from `WordApiHiddenDocument`. Think of this as a similar promotion process to an API moving from preview to release.

> [!IMPORTANT]
>
> - `WordApiHiddenDocument` is a superset of the WordApi 1.3 requirement set.
> - `WordApiHiddenDocument 1.3` is currently the only desktop-only requirement set.

## Recommended usage

Because `WordApiHiddenDocument` APIs are only supported by Word on Windows and on Mac, your add-in should check if the requirement set is supported before calling these APIs. This avoids any attempt to use desktop-only APIs on an unsupported platform.

```js
if (Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.3")) {
   // Any API exclusive to the desktop-only requirement set.
}
```

Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check. This will enable your add-in's feature on other platforms. Be sure to test the feature on those platforms when making this change.

> [!IMPORTANT]
> Your manifest cannot specify `WordApiHiddenDocument 1.3` as an activation requirement. It's not a valid value to use in the [Set element](/javascript/api/manifest/set).

## API list

The following table lists the Word JavaScript APIs currently included in the `WordApiHiddenDocument` requirement set. For a complete list of all Word JavaScript APIs (including `WordApiHiddenDocument` APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-desktop&preserve-view=true).

[!INCLUDE[API table](../../includes/word-desktop.md)]

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word?view=word-js-desktop&preserve-view=true)
- [Word JavaScript preview APIs](word-preview-apis.md)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
