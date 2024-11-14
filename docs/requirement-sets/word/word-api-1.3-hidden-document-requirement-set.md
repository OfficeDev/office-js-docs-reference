---
title: Word JavaScript API Hidden Document requirement set 1.3
description: Details about the WordApiHiddenDocument 1.3 requirement set.
ms.date: 08/29/2024
ms.localizationpriority: medium
---

# Word JavaScript API Hidden Document requirement set 1.3

The `WordApiHiddenDocument 1.3` requirement set is a special requirement set that includes features that are only available for Word on Windows and on Mac. APIs in this requirement set are considered to be production APIs for the Word application on Windows and on Mac. They follow [Microsoft 365 developer support policies](/office/dev/add-ins/publish/maintain-breaking-changes). `WordApiHiddenDocument` APIs are considered to be "preview" APIs for other platforms (web, iPad) and may not be supported by any of those platforms.

> [!IMPORTANT]
> `WordApiHiddenDocument 1.3` is a superset of the WordApi 1.3 requirement set and is a desktop-only requirement set.

## Recommended usage

Because `WordApiHiddenDocument` APIs are only supported by Word on Windows and on Mac, your add-in should check if the requirement set is supported before calling these APIs. This avoids any attempt to use desktop-only APIs on an unsupported platform.

```js
if (Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.3")) {
   // Any API exclusive to this desktop-only requirement set.
}
```

Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check. This will enable your add-in's feature on other platforms. Be sure to test the feature on those platforms when making this change.

> [!IMPORTANT]
> Your manifest cannot specify `WordApiHiddenDocument 1.3` as an activation requirement. It's not a valid value to use in the [Set element](../../manifest/set.md).

## API list

The following table lists the Word JavaScript APIs currently included in the `WordApiHiddenDocument 1.3` requirement set. For a complete list of all Word JavaScript APIs (including `WordApiHiddenDocument 1.3` APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-1.3-hidden-document&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[body](/javascript/api/word/word.documentcreated#word-word-documentcreated-body-member)|Gets the body object of the document.|
||[contentControls](/javascript/api/word/word.documentcreated#word-word-documentcreated-contentcontrols-member)|Gets the collection of content control objects in the document.|
||[properties](/javascript/api/word/word.documentcreated#word-word-documentcreated-properties-member)|Gets the properties of the document.|
||[save()](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|Saves the document.|
||[saved](/javascript/api/word/word.documentcreated#word-word-documentcreated-saved-member)|Indicates whether the changes in the document have been saved.|
||[sections](/javascript/api/word/word.documentcreated#word-word-documentcreated-sections-member)|Gets the collection of section objects in the document.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word?view=word-js-1.3-hidden-document&preserve-view=true)
- [Word JavaScript preview APIs](word-preview-apis.md)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
- [Understanding platform-specific requirement sets](https://aka.ms/PlatformSpecificReqtSets)
