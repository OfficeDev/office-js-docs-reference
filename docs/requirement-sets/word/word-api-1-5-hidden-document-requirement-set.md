---
title: Word JavaScript API Hidden Document requirement set 1.5
description: Details about the WordApiHiddenDocument 1.5 requirement set.
ms.date: 08/29/2024
ms.localizationpriority: medium
---

# Word JavaScript API Hidden Document requirement set 1.5

The `WordApiHiddenDocument 1.5` requirement set is a special requirement set that includes features that are only available for Word on Windows and on Mac. APIs in this requirement set are considered to be production APIs for the Word application on Windows and on Mac. They follow [Microsoft 365 developer support policies](/office/dev/add-ins/publish/maintain-breaking-changes). `WordApiHiddenDocument` APIs are considered to be "preview" APIs for other platforms (web, iPad) and may not be supported by any of those platforms.

> [!IMPORTANT]
> `WordApiHiddenDocument 1.5` is a superset of the WordApi 1.5 and WordApiHiddenDocument 1.4 requirement sets, and is a desktop-only requirement set.

## Recommended usage

Because `WordApiHiddenDocument` APIs are only supported by Word on Windows and on Mac, your add-in should check if the requirement set is supported before calling these APIs. This avoids any attempt to use desktop-only APIs on an unsupported platform.

```js
if (Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.5")) {
   // Any API exclusive to this desktop-only requirement set.
}
```

Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check. This will enable your add-in's feature on other platforms. Be sure to test the feature on those platforms when making this change.

> [!IMPORTANT]
> Your manifest cannot specify `WordApiHiddenDocument 1.5` as an activation requirement. It's not a valid value to use in the [Set element](../../manifest/set.md).

## API list

The following table lists the Word JavaScript APIs currently included in the `WordApiHiddenDocument 1.5` requirement set. For a complete list of all Word JavaScript APIs (including `WordApiHiddenDocument 1.5` APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-1.5-hidden-document&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[addStyle(name: string, type: Word.StyleType)](/javascript/api/word/word.documentcreated#word-word-documentcreated-addstyle-member(1))|Adds a style into the document by name and type.|
||[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getcontentcontrols-member(1))|Gets the currently supported content controls in the document.|
||[getStyles()](/javascript/api/word/word.documentcreated#word-word-documentcreated-getstyles-member(1))|Gets a StyleCollection object that represents the whole style set of the document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End", insertFileOptions?: Word.InsertFileOptions)](/javascript/api/word/word.documentcreated#word-word-documentcreated-insertfilefrombase64-member(1))|Inserts a document into the target document at a specific location with additional properties.|
||[save(saveBehavior?: Word.SaveBehavior, fileName?: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|Saves the document.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word?view=word-js-1.5-hidden-document&preserve-view=true)
- [Word JavaScript preview APIs](word-preview-apis.md)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
- [Understanding platform-specific requirement sets](https://aka.ms/PlatformSpecificReqtSets)
