---
title: Word JavaScript preview APIs
description: Details about upcoming Word JavaScript APIs.
ms.date: 04/20/2023
ms.topic: whats-new
ms.localizationpriority: medium
---

# Word JavaScript preview APIs

New Word JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

> [!IMPORTANT]
> Note that the following Word preview APIs may be available on the following platforms.
>
> - Word on Windows
> - Word on Mac
>
> Word preview APIs are currently not supported on iPad. However, bookmark feature APIs are also available in Word on the web. For APIs available only in Word on the web, see the [Web-only API list](#web-only-api-list).

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## API list

The following table lists the Word JavaScript APIs currently in preview, except those that are [available only in Word on the web](#web-only-api-list). To see a complete list of all Word JavaScript APIs (including preview APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[retrieveStylesFromBase64(base64File: string)](/javascript/api/word/word.application#word-word-application-retrievestylesfrombase64-member(1))|Parse styles from template base 64 file and return JSON format of retrieved styles as a string.|
|[Body](/javascript/api/word/word.body)|[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.body#word-word-body-getcontentcontrols-member(1))|Gets the currently supported content controls in the body.|
||[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| "RichText" \| "PlainText")](/javascript/api/word/word.body#word-word-body-insertcontentcontrol-member(1))|Wraps the Body object with a content control.|
||[styleBuiltIn](/javascript/api/word/word.body#word-word-body-stylebuiltin-member)|Specifies the built-in style name for the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcontentcontrols-member(1))|Gets the currently supported child content controls in this content control.|
||[onDataChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member)|Occurs when data within the content control are changed.|
||[onDeleted](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member)|Occurs when the content control is deleted.|
||[onEntered](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onentered-member)|Occurs when the content control is entered.|
||[onExited](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onexited-member)|Occurs when the content control is exited, for example, when the cursor leaves the content control.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member)|Occurs when selection within the content control is changed.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-stylebuiltin-member)|Specifies the built-in style name for the content control.|
|[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs)|[eventType](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-source-member)|The source of the event.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByChangeTrackingStates(changeTrackingStates: Word.ChangeTrackingState[])](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbychangetrackingstates-member(1))|Gets the content controls that have the specified tracking state.|
|[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs)|[eventType](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-source-member)|The source of the event.|
|[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs)|[eventType](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-source-member)|The source of the event.|
|[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs)|[eventType](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-source-member)|The source of the event.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentcontrol-member)|The object that raised the event.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-source-member)|The source of the event.|
|[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs)|[eventType](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-source-member)|The source of the event.|
|[ContentControlOptions](/javascript/api/word/word.contentcontroloptions)|[types](/javascript/api/word/word.contentcontroloptions#word-word-contentcontroloptions-types-member)|An array of content control types, item must be 'RichText' or 'PlainText'.|
|[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs)|[eventType](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-source-member)|The source of the event.|
|[Document](/javascript/api/word/word.document)|[addStyle(name: string, type: Word.StyleType)](/javascript/api/word/word.document#word-word-document-addstyle-member(1))|Adds a style into the document by name and type.|
||[close(closeBehavior?: Word.CloseBehavior)](/javascript/api/word/word.document#word-word-document-close-member(1))|Close current document.|
||[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.document#word-word-document-getcontentcontrols-member(1))|Gets the currently supported content controls in the document.|
||[getStyles()](/javascript/api/word/word.document#word-word-document-getstyles-member(1))|Gets a StyleCollection object that represents the whole style set of the document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End", insertFileOptions?: Word.InsertFileOptions)](/javascript/api/word/word.document#word-word-document-insertfilefrombase64-member(1))|Inserts a document into the target document at a specific location with additional properties.|
||[onContentControlAdded](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member)|Occurs when a content control is added.|
||[save(saveBehavior?: Word.SaveBehavior, fileName?: string)](/javascript/api/word/word.document#word-word-document-save-member(1))|Saves the document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[addStyle(name: string, type: Word.StyleType)](/javascript/api/word/word.documentcreated#word-word-documentcreated-addstyle-member(1))|Adds a style into the document by name and type.|
||[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getcontentcontrols-member(1))|Gets the currently supported content controls in the document.|
||[getStyles()](/javascript/api/word/word.documentcreated#word-word-documentcreated-getstyles-member(1))|Gets a StyleCollection object that represents the whole style set of the document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End", insertFileOptions?: Word.InsertFileOptions)](/javascript/api/word/word.documentcreated#word-word-documentcreated-insertfilefrombase64-member(1))|Inserts a document into the target document at a specific location with additional properties.|
||[save(saveBehavior?: Word.SaveBehavior, fileName?: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|Saves the document.|
|[Field](/javascript/api/word/word.field)|[code](/javascript/api/word/word.field#word-word-field-code-member)|Specifies the field's code instruction.|
||[data](/javascript/api/word/word.field#word-word-field-data-member)|Specifies data in an "Addin" field.|
||[delete()](/javascript/api/word/word.field#word-word-field-delete-member(1))|Deletes the field.|
||[kind](/javascript/api/word/word.field#word-word-field-kind-member)|Gets the field's kind.|
||[locked](/javascript/api/word/word.field#word-word-field-locked-member)|Specifies whether the field is locked.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.field#word-word-field-select-member(1))|Selects the field.|
||[showCodes](/javascript/api/word/word.field#word-word-field-showcodes-member)|Specifies whether the field codes are displayed for the specified field.|
||[type](/javascript/api/word/word.field#word-word-field-type-member)|Gets the field's type.|
||[updateResult()](/javascript/api/word/word.field#word-word-field-updateresult-member(1))|Updates the field.|
|[FieldCollection](/javascript/api/word/word.fieldcollection)|[getByTypes(types: Word.FieldType[])](/javascript/api/word/word.fieldcollection#word-word-fieldcollection-getbytypes-member(1))|Gets the Field object collection including the specified types of fields.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|Gets the format of the inline image.|
|[InsertFileOptions](/javascript/api/word/word.insertfileoptions)|[importChangeTrackingMode](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importchangetrackingmode-member)|Represents whether the change tracking mode status from the source document should be imported.|
||[importPageColor](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importpagecolor-member)|Represents whether the page color and other background information from the source document should be imported.|
||[importParagraphSpacing](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importparagraphspacing-member)|Represents whether the paragraph spacing from the source document should be imported.|
||[importStyles](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importstyles-member)|Represents whether the styles from the source document should be imported.|
||[importTheme](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importtheme-member)|Represents whether the theme from the source document should be imported.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#word-word-list-getlevelfont-member(1))|Gets the font of the bullet, number, or picture at the specified level in the list.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#word-word-list-getlevelpicture-member(1))|Gets the Base64-encoded string representation of the picture at the specified level in the list.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|Resets the font of the bullet, number, or picture at the specified level in the list.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#word-word-list-setlevelpicture-member(1))|Sets the picture at the specified level in the list.|
|[ListLevel](/javascript/api/word/word.listlevel)|[alignment](/javascript/api/word/word.listlevel#word-word-listlevel-alignment-member)|Specifies the horizontal alignment of the list level.|
||[font](/javascript/api/word/word.listlevel#word-word-listlevel-font-member)|Gets a Font object that represents the character formatting of the specified object.|
||[linkedStyle](/javascript/api/word/word.listlevel#word-word-listlevel-linkedstyle-member)|Specifies the name of the style that's linked to the specified list level object.|
||[numberFormat](/javascript/api/word/word.listlevel#word-word-listlevel-numberformat-member)|Specifies the number format for the specified list level.|
||[numberPosition](/javascript/api/word/word.listlevel#word-word-listlevel-numberposition-member)|Specifies the position (in points) of the number or bullet for the specified list level object.|
||[numberStyle](/javascript/api/word/word.listlevel#word-word-listlevel-numberstyle-member)|Specifies the number style for the list level object.|
||[resetOnHigher](/javascript/api/word/word.listlevel#word-word-listlevel-resetonhigher-member)|Specifies the list level that must appear before the specified list level restarts numbering at 1.|
||[startAt](/javascript/api/word/word.listlevel#word-word-listlevel-startat-member)|Specifies the starting number for the specified list level object.|
||[tabPosition](/javascript/api/word/word.listlevel#word-word-listlevel-tabposition-member)|Specifies the tab position for the specified list level object.|
||[textPosition](/javascript/api/word/word.listlevel#word-word-listlevel-textposition-member)|Specifies the position (in points) for the second line of wrapping text for the specified list level object.|
||[trailingCharacter](/javascript/api/word/word.listlevel#word-word-listlevel-trailingcharacter-member)|Specifies the character inserted after the number for the specified list level.|
|[ListLevelCollection](/javascript/api/word/word.listlevelcollection)|[getFirst()](/javascript/api/word/word.listlevelcollection#word-word-listlevelcollection-getfirst-member(1))|Gets the first list level in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.listlevelcollection#word-word-listlevelcollection-getfirstornullobject-member(1))|Gets the first list level in this collection.|
||[items](/javascript/api/word/word.listlevelcollection#word-word-listlevelcollection-items-member)|Gets the loaded child items in this collection.|
|[ListTemplate](/javascript/api/word/word.listtemplate)|[listLevels](/javascript/api/word/word.listtemplate#word-word-listtemplate-listlevels-member)|Gets a ListLevels collection that represents all the levels for the specified ListTemplate.|
||[outlineNumbered](/javascript/api/word/word.listtemplate#word-word-listtemplate-outlinenumbered-member)|Specifies whether the specified ListTemplate object is outline numbered.|
|[Paragraph](/javascript/api/word/word.paragraph)|[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-getcontentcontrols-member(1))|Gets the currently supported content controls in the paragraph.|
||[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| "RichText" \| "PlainText")](/javascript/api/word/word.paragraph#word-word-paragraph-insertcontentcontrol-member(1))|Wraps the Paragraph object with a content control.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#word-word-paragraph-stylebuiltin-member)|Specifies the built-in style name for the paragraph.|
|[ParagraphFormat](/javascript/api/word/word.paragraphformat)|[alignment](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-alignment-member)|Specifies the alignment for the specified paragraphs.|
||[firstLineIndent](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-firstlineindent-member)|Specifies the value (in points) for a first line or hanging indent.|
||[keepTogether](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-keeptogether-member)|Specifies whether all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.|
||[keepWithNext](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-keepwithnext-member)|Specifies whether the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.|
||[leftIndent](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-leftindent-member)|Specifies the left indent.|
||[lineSpacing](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-linespacing-member)|Specifies the line spacing (in points) for the specified paragraphs.|
||[lineUnitAfter](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-lineunitafter-member)|Specifies the amount of spacing (in gridlines) after the specified paragraphs.|
||[lineUnitBefore](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-lineunitbefore-member)|Specifies the amount of spacing (in gridlines) before the specified paragraphs.|
||[mirrorIndents](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-mirrorindents-member)|Specifies whether left and right indents are the same width.|
||[outlineLevel](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-outlinelevel-member)|Specifies the outline level for the specified paragraphs.|
||[rightIndent](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-rightindent-member)|Specifies the right indent (in points) for the specified paragraphs.|
||[spaceAfter](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-spaceafter-member)|Specifies the amount of spacing (in points) after the specified paragraph or text column.|
||[spaceBefore](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-spacebefore-member)|Specifies the spacing (in points) before the specified paragraphs.|
||[widowControl](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-widowcontrol-member)|Specifies whether the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Microsoft Word repaginates the document.|
|[Range](/javascript/api/word/word.range)|[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.range#word-word-range-getcontentcontrols-member(1))|Gets the currently supported content controls in the range.|
||[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| "RichText" \| "PlainText")](/javascript/api/word/word.range#word-word-range-insertcontentcontrol-member(1))|Wraps the Range object with a content control.|
||[insertField(insertLocation: Word.InsertLocation \| "Replace" \| "Start" \| "End" \| "Before" \| "After", fieldType?: Word.FieldType, text?: string, removeFormatting?: boolean)](/javascript/api/word/word.range#word-word-range-insertfield-member(1))|Inserts a field at the specified location.|
||[styleBuiltIn](/javascript/api/word/word.range#word-word-range-stylebuiltin-member)|Specifies the built-in style name for the range.|
|[Style](/javascript/api/word/word.style)|[baseStyle](/javascript/api/word/word.style#word-word-style-basestyle-member)|Gets the name of an existing style to use as the base formatting of another style.|
||[builtIn](/javascript/api/word/word.style#word-word-style-builtin-member)|Gets whether the specified style is a built-in style.|
||[delete()](/javascript/api/word/word.style#word-word-style-delete-member(1))|Deletes the style.|
||[description](/javascript/api/word/word.style#word-word-style-description-member)|Gets the description of the specified style.|
||[font](/javascript/api/word/word.style#word-word-style-font-member)|Gets a font object that represents the character formatting of the specified style.|
||[inUse](/javascript/api/word/word.style#word-word-style-inuse-member)|Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.|
||[linked](/javascript/api/word/word.style#word-word-style-linked-member)|Gets whether a style is a linked style that can be used for both paragraph and character formatting.|
||[listTemplate](/javascript/api/word/word.style#word-word-style-listtemplate-member)|Gets a ListTemplate object that represents the list formatting for the specified Style object.|
||[nameLocal](/javascript/api/word/word.style#word-word-style-namelocal-member)|Gets the name of a style in the language of the user.|
||[nextParagraphStyle](/javascript/api/word/word.style#word-word-style-nextparagraphstyle-member)|Gets the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.|
||[paragraphFormat](/javascript/api/word/word.style#word-word-style-paragraphformat-member)|Gets a ParagraphFormat object that represents the paragraph settings for the specified style.|
||[priority](/javascript/api/word/word.style#word-word-style-priority-member)|Specifies the priority.|
||[quickStyle](/javascript/api/word/word.style#word-word-style-quickstyle-member)|Specifies whether the style corresponds to an available quick style.|
||[type](/javascript/api/word/word.style#word-word-style-type-member)|Gets the style type.|
||[unhideWhenUsed](/javascript/api/word/word.style#word-word-style-unhidewhenused-member)|Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.|
||[visibility](/javascript/api/word/word.style#word-word-style-visibility-member)|Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.|
|[StyleCollection](/javascript/api/word/word.stylecollection)|[getByName(name: string)](/javascript/api/word/word.stylecollection#word-word-stylecollection-getbyname-member(1))|Get the style object by its name.|
||[getByNameOrNullObject(name: string)](/javascript/api/word/word.stylecollection#word-word-stylecollection-getbynameornullobject-member(1))|If the corresponding style doesn't exist, then this method returns an object with its `isNullObject` property set to `true`.|
||[getCount()](/javascript/api/word/word.stylecollection#word-word-stylecollection-getcount-member(1))|Gets the number of the styles in the collection.|
||[getItem(index: number)](/javascript/api/word/word.stylecollection#word-word-stylecollection-getitem-member(1))|Gets a style object by its index in the collection.|
||[items](/javascript/api/word/word.stylecollection#word-word-stylecollection-items-member)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/word/word.table)|[styleBuiltIn](/javascript/api/word/word.table#word-word-table-stylebuiltin-member)|Specifies the built-in style name for the table.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|Inserts a content control on the row.|

## Web-only API list

The following table lists the Word JavaScript APIs currently in preview only in Word on the web. To see a complete list of all Word JavaScript APIs (including preview APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[onCommentAdded](/javascript/api/word/word.body#word-word-body-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.body#word-word-body-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.body#word-word-body-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.body#word-word-body-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.body#word-word-body-oncommentselected-member)|Occurs when a comment is selected.|
|[CommentDetail](/javascript/api/word/word.commentdetail)|[id](/javascript/api/word/word.commentdetail#word-word-commentdetail-id-member)|Represents the ID of this comment.|
||[replyIds](/javascript/api/word/word.commentdetail#word-word-commentdetail-replyids-member)|Represents the IDs of the replies to this comment.|
|[CommentEventArgs](/javascript/api/word/word.commenteventargs)|[changeType](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-changetype-member)|Represents how the comment changed event is triggered.|
||[commentDetails](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-commentdetails-member)|Gets the CommentDetail array which contains the IDs and reply IDs of the involved comments.|
||[source](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-type-member)|The event type.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onCommentAdded](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentselected-member)|Occurs when a comment is selected.|
|[Paragraph](/javascript/api/word/word.paragraph)|[onCommentAdded](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentselected-member)|Occurs when a comment is selected.|
|[Range](/javascript/api/word/word.range)|[onCommentAdded](/javascript/api/word/word.range#word-word-range-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.range#word-word-range-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.range#word-word-range-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.range#word-word-range-oncommentselected-member)|Occurs when a comment is selected.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
