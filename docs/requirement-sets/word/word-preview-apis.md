---
title: Word JavaScript preview APIs
description: Details about upcoming Word JavaScript APIs.
ms.date: 01/25/2024
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
> Word preview APIs are currently not supported on iPad. However, several APIs may also be available in Word on the web. For APIs available only in Word on the web, see the [Web-only API list](#web-only-api-list).

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## API list

The following table lists the Word JavaScript APIs currently in preview, except those that are [available only in Word on the web](#web-only-api-list). To see a complete list of all Word JavaScript APIs (including preview APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Annotation](/javascript/api/word/word.annotation)|[critiqueAnnotation](/javascript/api/word/word.annotation#word-word-annotation-critiqueannotation-member)|Gets the critique annotation object.|
||[delete()](/javascript/api/word/word.annotation#word-word-annotation-delete-member(1))|Deletes the annotation.|
||[id](/javascript/api/word/word.annotation#word-word-annotation-id-member)|Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.|
||[state](/javascript/api/word/word.annotation#word-word-annotation-state-member)|Gets the state of the annotation.|
|[AnnotationClickedEventArgs](/javascript/api/word/word.annotationclickedeventargs)|[id](/javascript/api/word/word.annotationclickedeventargs#word-word-annotationclickedeventargs-id-member)|Specifies the annotation ID for which the event was fired.|
|[AnnotationCollection](/javascript/api/word/word.annotationcollection)|[getFirst()](/javascript/api/word/word.annotationcollection#word-word-annotationcollection-getfirst-member(1))|Gets the first annotation in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.annotationcollection#word-word-annotationcollection-getfirstornullobject-member(1))|Gets the first annotation in this collection.|
||[items](/javascript/api/word/word.annotationcollection#word-word-annotationcollection-items-member)|Gets the loaded child items in this collection.|
|[AnnotationHoveredEventArgs](/javascript/api/word/word.annotationhoveredeventargs)|[id](/javascript/api/word/word.annotationhoveredeventargs#word-word-annotationhoveredeventargs-id-member)|Specifies the annotation ID for which the event was fired.|
|[AnnotationInsertedEventArgs](/javascript/api/word/word.annotationinsertedeventargs)|[ids](/javascript/api/word/word.annotationinsertedeventargs#word-word-annotationinsertedeventargs-ids-member)|Specifies the annotation IDs for which the event was fired.|
|[AnnotationPopupActionEventArgs](/javascript/api/word/word.annotationpopupactioneventargs)|[action](/javascript/api/word/word.annotationpopupactioneventargs#word-word-annotationpopupactioneventargs-action-member)|Specifies the chosen action in the pop-up menu.|
||[critiqueSuggestion](/javascript/api/word/word.annotationpopupactioneventargs#word-word-annotationpopupactioneventargs-critiquesuggestion-member)|Specifies the suggestion accepted (only populated when accepting a critique suggestion).|
||[id](/javascript/api/word/word.annotationpopupactioneventargs#word-word-annotationpopupactioneventargs-id-member)|Specifies the annotation ID for which the event was fired.|
|[AnnotationRemovedEventArgs](/javascript/api/word/word.annotationremovedeventargs)|[ids](/javascript/api/word/word.annotationremovedeventargs#word-word-annotationremovedeventargs-ids-member)|Specifies the annotation IDs for which the event was fired.|
|[AnnotationSet](/javascript/api/word/word.annotationset)|[critiques](/javascript/api/word/word.annotationset#word-word-annotationset-critiques-member)|Critiques.|
|[Body](/javascript/api/word/word.body)|[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| Word.ContentControlType.checkBox \| "RichText" \| "PlainText" \| "CheckBox")](/javascript/api/word/word.body#word-word-body-insertcontentcontrol-member(1))|Wraps the Body object with a content control.|
|[Border](/javascript/api/word/word.border)|[color](/javascript/api/word/word.border#word-word-border-color-member)|Specifies the color for the border.|
||[location](/javascript/api/word/word.border#word-word-border-location-member)|Gets the location of the border.|
||[type](/javascript/api/word/word.border#word-word-border-type-member)|Specifies the border type for the border.|
||[visible](/javascript/api/word/word.border#word-word-border-visible-member)|Specifies whether the border is visible.|
||[width](/javascript/api/word/word.border#word-word-border-width-member)|Specifies the width for the border.|
|[BorderCollection](/javascript/api/word/word.bordercollection)|[getByLocation(borderLocation: Word.BorderLocation.top \| Word.BorderLocation.left \| Word.BorderLocation.bottom \| Word.BorderLocation.right \| Word.BorderLocation.insideHorizontal \| Word.BorderLocation.insideVertical \| "Top" \| "Left" \| "Bottom" \| "Right" \| "InsideHorizontal" \| "InsideVertical")](/javascript/api/word/word.bordercollection#word-word-bordercollection-getbylocation-member(1))|Gets the border that has the specified location.|
||[getFirst()](/javascript/api/word/word.bordercollection#word-word-bordercollection-getfirst-member(1))|Gets the first border in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.bordercollection#word-word-bordercollection-getfirstornullobject-member(1))|Gets the first border in this collection.|
||[getItem(index: number)](/javascript/api/word/word.bordercollection#word-word-bordercollection-getitem-member(1))|Gets a Border object by its index in the collection.|
||[insideBorderColor](/javascript/api/word/word.bordercollection#word-word-bordercollection-insidebordercolor-member)|Specifies the 24-bit color of the inside borders.|
||[insideBorderType](/javascript/api/word/word.bordercollection#word-word-bordercollection-insidebordertype-member)|Specifies the border type of the inside borders.|
||[insideBorderWidth](/javascript/api/word/word.bordercollection#word-word-bordercollection-insideborderwidth-member)|Specifies the width of the inside borders.|
||[items](/javascript/api/word/word.bordercollection#word-word-bordercollection-items-member)|Gets the loaded child items in this collection.|
||[outsideBorderColor](/javascript/api/word/word.bordercollection#word-word-bordercollection-outsidebordercolor-member)|Specifies the 24-bit color of the outside borders.|
||[outsideBorderType](/javascript/api/word/word.bordercollection#word-word-bordercollection-outsidebordertype-member)|Specifies the border type of the outside borders.|
||[outsideBorderWidth](/javascript/api/word/word.bordercollection#word-word-bordercollection-outsideborderwidth-member)|Specifies the width of the outside borders.|
|[CheckboxContentControl](/javascript/api/word/word.checkboxcontentcontrol)|[isChecked](/javascript/api/word/word.checkboxcontentcontrol#word-word-checkboxcontentcontrol-ischecked-member)|Specifies the current state of the checkbox.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[checkboxContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-checkboxcontentcontrol-member)|Specifies the checkbox-related data if the content control's type is 'CheckBox'.|
|[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs)|[eventType](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-eventtype-member)|The event type.|
|[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs)|[eventType](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
|[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs)|[eventType](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-eventtype-member)|The event type.|
|[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs)|[eventType](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-eventtype-member)|The event type.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentcontrol-member)|The object that raised the event.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-source-member)|The source of the event.|
|[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs)|[eventType](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-eventtype-member)|The event type.|
|[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs)|[eventType](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
|[Critique](/javascript/api/word/word.critique)|[colorScheme](/javascript/api/word/word.critique#word-word-critique-colorscheme-member)|Gets the color scheme of the critique.|
||[length](/javascript/api/word/word.critique#word-word-critique-length-member)|Gets the length of the critique inside paragraph.|
||[popupOptions](/javascript/api/word/word.critique#word-word-critique-popupoptions-member)|Specifies the behavior of the pop-up menu for the critique.|
||[start](/javascript/api/word/word.critique#word-word-critique-start-member)|Gets the start index of the critique inside paragraph.|
|[CritiqueAnnotation](/javascript/api/word/word.critiqueannotation)|[accept()](/javascript/api/word/word.critiqueannotation#word-word-critiqueannotation-accept-member(1))|Accepts the critique.|
||[critique](/javascript/api/word/word.critiqueannotation#word-word-critiqueannotation-critique-member)|Gets the critique that was passed when the annotation was inserted.|
||[range](/javascript/api/word/word.critiqueannotation#word-word-critiqueannotation-range-member)|Gets the range of text that is annotated.|
||[reject()](/javascript/api/word/word.critiqueannotation#word-word-critiqueannotation-reject-member(1))|Rejects the critique.|
|[CritiquePopupOptions](/javascript/api/word/word.critiquepopupoptions)|[brandingTextResourceId](/javascript/api/word/word.critiquepopupoptions#word-word-critiquepopupoptions-brandingtextresourceid-member)|Gets the manifest resource ID of the string to use for branding.|
||[subtitleResourceId](/javascript/api/word/word.critiquepopupoptions#word-word-critiquepopupoptions-subtitleresourceid-member)|Gets the manifest resource ID of the string to use as the subtitle.|
||[suggestions](/javascript/api/word/word.critiquepopupoptions#word-word-critiquepopupoptions-suggestions-member)|Gets the suggestions to display in the critique pop-up menu.|
||[titleResourceId](/javascript/api/word/word.critiquepopupoptions#word-word-critiquepopupoptions-titleresourceid-member)|Gets the manifest resource ID of the string to use as the title.|
|[Document](/javascript/api/word/word.document)|[compare(filePath: string, documentCompareOptions?: Word.DocumentCompareOptions)](/javascript/api/word/word.document#word-word-document-compare-member(1))|Displays revision marks that indicate where the specified document differs from another document.|
||[getAnnotationById(id: string)](/javascript/api/word/word.document#word-word-document-getannotationbyid-member(1))|Gets the annotation by ID.|
||[onAnnotationClicked](/javascript/api/word/word.document#word-word-document-onannotationclicked-member)|Occurs when the user clicks an annotation (or selects it using **Alt+Down**).|
||[onAnnotationHovered](/javascript/api/word/word.document#word-word-document-onannotationhovered-member)|Occurs when the user hovers the cursor over an annotation.|
||[onAnnotationInserted](/javascript/api/word/word.document#word-word-document-onannotationinserted-member)|Occurs when the user adds one or more annotations.|
||[onAnnotationPopupAction](/javascript/api/word/word.document#word-word-document-onannotationpopupaction-member)|Occurs when the user performs an action in an annotation pop-up menu.|
||[onAnnotationRemoved](/javascript/api/word/word.document#word-word-document-onannotationremoved-member)|Occurs when the user deletes one or more annotations.|
|[DocumentCompareOptions](/javascript/api/word/word.documentcompareoptions)|[addToRecentFiles](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-addtorecentfiles-member)|True adds the document to the list of recently used files on the File menu.|
||[authorName](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-authorname-member)|The reviewer name associated with the differences generated by the comparison.|
||[compareTarget](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-comparetarget-member)|The target document for the comparison.|
||[detectFormatChanges](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-detectformatchanges-member)|True (default) for the comparison to include detection of format changes.|
||[ignoreAllComparisonWarnings](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-ignoreallcomparisonwarnings-member)|True compares the documents without notifying a user of problems.|
||[removeDateAndTime](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-removedateandtime-member)|True removes date and time stamp information from tracked changes in the returned Document object.|
||[removePersonalInformation](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-removepersonalinformation-member)|True removes all user information from comments, revisions, and the properties dialog box in the returned Document object.|
|[Field](/javascript/api/word/word.field)|[showCodes](/javascript/api/word/word.field#word-word-field-showcodes-member)|Specifies whether the field codes are displayed for the specified field.|
|[Font](/javascript/api/word/word.font)|[hidden](/javascript/api/word/word.font#word-word-font-hidden-member)|Specifies a value that indicates whether the font is tagged as hidden.|
|[GetTextOptions](/javascript/api/word/word.gettextoptions)|[includeHiddenText](/javascript/api/word/word.gettextoptions#word-word-gettextoptions-includehiddentext-member)|Specifies a value that indicates whether to include hidden text in the result of the GetText method.|
||[includeTextMarkedAsDeleted](/javascript/api/word/word.gettextoptions#word-word-gettextoptions-includetextmarkedasdeleted-member)|Specifies a value that indicates whether to include text marked as deleted in the result of the GetText method.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|Gets the format of the inline image.|
|[InsertFileOptions](/javascript/api/word/word.insertfileoptions)|[importDifferentOddEvenPages](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importdifferentoddevenpages-member)|Represents whether to import the Different Odd and Even Pages setting for the header and footer from the source document.|
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
|[Paragraph](/javascript/api/word/word.paragraph)|[getAnnotations()](/javascript/api/word/word.paragraph#word-word-paragraph-getannotations-member(1))|Gets annotations set on this Paragraph object.|
||[getText(options?: Word.GetTextOptions \| { IncludeHiddenText?: boolean IncludeTextMarkedAsDeleted?: boolean })](/javascript/api/word/word.paragraph#word-word-paragraph-gettext-member(1))|Returns the text of the paragraph.|
||[insertAnnotations(annotations: Word.AnnotationSet)](/javascript/api/word/word.paragraph#word-word-paragraph-insertannotations-member(1))|Inserts annotations on this Paragraph object.|
||[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| Word.ContentControlType.checkBox \| "RichText" \| "PlainText" \| "CheckBox")](/javascript/api/word/word.paragraph#word-word-paragraph-insertcontentcontrol-member(1))|Wraps the Paragraph object with a content control.|
|[ParagraphAddedEventArgs](/javascript/api/word/word.paragraphaddedeventargs)|[type](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-type-member)|The event type.|
|[ParagraphChangedEventArgs](/javascript/api/word/word.paragraphchangedeventargs)|[type](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-type-member)|The event type.|
|[ParagraphDeletedEventArgs](/javascript/api/word/word.paragraphdeletedeventargs)|[type](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-type-member)|The event type.|
|[Range](/javascript/api/word/word.range)|[highlight()](/javascript/api/word/word.range#word-word-range-highlight-member(1))|Highlights the range temporarily without changing document content.|
||[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| Word.ContentControlType.checkBox \| "RichText" \| "PlainText" \| "CheckBox")](/javascript/api/word/word.range#word-word-range-insertcontentcontrol-member(1))|Wraps the Range object with a content control.|
||[removeHighlight()](/javascript/api/word/word.range#word-word-range-removehighlight-member(1))|Removes the highlight added by the Highlight function if any.|
|[Shading](/javascript/api/word/word.shading)|[foregroundPatternColor](/javascript/api/word/word.shading#word-word-shading-foregroundpatterncolor-member)|Specifies the color for the foreground of the object.|
||[texture](/javascript/api/word/word.shading#word-word-shading-texture-member)|Specifies the shading texture of the object.|
|[Style](/javascript/api/word/word.style)|[borders](/javascript/api/word/word.style#word-word-style-borders-member)|Specifies a BorderCollection object that represents all the borders for the specified style.|
||[description](/javascript/api/word/word.style#word-word-style-description-member)|Gets the description of the specified style.|
||[listTemplate](/javascript/api/word/word.style#word-word-style-listtemplate-member)|Gets a ListTemplate object that represents the list formatting for the specified Style object.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|Inserts a content control on the row.|
|[TableStyle](/javascript/api/word/word.tablestyle)|[alignment](/javascript/api/word/word.tablestyle#word-word-tablestyle-alignment-member)|Specifies the table's alignment against the page margin.|
||[allowBreakAcrossPage](/javascript/api/word/word.tablestyle#word-word-tablestyle-allowbreakacrosspage-member)|Specifies whether lines in tables formatted with a specified style break across pages.|

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
