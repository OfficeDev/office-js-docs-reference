---
title: Word JavaScript preview APIs
description: Details about upcoming Word JavaScript APIs.
ms.date: 02/11/2025
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
|[Body](/javascript/api/word/word.body)|[shapes](/javascript/api/word/word.body#word-word-body-shapes-member)|Gets the collection of shape objects in the body, including both inline and floating shapes.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Gets the type of the body.|
|[Canvas](/javascript/api/word/word.canvas)|[id](/javascript/api/word/word.canvas#word-word-canvas-id-member)|Gets an integer that represents the canvas identifier.|
||[shape](/javascript/api/word/word.canvas#word-word-canvas-shape-member)|Gets the Shape object associated with the canvas.|
||[shapes](/javascript/api/word/word.canvas#word-word-canvas-shapes-member)|Gets the collection of Shape objects.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[resetState()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-resetstate-member(1))|Resets the state of the content control.|
||[setState(contentControlState: Word.ContentControlState)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-setstate-member(1))|Sets the state of the content control.|
|[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs)|[eventType](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-eventtype-member)|The event type.|
|[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs)|[eventType](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
|[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs)|[eventType](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-eventtype-member)|The event type.|
|[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs)|[eventType](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-eventtype-member)|The event type.|
|[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs)|[eventType](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-eventtype-member)|The event type.|
|[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs)|[eventType](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
|[Document](/javascript/api/word/word.document)|[activeWindow](/javascript/api/word/word.document#word-word-document-activewindow-member)|Gets the active window for the document.|
||[compareFromBase64(base64File: string, documentCompareOptions?: Word.DocumentCompareOptions)](/javascript/api/word/word.document#word-word-document-comparefrombase64-member(1))|Displays revision marks that indicate where the specified document differs from another document.|
||[windows](/javascript/api/word/word.document#word-word-document-windows-member)|Gets the collection of `Word.Window` objects for the document.|
|[Font](/javascript/api/word/word.font)|[hidden](/javascript/api/word/word.font#word-word-font-hidden-member)|Specifies a value that indicates whether the font is tagged as hidden.|
|[InsertShapeOptions](/javascript/api/word/word.insertshapeoptions)|[height](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-height-member)|Represents the height of the shape being inserted.|
||[left](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-left-member)|Represents the left position of the shape being inserted.|
||[top](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-top-member)|Represents the top position of the shape being inserted.|
||[width](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-width-member)|Represents the width of the shape being inserted.|
|[Page](/javascript/api/word/word.page)|[getNext()](/javascript/api/word/word.page#word-word-page-getnext-member(1))|Gets the next page in the pane.|
||[getNextOrNullObject()](/javascript/api/word/word.page#word-word-page-getnextornullobject-member(1))|Gets the next page.|
||[getRange(rangeLocation?: Word.RangeLocation.whole \| Word.RangeLocation.start \| Word.RangeLocation.end \| "Whole" \| "Start" \| "End")](/javascript/api/word/word.page#word-word-page-getrange-member(1))|Gets the whole page, or the starting or ending point of the page, as a range.|
||[height](/javascript/api/word/word.page#word-word-page-height-member)|Gets the height, in points, of the paper defined in the Page Setup dialog box.|
||[index](/javascript/api/word/word.page#word-word-page-index-member)|Gets the index of the page.|
||[width](/javascript/api/word/word.page#word-word-page-width-member)|Gets the width, in points, of the paper defined in the Page Setup dialog box.|
|[PageCollection](/javascript/api/word/word.pagecollection)|[getFirst()](/javascript/api/word/word.pagecollection#word-word-pagecollection-getfirst-member(1))|Gets the first page in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.pagecollection#word-word-pagecollection-getfirstornullobject-member(1))|Gets the first page in this collection.|
||[getItem(index: number)](/javascript/api/word/word.pagecollection#word-word-pagecollection-getitem-member(1))|Gets a Page object by its index in the collection.|
||[items](/javascript/api/word/word.pagecollection#word-word-pagecollection-items-member)|Gets the loaded child items in this collection.|
|[Pane](/javascript/api/word/word.pane)|[getNext()](/javascript/api/word/word.pane#word-word-pane-getnext-member(1))|Gets the next pane in the window.|
||[getNextOrNullObject()](/javascript/api/word/word.pane#word-word-pane-getnextornullobject-member(1))|Gets the next pane.|
||[pages](/javascript/api/word/word.pane#word-word-pane-pages-member)|Gets the collection of pages in the pane.|
||[pagesEnclosingViewport](/javascript/api/word/word.pane#word-word-pane-pagesenclosingviewport-member)|Gets the `PageCollection` shown in the viewport of the pane.|
|[PaneCollection](/javascript/api/word/word.panecollection)|[getFirst()](/javascript/api/word/word.panecollection#word-word-panecollection-getfirst-member(1))|Gets the first pane in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.panecollection#word-word-panecollection-getfirstornullobject-member(1))|Gets the first pane in this collection.|
||[getItem(index: number)](/javascript/api/word/word.panecollection#word-word-panecollection-getitem-member(1))|Gets a Pane object by its index in the collection.|
||[items](/javascript/api/word/word.panecollection#word-word-panecollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[insertCanvas(insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-insertcanvas-member(1))|Inserts a floating canvas in front of text with its anchor at the beginning of the paragraph.|
||[insertGeometricShape(geometricShapeType: Word.GeometricShapeType, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-insertgeometricshape-member(1))|Inserts a geometric shape in front of text with its anchor at the beginning of the paragraph.|
||[insertPictureFromBase64(base64EncodedImage: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-insertpicturefrombase64-member(1))|Inserts a floating picture in front of text with its anchor at the beginning of the paragraph.|
||[insertTextBox(text?: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-inserttextbox-member(1))|Inserts a floating text box in front of text with its anchor at the beginning of the paragraph.|
||[shapes](/javascript/api/word/word.paragraph#word-word-paragraph-shapes-member)|Gets the collection of shape objects anchored in the paragraph, including both inline and floating shapes.|
|[ParagraphAddedEventArgs](/javascript/api/word/word.paragraphaddedeventargs)|[type](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-type-member)|The event type.|
|[ParagraphChangedEventArgs](/javascript/api/word/word.paragraphchangedeventargs)|[type](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-type-member)|The event type.|
|[ParagraphDeletedEventArgs](/javascript/api/word/word.paragraphdeletedeventargs)|[type](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-type-member)|The event type.|
|[Range](/javascript/api/word/word.range)|[insertCanvas(insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-insertcanvas-member(1))|Inserts a floating canvas in front of text with its anchor at the beginning of the range.|
||[insertGeometricShape(geometricShapeType: Word.GeometricShapeType, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-insertgeometricshape-member(1))|Inserts a geometric shape in front of text with its anchor at the beginning of the range.|
||[insertPictureFromBase64(base64EncodedImage: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-insertpicturefrombase64-member(1))|Inserts a floating picture in front of text with its anchor at the beginning of the range.|
||[insertTextBox(text?: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-inserttextbox-member(1))|Inserts a floating text box in front of text with its anchor at the beginning of the range.|
||[pages](/javascript/api/word/word.range#word-word-range-pages-member)|Gets the collection of pages in the range.|
||[shapes](/javascript/api/word/word.range#word-word-range-shapes-member)|Gets the collection of shape objects anchored in the range, including both inline and floating shapes.|
|[Shape](/javascript/api/word/word.shape)|[body](/javascript/api/word/word.shape#word-word-shape-body-member)|Represents the body object of the shape.|
||[canvas](/javascript/api/word/word.shape#word-word-shape-canvas-member)|Gets the canvas associated with the shape.|
||[delete()](/javascript/api/word/word.shape#word-word-shape-delete-member(1))|Deletes the shape and its content.|
||[fill](/javascript/api/word/word.shape#word-word-shape-fill-member)|Returns the fill formatting of this shape.|
||[geometricShapeType](/javascript/api/word/word.shape#word-word-shape-geometricshapetype-member)|The geometric shape type of the shape.|
||[height](/javascript/api/word/word.shape#word-word-shape-height-member)|The height, in points, of the shape.|
||[id](/javascript/api/word/word.shape#word-word-shape-id-member)|Gets an integer that represents the shape identifier.|
||[isChild](/javascript/api/word/word.shape#word-word-shape-ischild-member)|Check whether this shape is a child of a group shape or a canvas shape.|
||[left](/javascript/api/word/word.shape#word-word-shape-left-member)|The distance, in points, from the left side of the shape to the horizontal relative position, see Word.RelativeHorizontalPosition.|
||[moveHorizontally(distance: number)](/javascript/api/word/word.shape#word-word-shape-movehorizontally-member(1))|Moves the shape horizontally by the number of points.|
||[moveVertically(distance: number)](/javascript/api/word/word.shape#word-word-shape-movevertically-member(1))|Moves the shape vertically by the number of points.|
||[name](/javascript/api/word/word.shape#word-word-shape-name-member)|The name of the shape.|
||[parentCanvas](/javascript/api/word/word.shape#word-word-shape-parentcanvas-member)|Gets the top-level parent canvas shape of this child shape.|
||[parentGroup](/javascript/api/word/word.shape#word-word-shape-parentgroup-member)|Gets the top-level parent group shape of this child shape.|
||[relativeHorizontalPosition](/javascript/api/word/word.shape#word-word-shape-relativehorizontalposition-member)|The relative horizontal position of the shape.|
||[relativeVerticalPosition](/javascript/api/word/word.shape#word-word-shape-relativeverticalposition-member)|The relative vertical position of the shape.|
||[select(selectMultipleShapes?: boolean)](/javascript/api/word/word.shape#word-word-shape-select-member(1))|Selects the shape.|
||[shapeGroup](/javascript/api/word/word.shape#word-word-shape-shapegroup-member)|Gets the shape group associated with the shape.|
||[textFrame](/javascript/api/word/word.shape#word-word-shape-textframe-member)|Gets the text frame object of the shape.|
||[top](/javascript/api/word/word.shape#word-word-shape-top-member)|The distance, in points, from the top edge of the shape to the vertical relative position, see Word.RelativeVerticalPosition.|
||[type](/javascript/api/word/word.shape#word-word-shape-type-member)|Gets the shape type.|
||[width](/javascript/api/word/word.shape#word-word-shape-width-member)|The width, in points, of the shape.|
|[ShapeCollection](/javascript/api/word/word.shapecollection)|[getByGeometricTypes(types: Word.GeometricShapeType[])](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbygeometrictypes-member(1))|Gets the shapes that have the specified geometric types.|
||[getById(id: number)](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbyid-member(1))|Gets a shape by its identifier.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbyidornullobject-member(1))|Gets a shape by its identifier.|
||[getByIds(ids: number[])](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbyids-member(1))|Gets the shapes by the identifiers.|
||[getByNames(names: string[])](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbynames-member(1))|Gets the shapes that have the specified names.|
||[getByTypes(types: Word.ShapeType[])](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbytypes-member(1))|Gets the shapes that have the specified types.|
||[getFirst()](/javascript/api/word/word.shapecollection#word-word-shapecollection-getfirst-member(1))|Gets the first shape in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.shapecollection#word-word-shapecollection-getfirstornullobject-member(1))|Gets the first shape in this collection.|
||[group()](/javascript/api/word/word.shapecollection#word-word-shapecollection-group-member(1))|Groups floating shapes in this collection, inline shapes will be skipped.|
||[items](/javascript/api/word/word.shapecollection#word-word-shapecollection-items-member)|Gets the loaded child items in this collection.|
|[ShapeFill](/javascript/api/word/word.shapefill)|[backgroundColor](/javascript/api/word/word.shapefill#word-word-shapefill-backgroundcolor-member)|Specifies the shape fill background color.|
||[clear()](/javascript/api/word/word.shapefill#word-word-shapefill-clear-member(1))|Clears the fill formatting of this shape and set it to `Word.ShapeFillType.NoFill`;|
||[foregroundColor](/javascript/api/word/word.shapefill#word-word-shapefill-foregroundcolor-member)|Specifies the shape fill foreground color.|
||[setSolidColor(color: string)](/javascript/api/word/word.shapefill#word-word-shapefill-setsolidcolor-member(1))|Sets the fill formatting of the shape to a uniform color.|
||[transparency](/javascript/api/word/word.shapefill#word-word-shapefill-transparency-member)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
||[type](/javascript/api/word/word.shapefill#word-word-shapefill-type-member)|Returns the fill type of the shape.|
|[ShapeGroup](/javascript/api/word/word.shapegroup)|[id](/javascript/api/word/word.shapegroup#word-word-shapegroup-id-member)|Gets an integer that represents the shape group identifier.|
||[shape](/javascript/api/word/word.shapegroup#word-word-shapegroup-shape-member)|Gets the Shape object associated with the group.|
||[shapes](/javascript/api/word/word.shapegroup#word-word-shapegroup-shapes-member)|Gets the collection of Shape objects.|
||[ungroup()](/javascript/api/word/word.shapegroup#word-word-shapegroup-ungroup-member(1))|Ungroups any grouped shapes in the specified shape group.|
|[Style](/javascript/api/word/word.style)|[description](/javascript/api/word/word.style#word-word-style-description-member)|Gets the description of the specified style.|
|[TextFrame](/javascript/api/word/word.textframe)|[autoSizeSetting](/javascript/api/word/word.textframe#word-word-textframe-autosizesetting-member)|The automatic sizing settings for the text frame.|
||[bottomMargin](/javascript/api/word/word.textframe#word-word-textframe-bottommargin-member)|Represents the bottom margin, in points, of the text frame.|
||[hasText](/javascript/api/word/word.textframe#word-word-textframe-hastext-member)|Specifies if the text frame contains text.|
||[leftMargin](/javascript/api/word/word.textframe#word-word-textframe-leftmargin-member)|Represents the left margin, in points, of the text frame.|
||[noTextRotation](/javascript/api/word/word.textframe#word-word-textframe-notextrotation-member)|Returns True if text in the text frame shouldn't rotate when the shape is rotated.|
||[orientation](/javascript/api/word/word.textframe#word-word-textframe-orientation-member)|Represents the angle to which the text is oriented for the text frame.|
||[rightMargin](/javascript/api/word/word.textframe#word-word-textframe-rightmargin-member)|Represents the right margin, in points, of the text frame.|
||[topMargin](/javascript/api/word/word.textframe#word-word-textframe-topmargin-member)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/word/word.textframe#word-word-textframe-verticalalignment-member)|Represents the vertical alignment of the text frame.|
||[wordWrap](/javascript/api/word/word.textframe#word-word-textframe-wordwrap-member)|Determines whether lines break automatically to fit text inside the shape.|
|[Window](/javascript/api/word/word.window)|[activePane](/javascript/api/word/word.window#word-word-window-activepane-member)|Gets the active pane in the window.|
||[panes](/javascript/api/word/word.window#word-word-window-panes-member)|Gets the collection of panes in the window.|
|[WindowCollection](/javascript/api/word/word.windowcollection)|[getFirst()](/javascript/api/word/word.windowcollection#word-word-windowcollection-getfirst-member(1))|Gets the first window in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.windowcollection#word-word-windowcollection-getfirstornullobject-member(1))|Gets the first window in this collection.|
||[getItem(index: number)](/javascript/api/word/word.windowcollection#word-word-windowcollection-getitem-member(1))|Gets a Window object by its index in the collection.|
||[items](/javascript/api/word/word.windowcollection#word-word-windowcollection-items-member)|Gets the loaded child items in this collection.|

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
