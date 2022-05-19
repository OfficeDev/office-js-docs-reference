| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#word-word-body-clear-member(1))|Clears the contents of the body object.|
||[endnotes](/javascript/api/word/word.body#word-word-body-endnotes-member)|Gets the collection of endnotes in the body.|
||[font](/javascript/api/word/word.body#word-word-body-font-member)|Gets the text format of the body.|
||[footnotes](/javascript/api/word/word.body#word-word-body-footnotes-member)|Gets the collection of footnotes in the body.|
||[getComments()](/javascript/api/word/word.body#word-word-body-getcomments-member(1))|Gets comments associated with the body.|
||[getHtml()](/javascript/api/word/word.body#word-word-body-gethtml-member(1))|Gets an HTML representation of the body object.|
||[getOoxml()](/javascript/api/word/word.body#word-word-body-getooxml-member(1))|Gets the OOXML (Office Open XML) representation of the body object.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#word-word-body-getrange-member(1))|Gets the whole body, or the starting or ending point of the body, as a range.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#word-word-body-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[ignorePunct](/javascript/api/word/word.body#word-word-body-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.body#word-word-body-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.body#word-word-body-inlinepictures-member)|Gets the collection of InlinePicture objects in the body.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertbreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.body#word-word-body-insertcontentcontrol-member(1))|Wraps the body object with a Rich Text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1))|Inserts a document into the body at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-inserthtml-member(1))|Inserts HTML at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertinlinepicturefrombase64-member(1))|Inserts a picture into the body at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertooxml-member(1))|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#word-word-body-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-inserttext-member(1))|Inserts text into the body at the specified location.|
||[lists](/javascript/api/word/word.body#word-word-body-lists-member)|Gets the collection of list objects in the body.|
||[matchCase](/javascript/api/word/word.body#word-word-body-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.body#word-word-body-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.body#word-word-body-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.body#word-word-body-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.body#word-word-body-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.body#word-word-body-paragraphs-member)|Gets the collection of paragraph objects in the body.|
||[parentBody](/javascript/api/word/word.body#word-word-body-parentbody-member)|Gets the parent body of the body.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#word-word-body-parentbodyornullobject-member)|Gets the parent body of the body.|
||[parentContentControl](/javascript/api/word/word.body#word-word-body-parentcontentcontrol-member)|Gets the content control that contains the body.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#word-word-body-parentcontentcontrolornullobject-member)|Gets the content control that contains the body.|
||[parentSection](/javascript/api/word/word.body#word-word-body-parentsection-member)|Gets the parent section of the body.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#word-word-body-parentsectionornullobject-member)|Gets the parent section of the body.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.body#word-word-body-search-member(1))|Performs a search with the specified SearchOptions on the scope of the body object.|
||[style](/javascript/api/word/word.body#word-word-body-style-member)|Gets or sets the style name for the body.|
||[styleBuiltIn](/javascript/api/word/word.body#word-word-body-stylebuiltin-member)|Gets or sets the built-in style name for the body.|
||[tables](/javascript/api/word/word.body#word-word-body-tables-member)|Gets the collection of table objects in the body.|
||[text](/javascript/api/word/word.body#word-word-body-text-member)|Gets the text of the body.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Gets the type of the body.|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#word-word-comment-authoremail-member)|Gets the email of the comment's author.|
||[authorName](/javascript/api/word/word.comment#word-word-comment-authorname-member)|Gets the name of the comment's author.|
||[content](/javascript/api/word/word.comment#word-word-comment-content-member)|Gets or sets the comment's content as plain text.|
||[contentRange](/javascript/api/word/word.comment#word-word-comment-contentrange-member)|Gets or sets the comment's content range.|
||[creationDate](/javascript/api/word/word.comment#word-word-comment-creationdate-member)|Gets the creation date of the comment.|
||[delete()](/javascript/api/word/word.comment#word-word-comment-delete-member(1))|Deletes the comment and its replies.|
||[getRange()](/javascript/api/word/word.comment#word-word-comment-getrange-member(1))|Gets the range in the main document where the comment is on.|
||[id](/javascript/api/word/word.comment#word-word-comment-id-member)|Gets the Id of the comment.|
||[replies](/javascript/api/word/word.comment#word-word-comment-replies-member)|Gets the collection of reply objects associated with the comment.|
||[reply(replyText: string)](/javascript/api/word/word.comment#word-word-comment-reply-member(1))|Adds a new reply to the end of the comment thread.|
||[resolved](/javascript/api/word/word.comment#word-word-comment-resolved-member)|Gets or sets the comment thread's status.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirst-member(1))|Gets the first comment in the collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirstornullobject-member(1))|Gets the first comment in the collection.|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#word-word-commentcollection-getitem-member(1))|Gets a comment object by its index in the collection.|
||[items](/javascript/api/word/word.commentcollection#word-word-commentcollection-items-member)|Gets the loaded child items in this collection.|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-bold-member)|Gets or sets a value that indicates whether the comment text is bold.|
||[hyperlink](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-hyperlink-member)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-inserttext-member(1))|Inserts text into at the specified location.|
||[isEmpty](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-isempty-member)|Checks whether the range length is zero.|
||[italic](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-italic-member)|Gets or sets a value that indicates whether the comment text is italicized.|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-strikethrough-member)|Gets or sets a value that indicates whether the comment text has a strikethrough.|
||[text](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-text-member)|Gets the text of the comment range.|
||[underline](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-underline-member)|Gets or sets a value that indicates the comment text's underline type.|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#word-word-commentreply-authoremail-member)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/word/word.commentreply#word-word-commentreply-authorname-member)|Gets the name of the comment reply's author.|
||[content](/javascript/api/word/word.commentreply#word-word-commentreply-content-member)|Gets or sets the comment reply's content.|
||[contentRange](/javascript/api/word/word.commentreply#word-word-commentreply-contentrange-member)|Gets or sets the commentReply's content range.|
||[creationDate](/javascript/api/word/word.commentreply#word-word-commentreply-creationdate-member)|Gets the creation date of the comment reply.|
||[delete()](/javascript/api/word/word.commentreply#word-word-commentreply-delete-member(1))|Deletes the comment reply.|
||[id](/javascript/api/word/word.commentreply#word-word-commentreply-id-member)|Gets the Id of the comment reply.|
||[parentComment](/javascript/api/word/word.commentreply#word-word-commentreply-parentcomment-member)|Gets the parent comment of this reply.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirst-member(1))|Gets the first comment reply in the collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirstornullobject-member(1))|Gets the first comment reply in the collection.|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getitem-member(1))|Gets a comment reply object by its index in the collection.|
||[items](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-items-member)|Gets the loaded child items in this collection.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[appearance](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-appearance-member)|Gets or sets the appearance of the content control.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotdelete-member)|Gets or sets a value that indicates whether the user can delete the content control.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotedit-member)|Gets or sets a value that indicates whether the user can edit the contents of the content control.|
||[clear()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-clear-member(1))|Clears the contents of the content control.|
||[color](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-color-member)|Gets or sets the color of the content control.|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-delete-member(1))|Deletes the content control and its content.|
||[endnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|Gets the collection of endnotes in the contentcontrol.|
||[font](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-font-member)|Gets the text format of the content control.|
||[footnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|Gets the collection of footnotes in the contentcontrol.|
||[getComments()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcomments-member(1))|Gets comments associated with the body.|
||[getHtml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gethtml-member(1))|Gets an HTML representation of the content control object.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getooxml-member(1))|Gets the Office Open XML (OOXML) representation of the content control object.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getrange-member(1))|Gets the whole content control, or the starting or ending point of the content control, as a range.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gettextranges-member(1))|Gets the text ranges in the content control by using punctuation marks and/or other ending marks.|
||[id](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-id-member)|Gets an integer that represents the content control identifier.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inlinepictures-member)|Gets the collection of inlinePicture objects in the content control.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertbreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertfilefrombase64-member(1))|Inserts a document into the content control at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserthtml-member(1))|Inserts HTML into the content control at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertinlinepicturefrombase64-member(1))|Inserts an inline picture into the content control at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertooxml-member(1))|Inserts OOXML into the content control at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttable-member(1))|Inserts a table with the specified number of rows and columns into, or next to, a content control.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttext-member(1))|Inserts text into the content control at the specified location.|
||[lists](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-lists-member)|Gets the collection of list objects in the content control.|
||[matchCase](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-paragraphs-member)|Gets the collection of paragraph objects in the content control.|
||[parentBody](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentbody-member)|Gets the parent body of the content control.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrol-member)|Gets the content control that contains the content control.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrolornullobject-member)|Gets the content control that contains the content control.|
||[parentTable](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttable-member)|Gets the table that contains the content control.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecell-member)|Gets the table cell that contains the content control.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecellornullobject-member)|Gets the table cell that contains the content control.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttableornullobject-member)|Gets the table that contains the content control.|
||[placeholderText](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-placeholdertext-member)|Gets or sets the placeholder text of the content control.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-removewhenedited-member)|Gets or sets a value that indicates whether the content control is removed after it is edited.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-search-member(1))|Performs a search with the specified SearchOptions on the scope of the content control object.|
||[style](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-style-member)|Gets or sets the style name for the content control.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-stylebuiltin-member)|Gets or sets the built-in style name for the content control.|
||[subtype](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-subtype-member)|Gets the content control subtype.|
||[tables](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tables-member)|Gets the collection of table objects in the content control.|
||[tag](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tag-member)|Gets or sets a tag to identify a content control.|
||[text](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-text-member)|Gets the text of the content control.|
||[title](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-title-member)|Gets or sets the title for a content control.|
||[type](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-type-member)|Gets the content control type.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#word-word-document-changetrackingmode-member)|Gets or sets the ChangeTracking mode.|
||[getEndnoteBody()](/javascript/api/word/word.document#word-word-document-getendnotebody-member(1))|Gets the document's endnotes in a single body.|
||[getFootnoteBody()](/javascript/api/word/word.document#word-word-document-getfootnotebody-member(1))|Gets the document's footnotes in a single body.|
||[getSelection()](/javascript/api/word/word.document#word-word-document-getselection-member(1))|Gets the current selection of the document.|
||[ignorePunct](/javascript/api/word/word.document#word-word-document-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.document#word-word-document-ignorespace-member)||
||[matchCase](/javascript/api/word/word.document#word-word-document-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.document#word-word-document-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.document#word-word-document-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.document#word-word-document-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.document#word-word-document-matchwildcards-member)||
||[save()](/javascript/api/word/word.document#word-word-document-save-member(1))|Saves the document.|
||[saved](/javascript/api/word/word.document#word-word-document-saved-member)|Indicates whether the changes in the document have been saved.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.document#word-word-document-search-member(1))|Performs a search with the specified search options on the scope of the whole document.|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#word-word-noteitem-body-member)|Represents the body object of the note item.|
||[delete()](/javascript/api/word/word.noteitem#word-word-noteitem-delete-member(1))|Deletes the note item.|
||[getNext()](/javascript/api/word/word.noteitem#word-word-noteitem-getnext-member(1))|Gets the next note item of the same type.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#word-word-noteitem-getnextornullobject-member(1))|Gets the next note item of the same type.|
||[reference](/javascript/api/word/word.noteitem#word-word-noteitem-reference-member)|Represents a footnote or endnote reference in the main document.|
||[type](/javascript/api/word/word.noteitem#word-word-noteitem-type-member)|Represents the note item type: footnote or endnote.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirst-member(1))|Gets the first note item in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirstornullobject-member(1))|Gets the first note item in this collection.|
||[items](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[alignment](/javascript/api/word/word.paragraph#word-word-paragraph-alignment-member)|Gets or sets the alignment for a paragraph.|
||[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#word-word-paragraph-attachtolist-member(1))|Lets the paragraph join an existing list at the specified level.|
||[clear()](/javascript/api/word/word.paragraph#word-word-paragraph-clear-member(1))|Clears the contents of the paragraph object.|
||[delete()](/javascript/api/word/word.paragraph#word-word-paragraph-delete-member(1))|Deletes the paragraph and its content from the document.|
||[detachFromList()](/javascript/api/word/word.paragraph#word-word-paragraph-detachfromlist-member(1))|Moves this paragraph out of its list, if the paragraph is a list item.|
||[endnotes](/javascript/api/word/word.paragraph#word-word-paragraph-endnotes-member)|Gets the collection of endnotes in the paragraph.|
||[firstLineIndent](/javascript/api/word/word.paragraph#word-word-paragraph-firstlineindent-member)|Gets or sets the value, in points, for a first line or hanging indent.|
||[font](/javascript/api/word/word.paragraph#word-word-paragraph-font-member)|Gets the text format of the paragraph.|
||[footnotes](/javascript/api/word/word.paragraph#word-word-paragraph-footnotes-member)|Gets the collection of footnotes in the paragraph.|
||[getComments()](/javascript/api/word/word.paragraph#word-word-paragraph-getcomments-member(1))|Gets comments associated with the paragraph.|
||[getHtml()](/javascript/api/word/word.paragraph#word-word-paragraph-gethtml-member(1))|Gets an HTML representation of the paragraph object.|
||[getNext()](/javascript/api/word/word.paragraph#word-word-paragraph-getnext-member(1))|Gets the next paragraph.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getnextornullobject-member(1))|Gets the next paragraph.|
||[getOoxml()](/javascript/api/word/word.paragraph#word-word-paragraph-getooxml-member(1))|Gets the Office Open XML (OOXML) representation of the paragraph object.|
||[getPrevious()](/javascript/api/word/word.paragraph#word-word-paragraph-getprevious-member(1))|Gets the previous paragraph.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getpreviousornullobject-member(1))|Gets the previous paragraph.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-getrange-member(1))|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#word-word-paragraph-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-gettextranges-member(1))|Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.|
||[ignorePunct](/javascript/api/word/word.paragraph#word-word-paragraph-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.paragraph#word-word-paragraph-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.paragraph#word-word-paragraph-inlinepictures-member)|Gets the collection of InlinePicture objects in the paragraph.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertbreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.paragraph#word-word-paragraph-insertcontentcontrol-member(1))|Wraps the paragraph object with a rich text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertfilefrombase64-member(1))|Inserts a document into the paragraph at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-inserthtml-member(1))|Inserts HTML into the paragraph at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertinlinepicturefrombase64-member(1))|Inserts a picture into the paragraph at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertooxml-member(1))|Inserts OOXML into the paragraph at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#word-word-paragraph-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-inserttext-member(1))|Inserts text into the paragraph at the specified location.|
||[isLastParagraph](/javascript/api/word/word.paragraph#word-word-paragraph-islastparagraph-member)|Indicates the paragraph is the last one inside its parent body.|
||[isListItem](/javascript/api/word/word.paragraph#word-word-paragraph-islistitem-member)|Checks whether the paragraph is a list item.|
||[leftIndent](/javascript/api/word/word.paragraph#word-word-paragraph-leftindent-member)|Gets or sets the left indent value, in points, for the paragraph.|
||[lineSpacing](/javascript/api/word/word.paragraph#word-word-paragraph-linespacing-member)|Gets or sets the line spacing, in points, for the specified paragraph.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitafter-member)|Gets or sets the amount of spacing, in grid lines, after the paragraph.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitbefore-member)|Gets or sets the amount of spacing, in grid lines, before the paragraph.|
||[list](/javascript/api/word/word.paragraph#word-word-paragraph-list-member)|Gets the List to which this paragraph belongs.|
||[listItem](/javascript/api/word/word.paragraph#word-word-paragraph-listitem-member)|Gets the ListItem for the paragraph.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listitemornullobject-member)|Gets the ListItem for the paragraph.|
||[listOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listornullobject-member)|Gets the List to which this paragraph belongs.|
||[matchCase](/javascript/api/word/word.paragraph#word-word-paragraph-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.paragraph#word-word-paragraph-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.paragraph#word-word-paragraph-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.paragraph#word-word-paragraph-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.paragraph#word-word-paragraph-matchwildcards-member)||
||[outlineLevel](/javascript/api/word/word.paragraph#word-word-paragraph-outlinelevel-member)|Gets or sets the outline level for the paragraph.|
||[parentBody](/javascript/api/word/word.paragraph#word-word-paragraph-parentbody-member)|Gets the parent body of the paragraph.|
||[parentContentControl](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrol-member)|Gets the content control that contains the paragraph.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrolornullobject-member)|Gets the content control that contains the paragraph.|
||[parentTable](/javascript/api/word/word.paragraph#word-word-paragraph-parenttable-member)|Gets the table that contains the paragraph.|
||[parentTableCell](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecell-member)|Gets the table cell that contains the paragraph.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecellornullobject-member)|Gets the table cell that contains the paragraph.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttableornullobject-member)|Gets the table that contains the paragraph.|
||[rightIndent](/javascript/api/word/word.paragraph#word-word-paragraph-rightindent-member)|Gets or sets the right indent value, in points, for the paragraph.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.paragraph#word-word-paragraph-search-member(1))|Performs a search with the specified SearchOptions on the scope of the paragraph object.|
||[spaceAfter](/javascript/api/word/word.paragraph#word-word-paragraph-spaceafter-member)|Gets or sets the spacing, in points, after the paragraph.|
||[spaceBefore](/javascript/api/word/word.paragraph#word-word-paragraph-spacebefore-member)|Gets or sets the spacing, in points, before the paragraph.|
||[style](/javascript/api/word/word.paragraph#word-word-paragraph-style-member)|Gets or sets the style name for the paragraph.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#word-word-paragraph-stylebuiltin-member)|Gets or sets the built-in style name for the paragraph.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#word-word-paragraph-tablenestinglevel-member)|Gets the level of the paragraph's table.|
||[text](/javascript/api/word/word.paragraph#word-word-paragraph-text-member)|Gets the text of the paragraph.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#word-word-range-clear-member(1))|Clears the contents of the range object.|
||[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-comparelocationwith-member(1))|Compares this range's location with another range's location.|
||[delete()](/javascript/api/word/word.range#word-word-range-delete-member(1))|Deletes the range and its content from the document.|
||[endnotes](/javascript/api/word/word.range#word-word-range-endnotes-member)|Gets the collection of endnotes in the range.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandto-member(1))|Returns a new range that extends from this range in either direction to cover another range.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandtoornullobject-member(1))|Returns a new range that extends from this range in either direction to cover another range.|
||[font](/javascript/api/word/word.range#word-word-range-font-member)|Gets the text format of the range.|
||[footnotes](/javascript/api/word/word.range#word-word-range-footnotes-member)|Gets the collection of footnotes in the range.|
||[getComments()](/javascript/api/word/word.range#word-word-range-getcomments-member(1))|Gets comments associated with the range.|
||[getHtml()](/javascript/api/word/word.range#word-word-range-gethtml-member(1))|Gets an HTML representation of the range object.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#word-word-range-gethyperlinkranges-member(1))|Gets hyperlink child ranges within the range.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrange-member(1))|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrangeornullobject-member(1))|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getOoxml()](/javascript/api/word/word.range#word-word-range-getooxml-member(1))|Gets the OOXML representation of the range object.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#word-word-range-getrange-member(1))|Clones the range, or gets the starting or ending point of the range as a new range.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#word-word-range-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-gettextranges-member(1))|Gets the text child ranges in the range by using punctuation marks and/or other ending marks.|
||[hyperlink](/javascript/api/word/word.range#word-word-range-hyperlink-member)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[ignorePunct](/javascript/api/word/word.range#word-word-range-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.range#word-word-range-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.range#word-word-range-inlinepictures-member)|Gets the collection of inline picture objects in the range.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertbreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertComment(commentText: string)](/javascript/api/word/word.range#word-word-range-insertcomment-member(1))|Insert a comment on the range.|
||[insertContentControl()](/javascript/api/word/word.range#word-word-range-insertcontentcontrol-member(1))|Wraps the range object with a rich text content control.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertendnote-member(1))|Inserts an endnote.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertfilefrombase64-member(1))|Inserts a document at the specified location.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertfootnote-member(1))|Inserts a footnote.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-inserthtml-member(1))|Inserts HTML at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertinlinepicturefrombase64-member(1))|Inserts a picture at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertooxml-member(1))|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#word-word-range-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-inserttext-member(1))|Inserts text at the specified location.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwith-member(1))|Returns a new range as the intersection of this range with another range.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwithornullobject-member(1))|Returns a new range as the intersection of this range with another range.|
||[isEmpty](/javascript/api/word/word.range#word-word-range-isempty-member)|Checks whether the range length is zero.|
||[lists](/javascript/api/word/word.range#word-word-range-lists-member)|Gets the collection of list objects in the range.|
||[matchCase](/javascript/api/word/word.range#word-word-range-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.range#word-word-range-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.range#word-word-range-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.range#word-word-range-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.range#word-word-range-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.range#word-word-range-paragraphs-member)|Gets the collection of paragraph objects in the range.|
||[parentBody](/javascript/api/word/word.range#word-word-range-parentbody-member)|Gets the parent body of the range.|
||[parentContentControl](/javascript/api/word/word.range#word-word-range-parentcontentcontrol-member)|Gets the content control that contains the range.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#word-word-range-parentcontentcontrolornullobject-member)|Gets the content control that contains the range.|
||[parentTable](/javascript/api/word/word.range#word-word-range-parenttable-member)|Gets the table that contains the range.|
||[parentTableCell](/javascript/api/word/word.range#word-word-range-parenttablecell-member)|Gets the table cell that contains the range.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#word-word-range-parenttablecellornullobject-member)|Gets the table cell that contains the range.|
||[parentTableOrNullObject](/javascript/api/word/word.range#word-word-range-parenttableornullobject-member)|Gets the table that contains the range.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.range#word-word-range-search-member(1))|Performs a search with the specified SearchOptions on the scope of the range object.|
||[style](/javascript/api/word/word.range#word-word-range-style-member)|Gets or sets the style name for the range.|
||[styleBuiltIn](/javascript/api/word/word.range#word-word-range-stylebuiltin-member)|Gets or sets the built-in style name for the range.|
||[tables](/javascript/api/word/word.range#word-word-range-tables-member)|Gets the collection of table objects in the range.|
||[text](/javascript/api/word/word.range#word-word-range-text-member)|Gets the text of the range.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addcolumns-member(1))|Adds columns to the start or end of the table, using the first or last existing column as a template.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addrows-member(1))|Adds rows to the start or end of the table, using the first or last existing row as a template.|
||[alignment](/javascript/api/word/word.table#word-word-table-alignment-member)|Gets or sets the alignment of the table against the page column.|
||[autoFitWindow()](/javascript/api/word/word.table#word-word-table-autofitwindow-member(1))|Autofits the table columns to the width of the window.|
||[clear()](/javascript/api/word/word.table#word-word-table-clear-member(1))|Clears the contents of the table.|
||[delete()](/javascript/api/word/word.table#word-word-table-delete-member(1))|Deletes the entire table.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#word-word-table-deletecolumns-member(1))|Deletes specific columns.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#word-word-table-deleterows-member(1))|Deletes specific rows.|
||[distributeColumns()](/javascript/api/word/word.table#word-word-table-distributecolumns-member(1))|Distributes the column widths evenly.|
||[endnotes](/javascript/api/word/word.table#word-word-table-endnotes-member)|Gets the collection of endnotes in the table.|
||[font](/javascript/api/word/word.table#word-word-table-font-member)|Gets the font.|
||[footnotes](/javascript/api/word/word.table#word-word-table-footnotes-member)|Gets the collection of footnotes in the table.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#word-word-table-getborder-member(1))|Gets the border style for the specified border.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcell-member(1))|Gets the table cell at a specified row and column.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcellornullobject-member(1))|Gets the table cell at a specified row and column.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#word-word-table-getcellpadding-member(1))|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.table#word-word-table-getnext-member(1))|Gets the next table.|
||[getNextOrNullObject()](/javascript/api/word/word.table#word-word-table-getnextornullobject-member(1))|Gets the next table.|
||[getParagraphAfter()](/javascript/api/word/word.table#word-word-table-getparagraphafter-member(1))|Gets the paragraph after the table.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphafterornullobject-member(1))|Gets the paragraph after the table.|
||[getParagraphBefore()](/javascript/api/word/word.table#word-word-table-getparagraphbefore-member(1))|Gets the paragraph before the table.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphbeforeornullobject-member(1))|Gets the paragraph before the table.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#word-word-table-getrange-member(1))|Gets the range that contains this table, or the range at the start or end of the table.|
||[headerRowCount](/javascript/api/word/word.table#word-word-table-headerrowcount-member)|Gets and sets the number of header rows.|
||[horizontalAlignment](/javascript/api/word/word.table#word-word-table-horizontalalignment-member)|Gets and sets the horizontal alignment of every cell in the table.|
||[ignorePunct](/javascript/api/word/word.table#word-word-table-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.table#word-word-table-ignorespace-member)||
||[insertContentControl()](/javascript/api/word/word.table#word-word-table-insertcontentcontrol-member(1))|Inserts a content control on the table.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#word-word-table-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#word-word-table-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[isUniform](/javascript/api/word/word.table#word-word-table-isuniform-member)|Indicates whether all of the table rows are uniform.|
||[matchCase](/javascript/api/word/word.table#word-word-table-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.table#word-word-table-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.table#word-word-table-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.table#word-word-table-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.table#word-word-table-matchwildcards-member)||
||[nestingLevel](/javascript/api/word/word.table#word-word-table-nestinglevel-member)|Gets the nesting level of the table.|
||[parentBody](/javascript/api/word/word.table#word-word-table-parentbody-member)|Gets the parent body of the table.|
||[parentContentControl](/javascript/api/word/word.table#word-word-table-parentcontentcontrol-member)|Gets the content control that contains the table.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#word-word-table-parentcontentcontrolornullobject-member)|Gets the content control that contains the table.|
||[parentTable](/javascript/api/word/word.table#word-word-table-parenttable-member)|Gets the table that contains this table.|
||[parentTableCell](/javascript/api/word/word.table#word-word-table-parenttablecell-member)|Gets the table cell that contains this table.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#word-word-table-parenttablecellornullobject-member)|Gets the table cell that contains this table.|
||[parentTableOrNullObject](/javascript/api/word/word.table#word-word-table-parenttableornullobject-member)|Gets the table that contains this table.|
||[rowCount](/javascript/api/word/word.table#word-word-table-rowcount-member)|Gets the number of rows in the table.|
||[rows](/javascript/api/word/word.table#word-word-table-rows-member)|Gets all of the table rows.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.table#word-word-table-search-member(1))|Performs a search with the specified SearchOptions on the scope of the table object.|
||[shadingColor](/javascript/api/word/word.table#word-word-table-shadingcolor-member)|Gets and sets the shading color.|
||[style](/javascript/api/word/word.table#word-word-table-style-member)|Gets or sets the style name for the table.|
||[styleBandedColumns](/javascript/api/word/word.table#word-word-table-stylebandedcolumns-member)|Gets and sets whether the table has banded columns.|
||[styleBandedRows](/javascript/api/word/word.table#word-word-table-stylebandedrows-member)|Gets and sets whether the table has banded rows.|
||[styleBuiltIn](/javascript/api/word/word.table#word-word-table-stylebuiltin-member)|Gets or sets the built-in style name for the table.|
||[styleFirstColumn](/javascript/api/word/word.table#word-word-table-stylefirstcolumn-member)|Gets and sets whether the table has a first column with a special style.|
||[styleLastColumn](/javascript/api/word/word.table#word-word-table-stylelastcolumn-member)|Gets and sets whether the table has a last column with a special style.|
||[styleTotalRow](/javascript/api/word/word.table#word-word-table-styletotalrow-member)|Gets and sets whether the table has a total (last) row with a special style.|
||[tables](/javascript/api/word/word.table#word-word-table-tables-member)|Gets the child tables nested one level deeper.|
||[values](/javascript/api/word/word.table#word-word-table-values-member)|Gets and sets the text values in the table, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.table#word-word-table-verticalalignment-member)|Gets and sets the vertical alignment of every cell in the table.|
||[width](/javascript/api/word/word.table#word-word-table-width-member)|Gets and sets the width of the table in points.|
|[TableRow](/javascript/api/word/word.tablerow)|[cellCount](/javascript/api/word/word.tablerow#word-word-tablerow-cellcount-member)|Gets the number of cells in the row.|
||[clear()](/javascript/api/word/word.tablerow#word-word-tablerow-clear-member(1))|Clears the contents of the row.|
||[delete()](/javascript/api/word/word.tablerow#word-word-tablerow-delete-member(1))|Deletes the entire row.|
||[endnotes](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|Gets the collection of endnotes in the table row.|
||[font](/javascript/api/word/word.tablerow#word-word-tablerow-font-member)|Gets the font.|
||[footnotes](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|Gets the collection of footnotes in the table row.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getborder-member(1))|Gets the border style of the cells in the row.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getcellpadding-member(1))|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablerow#word-word-tablerow-getnext-member(1))|Gets the next row.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#word-word-tablerow-getnextornullobject-member(1))|Gets the next row.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-horizontalalignment-member)|Gets and sets the horizontal alignment of every cell in the row.|
||[ignorePunct](/javascript/api/word/word.tablerow#word-word-tablerow-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.tablerow#word-word-tablerow-ignorespace-member)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#word-word-tablerow-insertrows-member(1))|Inserts rows using this row as a template.|
||[isHeader](/javascript/api/word/word.tablerow#word-word-tablerow-isheader-member)|Checks whether the row is a header row.|
||[matchCase](/javascript/api/word/word.tablerow#word-word-tablerow-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.tablerow#word-word-tablerow-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.tablerow#word-word-tablerow-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.tablerow#word-word-tablerow-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.tablerow#word-word-tablerow-matchwildcards-member)||
||[parentTable](/javascript/api/word/word.tablerow#word-word-tablerow-parenttable-member)|Gets parent table.|
||[preferredHeight](/javascript/api/word/word.tablerow#word-word-tablerow-preferredheight-member)|Gets and sets the preferred height of the row in points.|
||[rowIndex](/javascript/api/word/word.tablerow#word-word-tablerow-rowindex-member)|Gets the index of the row in its parent table.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.tablerow#word-word-tablerow-search-member(1))|Performs a search with the specified SearchOptions on the scope of the row.|
||[shadingColor](/javascript/api/word/word.tablerow#word-word-tablerow-shadingcolor-member)|Gets and sets the shading color.|
||[values](/javascript/api/word/word.tablerow#word-word-tablerow-values-member)|Gets and sets the text values in the row, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-verticalalignment-member)|Gets and sets the vertical alignment of the cells in the row.|
