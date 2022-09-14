| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#word-word-body-endnotes-member)|Gets the collection of endnotes in the body.|
||[footnotes](/javascript/api/word/word.body#word-word-body-footnotes-member)|Gets the collection of footnotes in the body.|
||[getComments()](/javascript/api/word/word.body#word-word-body-getcomments-member(1))|Gets comments associated with the body.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#word-word-body-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
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
||[items](/javascript/api/word/word.commentcollection#word-word-commentcollection-items-member)|Gets the loaded child items in this collection.|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-bold-member)|Gets or sets a value that indicates whether the comment text is bold.|
||[hyperlink](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-hyperlink-member)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertText(text: string, insertLocation: Word.InsertLocation \| "Replace" \| "Start" \| "End" \| "Before" \| "After")](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-inserttext-member(1))|Inserts text into at the specified location.|
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
||[items](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-items-member)|Gets the loaded child items in this collection.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|Gets the collection of endnotes in the contentcontrol.|
||[footnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|Gets the collection of footnotes in the contentcontrol.|
||[getComments()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcomments-member(1))|Gets comments associated with the content control.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#word-word-document-changetrackingmode-member)|Gets or sets the ChangeTracking mode.|
||[getEndnoteBody()](/javascript/api/word/word.document#word-word-document-getendnotebody-member(1))|Gets the document's endnotes in a single body.|
||[getFootnoteBody()](/javascript/api/word/word.document#word-word-document-getfootnotebody-member(1))|Gets the document's footnotes in a single body.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.document#word-word-document-search-member(1))|Performs a search with the specified search options on the scope of the whole document.|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#word-word-noteitem-body-member)|Represents the body object of the note item.|
||[delete()](/javascript/api/word/word.noteitem#word-word-noteitem-delete-member(1))|Deletes the note item.|
||[getNext()](/javascript/api/word/word.noteitem#word-word-noteitem-getnext-member(1))|Gets the next note item of the same type.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#word-word-noteitem-getnextornullobject-member(1))|Gets the next note item of the same type.|
||[reference](/javascript/api/word/word.noteitem#word-word-noteitem-reference-member)|Represents a footnote or endnote reference in the main document.|
||[type](/javascript/api/word/word.noteitem#word-word-noteitem-type-member)|Represents the note item type: footnote or endnote.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirst-member(1))|Gets the first note item in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirstornullobject-member(1))|Gets the first note item in this collection.|
||[items](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#word-word-paragraph-endnotes-member)|Gets the collection of endnotes in the paragraph.|
||[footnotes](/javascript/api/word/word.paragraph#word-word-paragraph-footnotes-member)|Gets the collection of footnotes in the paragraph.|
||[getComments()](/javascript/api/word/word.paragraph#word-word-paragraph-getcomments-member(1))|Gets comments associated with the paragraph.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#word-word-paragraph-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#word-word-range-endnotes-member)|Gets the collection of endnotes in the range.|
||[footnotes](/javascript/api/word/word.range#word-word-range-footnotes-member)|Gets the collection of footnotes in the range.|
||[getComments()](/javascript/api/word/word.range#word-word-range-getcomments-member(1))|Gets comments associated with the range.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#word-word-range-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[insertComment(commentText: string)](/javascript/api/word/word.range#word-word-range-insertcomment-member(1))|Insert a comment on the range.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertendnote-member(1))|Inserts an endnote.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertfootnote-member(1))|Inserts a footnote.|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#word-word-table-endnotes-member)|Gets the collection of endnotes in the table.|
||[footnotes](/javascript/api/word/word.table#word-word-table-footnotes-member)|Gets the collection of footnotes in the table.|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|Gets the collection of endnotes in the table row.|
||[footnotes](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|Gets the collection of footnotes in the table row.|
