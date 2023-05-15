| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[fields](/javascript/api/word/word.body#word-word-body-fields-member)|Gets the collection of field objects in the body.|
||[getComments()](/javascript/api/word/word.body#word-word-body-getcomments-member(1))|Gets comments associated with the body.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#word-word-body-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#word-word-comment-authoremail-member)|Gets the email of the comment's author.|
||[authorName](/javascript/api/word/word.comment#word-word-comment-authorname-member)|Gets the name of the comment's author.|
||[content](/javascript/api/word/word.comment#word-word-comment-content-member)|Specifies the comment's content as plain text.|
||[contentRange](/javascript/api/word/word.comment#word-word-comment-contentrange-member)|Specifies the comment's content range.|
||[creationDate](/javascript/api/word/word.comment#word-word-comment-creationdate-member)|Gets the creation date of the comment.|
||[delete()](/javascript/api/word/word.comment#word-word-comment-delete-member(1))|Deletes the comment and its replies.|
||[getRange()](/javascript/api/word/word.comment#word-word-comment-getrange-member(1))|Gets the range in the main document where the comment is on.|
||[id](/javascript/api/word/word.comment#word-word-comment-id-member)|Gets the ID of the comment.|
||[replies](/javascript/api/word/word.comment#word-word-comment-replies-member)|Gets the collection of reply objects associated with the comment.|
||[reply(replyText: string)](/javascript/api/word/word.comment#word-word-comment-reply-member(1))|Adds a new reply to the end of the comment thread.|
||[resolved](/javascript/api/word/word.comment#word-word-comment-resolved-member)|Specifies the comment thread's status.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirst-member(1))|Gets the first comment in the collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirstornullobject-member(1))|Gets the first comment in the collection.|
||[items](/javascript/api/word/word.commentcollection#word-word-commentcollection-items-member)|Gets the loaded child items in this collection.|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-bold-member)|Specifies a value that indicates whether the comment text is bold.|
||[hyperlink](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-hyperlink-member)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertText(text: string, insertLocation: Word.InsertLocation \| "Replace" \| "Start" \| "End" \| "Before" \| "After")](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-inserttext-member(1))|Inserts text into at the specified location.|
||[isEmpty](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-isempty-member)|Checks whether the range length is zero.|
||[italic](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-italic-member)|Specifies a value that indicates whether the comment text is italicized.|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-strikethrough-member)|Specifies a value that indicates whether the comment text has a strikethrough.|
||[text](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-text-member)|Gets the text of the comment range.|
||[underline](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-underline-member)|Specifies a value that indicates the comment text's underline type.|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#word-word-commentreply-authoremail-member)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/word/word.commentreply#word-word-commentreply-authorname-member)|Gets the name of the comment reply's author.|
||[content](/javascript/api/word/word.commentreply#word-word-commentreply-content-member)|Specifies the comment reply's content.|
||[contentRange](/javascript/api/word/word.commentreply#word-word-commentreply-contentrange-member)|Specifies the commentReply's content range.|
||[creationDate](/javascript/api/word/word.commentreply#word-word-commentreply-creationdate-member)|Gets the creation date of the comment reply.|
||[delete()](/javascript/api/word/word.commentreply#word-word-commentreply-delete-member(1))|Deletes the comment reply.|
||[id](/javascript/api/word/word.commentreply#word-word-commentreply-id-member)|Gets the ID of the comment reply.|
||[parentComment](/javascript/api/word/word.commentreply#word-word-commentreply-parentcomment-member)|Gets the parent comment of this reply.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirst-member(1))|Gets the first comment reply in the collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirstornullobject-member(1))|Gets the first comment reply in the collection.|
||[items](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-items-member)|Gets the loaded child items in this collection.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[fields](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-fields-member)|Gets the collection of field objects in the content control.|
||[getComments()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcomments-member(1))|Gets comments associated with the content control.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-delete-member(1))|Deletes the custom XML part.|
||[deleteAttribute(xpath: string, namespaceMappings: {            [key: string]: string        }, name: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteattribute-member(1))|Deletes an attribute with the given name from the element identified by xpath.|
||[deleteElement(xpath: string, namespaceMappings: {            [key: string]: string        })](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteelement-member(1))|Deletes the element identified by xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-getxml-member(1))|Gets the full XML content of the custom XML part.|
||[id](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-id-member)|Gets the ID of the custom XML part.|
||[insertAttribute(xpath: string, namespaceMappings: {            [key: string]: string        }, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertattribute-member(1))|Inserts an attribute with the given name and value to the element identified by xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: {            [key: string]: string        }, index?: number)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertelement-member(1))|Inserts the given XML under the parent element identified by xpath at child position index.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-namespaceuri-member)|Gets the namespace URI of the custom XML part.|
||[query(xpath: string, namespaceMappings: {            [key: string]: string        })](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-query-member(1))|Queries the XML content of the custom XML part.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-setxml-member(1))|Sets the full XML content of the custom XML part.|
||[updateAttribute(xpath: string, namespaceMappings: {            [key: string]: string        }, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateattribute-member(1))|Updates the value of an attribute with the given name of the element identified by xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: {            [key: string]: string        })](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateelement-member(1))|Updates the XML of the element identified by xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-add-member(1))|Adds a new custom XML part to the document.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getbynamespace-member(1))|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getcount-member(1))|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitemornullobject-member(1))|Gets a custom XML part based on its ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getcount-member(1))|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitemornullobject-member(1))|Gets a custom XML part based on its ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitem-member(1))|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|If the collection contains exactly one item, this method returns it.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#word-word-document-changetrackingmode-member)|Specifies the ChangeTracking mode.|
||[customXmlParts](/javascript/api/word/word.document#word-word-document-customxmlparts-member)|Gets the custom XML parts in the document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.document#word-word-document-deletebookmark-member(1))|Deletes a bookmark, if it exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#word-word-document-getbookmarkrange-member(1))|Gets a bookmark's range.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#word-word-document-getbookmarkrangeornullobject-member(1))|Gets a bookmark's range.|
||[settings](/javascript/api/word/word.document#word-word-document-settings-member)|Gets the add-in's settings in the document.|
|[Field](/javascript/api/word/word.field)|[code](/javascript/api/word/word.field#word-word-field-code-member)|Gets the field's code instruction.|
||[getNext()](/javascript/api/word/word.field#word-word-field-getnext-member(1))|Gets the next field.|
||[getNextOrNullObject()](/javascript/api/word/word.field#word-word-field-getnextornullobject-member(1))|Gets the next field.|
||[parentBody](/javascript/api/word/word.field#word-word-field-parentbody-member)|Gets the parent body of the field.|
||[parentContentControl](/javascript/api/word/word.field#word-word-field-parentcontentcontrol-member)|Gets the content control that contains the field.|
||[parentContentControlOrNullObject](/javascript/api/word/word.field#word-word-field-parentcontentcontrolornullobject-member)|Gets the content control that contains the field.|
||[parentTable](/javascript/api/word/word.field#word-word-field-parenttable-member)|Gets the table that contains the field.|
||[parentTableCell](/javascript/api/word/word.field#word-word-field-parenttablecell-member)|Gets the table cell that contains the field.|
||[parentTableCellOrNullObject](/javascript/api/word/word.field#word-word-field-parenttablecellornullobject-member)|Gets the table cell that contains the field.|
||[parentTableOrNullObject](/javascript/api/word/word.field#word-word-field-parenttableornullobject-member)|Gets the table that contains the field.|
||[result](/javascript/api/word/word.field#word-word-field-result-member)|Gets the field's result data.|
|[FieldCollection](/javascript/api/word/word.fieldcollection)|[getFirst()](/javascript/api/word/word.fieldcollection#word-word-fieldcollection-getfirst-member(1))|Gets the first field in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.fieldcollection#word-word-fieldcollection-getfirstornullobject-member(1))|Gets the first field in this collection.|
||[items](/javascript/api/word/word.fieldcollection#word-word-fieldcollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[fields](/javascript/api/word/word.paragraph#word-word-paragraph-fields-member)|Gets the collection of fields in the paragraph.|
||[getComments()](/javascript/api/word/word.paragraph#word-word-paragraph-getcomments-member(1))|Gets comments associated with the paragraph.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#word-word-paragraph-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Range](/javascript/api/word/word.range)|[fields](/javascript/api/word/word.range#word-word-range-fields-member)|Gets the collection of field objects in the range.|
||[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#word-word-range-getbookmarks-member(1))|Gets the names all bookmarks in or overlapping the range.|
||[getComments()](/javascript/api/word/word.range#word-word-range-getcomments-member(1))|Gets comments associated with the range.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#word-word-range-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#word-word-range-insertbookmark-member(1))|Inserts a bookmark on the range.|
||[insertComment(commentText: string)](/javascript/api/word/word.range#word-word-range-insertcomment-member(1))|Insert a comment on the range.|
|[Setting](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#word-word-setting-delete-member(1))|Deletes the setting.|
||[key](/javascript/api/word/word.setting#word-word-setting-key-member)|Gets the key of the setting.|
||[value](/javascript/api/word/word.setting#word-word-setting-value-member)|Specifies the value of the setting.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#word-word-settingcollection-add-member(1))|Creates a new setting or sets an existing setting.|
||[deleteAll()](/javascript/api/word/word.settingcollection#word-word-settingcollection-deleteall-member(1))|Deletes all settings in this add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#word-word-settingcollection-getcount-member(1))|Gets the count of settings.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitem-member(1))|Gets a setting object by its key, which is case-sensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitemornullobject-member(1))|Gets a setting object by its key, which is case-sensitive.|
||[items](/javascript/api/word/word.settingcollection#word-word-settingcollection-items-member)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/word/word.table)|[fields](/javascript/api/word/word.table#word-word-table-fields-member)|Gets the collection of field objects in the table.|
||[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#word-word-table-mergecells-member(1))|Merges the cells bounded inclusively by a first and last cell.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#word-word-tablecell-split-member(1))|Splits the cell into the specified number of rows and columns.|
|[TableRow](/javascript/api/word/word.tablerow)|[fields](/javascript/api/word/word.tablerow#word-word-tablerow-fields-member)|Gets the collection of field objects in the table row.|
||[merge()](/javascript/api/word/word.tablerow#word-word-tablerow-merge-member(1))|Merges the row into one cell.|
