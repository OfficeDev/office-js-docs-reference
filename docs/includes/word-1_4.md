| Class | Fields | Description |
|:---|:---|:---|
|[Body](/.body)|[fields](/.body#word-javascript/api/word/-body-fields-member)|Gets the collection of field objects in the body.|
||[getComments()](/.body#word-javascript/api/word/-body-getcomments-member(1))|Gets comments associated with the body.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/.body#word-javascript/api/word/-body-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Comment](/.comment)|[authorEmail](/.comment#word-javascript/api/word/-comment-authoremail-member)|Gets the email of the comment's author.|
||[authorName](/.comment#word-javascript/api/word/-comment-authorname-member)|Gets the name of the comment's author.|
||[content](/.comment#word-javascript/api/word/-comment-content-member)|Specifies the comment's content as plain text.|
||[contentRange](/.comment#word-javascript/api/word/-comment-contentrange-member)|Specifies the comment's content range.|
||[creationDate](/.comment#word-javascript/api/word/-comment-creationdate-member)|Gets the creation date of the comment.|
||[delete()](/.comment#word-javascript/api/word/-comment-delete-member(1))|Deletes the comment and its replies.|
||[getRange()](/.comment#word-javascript/api/word/-comment-getrange-member(1))|Gets the range in the main document where the comment is on.|
||[id](/.comment#word-javascript/api/word/-comment-id-member)|Gets the ID of the comment.|
||[replies](/.comment#word-javascript/api/word/-comment-replies-member)|Gets the collection of reply objects associated with the comment.|
||[reply(replyText: string)](/.comment#word-javascript/api/word/-comment-reply-member(1))|Adds a new reply to the end of the comment thread.|
||[resolved](/.comment#word-javascript/api/word/-comment-resolved-member)|Specifies the comment thread's status.|
|[CommentCollection](/.commentcollection)|[getFirst()](/.commentcollection#word-javascript/api/word/-commentcollection-getfirst-member(1))|Gets the first comment in the collection.|
||[getFirstOrNullObject()](/.commentcollection#word-javascript/api/word/-commentcollection-getfirstornullobject-member(1))|Gets the first comment in the collection.|
||[items](/.commentcollection#word-javascript/api/word/-commentcollection-items-member)|Gets the loaded child items in this collection.|
|[CommentContentRange](/.commentcontentrange)|[bold](/.commentcontentrange#word-javascript/api/word/-commentcontentrange-bold-member)|Specifies a value that indicates whether the comment text is bold.|
||[hyperlink](/.commentcontentrange#word-javascript/api/word/-commentcontentrange-hyperlink-member)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertText(text: string, insertLocation: Word.InsertLocation \| "Replace" \| "Start" \| "End" \| "Before" \| "After")](/.commentcontentrange#word-javascript/api/word/-commentcontentrange-inserttext-member(1))|Inserts text into at the specified location.|
||[isEmpty](/.commentcontentrange#word-javascript/api/word/-commentcontentrange-isempty-member)|Checks whether the range length is zero.|
||[italic](/.commentcontentrange#word-javascript/api/word/-commentcontentrange-italic-member)|Specifies a value that indicates whether the comment text is italicized.|
||[strikeThrough](/.commentcontentrange#word-javascript/api/word/-commentcontentrange-strikethrough-member)|Specifies a value that indicates whether the comment text has a strikethrough.|
||[text](/.commentcontentrange#word-javascript/api/word/-commentcontentrange-text-member)|Gets the text of the comment range.|
||[underline](/.commentcontentrange#word-javascript/api/word/-commentcontentrange-underline-member)|Specifies a value that indicates the comment text's underline type.|
|[CommentReply](/.commentreply)|[authorEmail](/.commentreply#word-javascript/api/word/-commentreply-authoremail-member)|Gets the email of the comment reply's author.|
||[authorName](/.commentreply#word-javascript/api/word/-commentreply-authorname-member)|Gets the name of the comment reply's author.|
||[content](/.commentreply#word-javascript/api/word/-commentreply-content-member)|Specifies the comment reply's content.|
||[contentRange](/.commentreply#word-javascript/api/word/-commentreply-contentrange-member)|Specifies the commentReply's content range.|
||[creationDate](/.commentreply#word-javascript/api/word/-commentreply-creationdate-member)|Gets the creation date of the comment reply.|
||[delete()](/.commentreply#word-javascript/api/word/-commentreply-delete-member(1))|Deletes the comment reply.|
||[id](/.commentreply#word-javascript/api/word/-commentreply-id-member)|Gets the ID of the comment reply.|
||[parentComment](/.commentreply#word-javascript/api/word/-commentreply-parentcomment-member)|Gets the parent comment of this reply.|
|[CommentReplyCollection](/.commentreplycollection)|[getFirst()](/.commentreplycollection#word-javascript/api/word/-commentreplycollection-getfirst-member(1))|Gets the first comment reply in the collection.|
||[getFirstOrNullObject()](/.commentreplycollection#word-javascript/api/word/-commentreplycollection-getfirstornullobject-member(1))|Gets the first comment reply in the collection.|
||[items](/.commentreplycollection#word-javascript/api/word/-commentreplycollection-items-member)|Gets the loaded child items in this collection.|
|[ContentControl](/.contentcontrol)|[fields](/.contentcontrol#word-javascript/api/word/-contentcontrol-fields-member)|Gets the collection of field objects in the content control.|
||[getComments()](/.contentcontrol#word-javascript/api/word/-contentcontrol-getcomments-member(1))|Gets comments associated with the content control.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/.contentcontrol#word-javascript/api/word/-contentcontrol-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[CustomXmlPart](/.customxmlpart)|[delete()](/.customxmlpart#word-javascript/api/word/-customxmlpart-delete-member(1))|Deletes the custom XML part.|
||[deleteAttribute(xpath: string, namespaceMappings: { [key: string]: string }, name: string)](/.customxmlpart#word-javascript/api/word/-customxmlpart-deleteattribute-member(1))|Deletes an attribute with the given name from the element identified by xpath.|
||[deleteElement(xpath: string, namespaceMappings: { [key: string]: string })](/.customxmlpart#word-javascript/api/word/-customxmlpart-deleteelement-member(1))|Deletes the element identified by xpath.|
||[getXml()](/.customxmlpart#word-javascript/api/word/-customxmlpart-getxml-member(1))|Gets the full XML content of the custom XML part.|
||[id](/.customxmlpart#word-javascript/api/word/-customxmlpart-id-member)|Gets the ID of the custom XML part.|
||[insertAttribute(xpath: string, namespaceMappings: { [key: string]: string }, name: string, value: string)](/.customxmlpart#word-javascript/api/word/-customxmlpart-insertattribute-member(1))|Inserts an attribute with the given name and value to the element identified by xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: { [key: string]: string }, index?: number)](/.customxmlpart#word-javascript/api/word/-customxmlpart-insertelement-member(1))|Inserts the given XML under the parent element identified by xpath at child position index.|
||[namespaceUri](/.customxmlpart#word-javascript/api/word/-customxmlpart-namespaceuri-member)|Gets the namespace URI of the custom XML part.|
||[query(xpath: string, namespaceMappings: { [key: string]: string })](/.customxmlpart#word-javascript/api/word/-customxmlpart-query-member(1))|Queries the XML content of the custom XML part.|
||[setXml(xml: string)](/.customxmlpart#word-javascript/api/word/-customxmlpart-setxml-member(1))|Sets the full XML content of the custom XML part.|
||[updateAttribute(xpath: string, namespaceMappings: { [key: string]: string }, name: string, value: string)](/.customxmlpart#word-javascript/api/word/-customxmlpart-updateattribute-member(1))|Updates the value of an attribute with the given name of the element identified by xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: { [key: string]: string })](/.customxmlpart#word-javascript/api/word/-customxmlpart-updateelement-member(1))|Updates the XML of the element identified by xpath.|
|[CustomXmlPartCollection](/.customxmlpartcollection)|[add(xml: string)](/.customxmlpartcollection#word-javascript/api/word/-customxmlpartcollection-add-member(1))|Adds a new custom XML part to the document.|
||[getByNamespace(namespaceUri: string)](/.customxmlpartcollection#word-javascript/api/word/-customxmlpartcollection-getbynamespace-member(1))|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/.customxmlpartcollection#word-javascript/api/word/-customxmlpartcollection-getcount-member(1))|Gets the number of items in the collection.|
||[getItem(id: string)](/.customxmlpartcollection#word-javascript/api/word/-customxmlpartcollection-getitem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/.customxmlpartcollection#word-javascript/api/word/-customxmlpartcollection-getitemornullobject-member(1))|Gets a custom XML part based on its ID.|
||[items](/.customxmlpartcollection#word-javascript/api/word/-customxmlpartcollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/.customxmlpartscopedcollection)|[getCount()](/.customxmlpartscopedcollection#word-javascript/api/word/-customxmlpartscopedcollection-getcount-member(1))|Gets the number of items in the collection.|
||[getItem(id: string)](/.customxmlpartscopedcollection#word-javascript/api/word/-customxmlpartscopedcollection-getitem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/.customxmlpartscopedcollection#word-javascript/api/word/-customxmlpartscopedcollection-getitemornullobject-member(1))|Gets a custom XML part based on its ID.|
||[getOnlyItem()](/.customxmlpartscopedcollection#word-javascript/api/word/-customxmlpartscopedcollection-getonlyitem-member(1))|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/.customxmlpartscopedcollection#word-javascript/api/word/-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|If the collection contains exactly one item, this method returns it.|
||[items](/.customxmlpartscopedcollection#word-javascript/api/word/-customxmlpartscopedcollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/.document)|[changeTrackingMode](/.document#word-javascript/api/word/-document-changetrackingmode-member)|Specifies the ChangeTracking mode.|
||[customXmlParts](/.document#word-javascript/api/word/-document-customxmlparts-member)|Gets the custom XML parts in the document.|
||[deleteBookmark(name: string)](/.document#word-javascript/api/word/-document-deletebookmark-member(1))|Deletes a bookmark, if it exists, from the document.|
||[getBookmarkRange(name: string)](/.document#word-javascript/api/word/-document-getbookmarkrange-member(1))|Gets a bookmark's range.|
||[getBookmarkRangeOrNullObject(name: string)](/.document#word-javascript/api/word/-document-getbookmarkrangeornullobject-member(1))|Gets a bookmark's range.|
||[settings](/.document#word-javascript/api/word/-document-settings-member)|Gets the add-in's settings in the document.|
|[Field](/.field)|[code](/.field#word-javascript/api/word/-field-code-member)|Specifies the field's code instruction.|
||[getNext()](/.field#word-javascript/api/word/-field-getnext-member(1))|Gets the next field.|
||[getNextOrNullObject()](/.field#word-javascript/api/word/-field-getnextornullobject-member(1))|Gets the next field.|
||[parentBody](/.field#word-javascript/api/word/-field-parentbody-member)|Gets the parent body of the field.|
||[parentContentControl](/.field#word-javascript/api/word/-field-parentcontentcontrol-member)|Gets the content control that contains the field.|
||[parentContentControlOrNullObject](/.field#word-javascript/api/word/-field-parentcontentcontrolornullobject-member)|Gets the content control that contains the field.|
||[parentTable](/.field#word-javascript/api/word/-field-parenttable-member)|Gets the table that contains the field.|
||[parentTableCell](/.field#word-javascript/api/word/-field-parenttablecell-member)|Gets the table cell that contains the field.|
||[parentTableCellOrNullObject](/.field#word-javascript/api/word/-field-parenttablecellornullobject-member)|Gets the table cell that contains the field.|
||[parentTableOrNullObject](/.field#word-javascript/api/word/-field-parenttableornullobject-member)|Gets the table that contains the field.|
||[result](/.field#word-javascript/api/word/-field-result-member)|Gets the field's result data.|
|[FieldCollection](/.fieldcollection)|[getFirst()](/.fieldcollection#word-javascript/api/word/-fieldcollection-getfirst-member(1))|Gets the first field in this collection.|
||[getFirstOrNullObject()](/.fieldcollection#word-javascript/api/word/-fieldcollection-getfirstornullobject-member(1))|Gets the first field in this collection.|
||[items](/.fieldcollection#word-javascript/api/word/-fieldcollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/.paragraph)|[fields](/.paragraph#word-javascript/api/word/-paragraph-fields-member)|Gets the collection of fields in the paragraph.|
||[getComments()](/.paragraph#word-javascript/api/word/-paragraph-getcomments-member(1))|Gets comments associated with the paragraph.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/.paragraph#word-javascript/api/word/-paragraph-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Range](/.range)|[fields](/.range#word-javascript/api/word/-range-fields-member)|Gets the collection of field objects in the range.|
||[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/.range#word-javascript/api/word/-range-getbookmarks-member(1))|Gets the names all bookmarks in or overlapping the range.|
||[getComments()](/.range#word-javascript/api/word/-range-getcomments-member(1))|Gets comments associated with the range.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/.range#word-javascript/api/word/-range-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[insertBookmark(name: string)](/.range#word-javascript/api/word/-range-insertbookmark-member(1))|Inserts a bookmark on the range.|
||[insertComment(commentText: string)](/.range#word-javascript/api/word/-range-insertcomment-member(1))|Insert a comment on the range.|
|[Setting](/.setting)|[delete()](/.setting#word-javascript/api/word/-setting-delete-member(1))|Deletes the setting.|
||[key](/.setting#word-javascript/api/word/-setting-key-member)|Gets the key of the setting.|
||[value](/.setting#word-javascript/api/word/-setting-value-member)|Specifies the value of the setting.|
|[SettingCollection](/.settingcollection)|[add(key: string, value: any)](/.settingcollection#word-javascript/api/word/-settingcollection-add-member(1))|Creates a new setting or sets an existing setting.|
||[deleteAll()](/.settingcollection#word-javascript/api/word/-settingcollection-deleteall-member(1))|Deletes all settings in this add-in.|
||[getCount()](/.settingcollection#word-javascript/api/word/-settingcollection-getcount-member(1))|Gets the count of settings.|
||[getItem(key: string)](/.settingcollection#word-javascript/api/word/-settingcollection-getitem-member(1))|Gets a setting object by its key, which is case-sensitive.|
||[getItemOrNullObject(key: string)](/.settingcollection#word-javascript/api/word/-settingcollection-getitemornullobject-member(1))|Gets a setting object by its key, which is case-sensitive.|
||[items](/.settingcollection#word-javascript/api/word/-settingcollection-items-member)|Gets the loaded child items in this collection.|
|[Table](/.table)|[fields](/.table#word-javascript/api/word/-table-fields-member)|Gets the collection of field objects in the table.|
||[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/.table#word-javascript/api/word/-table-mergecells-member(1))|Merges the cells bounded inclusively by a first and last cell.|
|[TableCell](/.tablecell)|[split(rowCount: number, columnCount: number)](/.tablecell#word-javascript/api/word/-tablecell-split-member(1))|Splits the cell into the specified number of rows and columns.|
|[TableRow](/.tablerow)|[fields](/.tablerow#word-javascript/api/word/-tablerow-fields-member)|Gets the collection of field objects in the table row.|
||[merge()](/.tablerow#word-javascript/api/word/-tablerow-merge-member(1))|Merges the row into one cell.|
