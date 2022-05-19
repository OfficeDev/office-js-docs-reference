| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#word-word-application-createdocument-member(1))|Creates a new document by using an optional base64 encoded .docx file.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-split-member(1))|Splits the content control into child ranges by using delimiters.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyidornullobject-member(1))|Gets a content control by its identifier.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytag-member(1))|Gets the content controls that have the specified tag.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytitle-member(1))|Gets the content controls that have the specified title.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytypes-member(1))|Gets the content controls that have the specified types and/or subtypes.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirst-member(1))|Gets the first content control in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirstornullobject-member(1))|Gets the first content control in this collection.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getitem-member(1))|Gets a content control by its index in the collection.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#word-word-customproperty-delete-member(1))|Deletes the custom property.|
||[key](/javascript/api/word/word.customproperty#word-word-customproperty-key-member)|Gets the key of the custom property.|
||[type](/javascript/api/word/word.customproperty#word-word-customproperty-type-member)|Gets the value type of the custom property.|
||[value](/javascript/api/word/word.customproperty#word-word-customproperty-value-member)|Gets or sets the value of the custom property.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-add-member(1))|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-deleteall-member(1))|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getcount-member(1))|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitem-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitemornullobject-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#word-word-document-properties-member)|Gets the properties of the document.|
||[sections](/javascript/api/word/word.document#word-word-document-sections-member)|Gets the collection of section objects in the document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[body](/javascript/api/word/word.documentcreated#word-word-documentcreated-body-member)|Gets the body object of the document.|
||[contentControls](/javascript/api/word/word.documentcreated#word-word-documentcreated-contentcontrols-member)|Gets the collection of content control objects in the document.|
||[open()](/javascript/api/word/word.documentcreated#word-word-documentcreated-open-member(1))|Opens the document.|
||[properties](/javascript/api/word/word.documentcreated#word-word-documentcreated-properties-member)|Gets the properties of the document.|
||[save()](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|Saves the document.|
||[saved](/javascript/api/word/word.documentcreated#word-word-documentcreated-saved-member)|Indicates whether the changes in the document have been saved.|
||[sections](/javascript/api/word/word.documentcreated#word-word-documentcreated-sections-member)|Gets the collection of section objects in the document.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[applicationName](/javascript/api/word/word.documentproperties#word-word-documentproperties-applicationname-member)|Gets the application name of the document.|
||[author](/javascript/api/word/word.documentproperties#word-word-documentproperties-author-member)|Gets or sets the author of the document.|
||[category](/javascript/api/word/word.documentproperties#word-word-documentproperties-category-member)|Gets or sets the category of the document.|
||[comments](/javascript/api/word/word.documentproperties#word-word-documentproperties-comments-member)|Gets or sets the comments of the document.|
||[company](/javascript/api/word/word.documentproperties#word-word-documentproperties-company-member)|Gets or sets the company of the document.|
||[creationDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-creationdate-member)|Gets the creation date of the document.|
||[customProperties](/javascript/api/word/word.documentproperties#word-word-documentproperties-customproperties-member)|Gets the collection of custom properties of the document.|
||[format](/javascript/api/word/word.documentproperties#word-word-documentproperties-format-member)|Gets or sets the format of the document.|
||[keywords](/javascript/api/word/word.documentproperties#word-word-documentproperties-keywords-member)|Gets or sets the keywords of the document.|
||[lastAuthor](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastauthor-member)|Gets the last author of the document.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastprintdate-member)|Gets the last print date of the document.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastsavetime-member)|Gets the last save time of the document.|
||[manager](/javascript/api/word/word.documentproperties#word-word-documentproperties-manager-member)|Gets or sets the manager of the document.|
||[revisionNumber](/javascript/api/word/word.documentproperties#word-word-documentproperties-revisionnumber-member)|Gets the revision number of the document.|
||[security](/javascript/api/word/word.documentproperties#word-word-documentproperties-security-member)|Gets security settings of the document.|
||[subject](/javascript/api/word/word.documentproperties#word-word-documentproperties-subject-member)|Gets or sets the subject of the document.|
||[template](/javascript/api/word/word.documentproperties#word-word-documentproperties-template-member)|Gets the template of the document.|
||[title](/javascript/api/word/word.documentproperties#word-word-documentproperties-title-member)|Gets or sets the title of the document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttextdescription-member)|Gets or sets a string that represents the alternative text associated with the inline image.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttexttitle-member)|Gets or sets a string that contains the title for the inline image.|
||[delete()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-delete-member(1))|Deletes the inline picture from the document.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getbase64imagesrc-member(1))|Gets the base64 encoded string representation of the inline image.|
||[getNext()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnext-member(1))|Gets the next inline image.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnextornullobject-member(1))|Gets the next inline image.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getrange-member(1))|Gets the picture, or the starting or ending point of the picture, as a range.|
||[height](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-height-member)|Gets or sets a number that describes the height of the inline image.|
||[hyperlink](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-hyperlink-member)|Gets or sets a hyperlink on the image.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertbreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertcontentcontrol-member(1))|Wraps the inline picture with a rich text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertfilefrombase64-member(1))|Inserts a document at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserthtml-member(1))|Inserts HTML at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertinlinepicturefrombase64-member(1))|Inserts an inline picture at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertooxml-member(1))|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserttext-member(1))|Inserts text at the specified location.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-lockaspectratio-member)|Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrolornullobject-member)|Gets the content control that contains the inline image.|
||[parentTable](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttable-member)|Gets the table that contains the inline image.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecell-member)|Gets the table cell that contains the inline image.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecellornullobject-member)|Gets the table cell that contains the inline image.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttableornullobject-member)|Gets the table that contains the inline image.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-select-member(1))|Selects the inline picture.|
||[width](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-width-member)|Gets or sets a number that describes the width of the inline image.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirst-member(1))|Gets the first inline image in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirstornullobject-member(1))|Gets the first inline image in this collection.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#word-word-list-getlevelparagraphs-member(1))|Gets the paragraphs that occur at the specified level in the list.|
||[getLevelString(level: number)](/javascript/api/word/word.list#word-word-list-getlevelstring-member(1))|Gets the bullet, number, or picture at the specified level as a string.|
||[id](/javascript/api/word/word.list#word-word-list-id-member)|Gets the list's id.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#word-word-list-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[levelExistences](/javascript/api/word/word.list#word-word-list-levelexistences-member)|Checks whether each of the 9 levels exists in the list.|
||[levelTypes](/javascript/api/word/word.list#word-word-list-leveltypes-member)|Gets all 9 level types in the list.|
||[paragraphs](/javascript/api/word/word.list#word-word-list-paragraphs-member)|Gets paragraphs in the list.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#word-word-list-setlevelalignment-member(1))|Sets the alignment of the bullet, number, or picture at the specified level in the list.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#word-word-list-setlevelbullet-member(1))|Sets the bullet format at the specified level in the list.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#word-word-list-setlevelindents-member(1))|Sets the two indents of the specified level in the list.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#word-word-list-setlevelnumbering-member(1))|Sets the numbering format at the specified level in the list.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#word-word-list-setlevelstartingnumber-member(1))|Sets the starting number at the specified level in the list.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyid-member(1))|Gets a list by its identifier.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyidornullobject-member(1))|Gets a list by its identifier.|
||[getFirst()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirst-member(1))|Gets the first list in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirstornullobject-member(1))|Gets the first list in this collection.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getitem-member(1))|Gets a list object by its index in the collection.|
||[items](/javascript/api/word/word.listcollection#word-word-listcollection-items-member)|Gets the loaded child items in this collection.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestor-member(1))|Gets the list item parent, or the closest ancestor if the parent does not exist.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestorornullobject-member(1))|Gets the list item parent, or the closest ancestor if the parent does not exist.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getdescendants-member(1))|Gets all descendant list items of the list item.|
||[level](/javascript/api/word/word.listitem#word-word-listitem-level-member)|Gets or sets the level of the item in the list.|
||[listString](/javascript/api/word/word.listitem#word-word-listitem-liststring-member)|Gets the list item bullet, number, or picture as a string.|
||[siblingIndex](/javascript/api/word/word.listitem#word-word-listitem-siblingindex-member)|Gets the list item order number in relation to its siblings.|
|[Paragraph](/javascript/api/word/word.paragraph)|[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-split-member(1))|Splits the paragraph into child ranges by using delimiters.|
||[startNewList()](/javascript/api/word/word.paragraph#word-word-paragraph-startnewlist-member(1))|Starts a new list with this paragraph.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirst-member(1))|Gets the first paragraph in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirstornullobject-member(1))|Gets the first paragraph in this collection.|
||[getLast()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlast-member(1))|Gets the last paragraph in this collection.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlastornullobject-member(1))|Gets the last paragraph in this collection.|
|[Range](/javascript/api/word/word.range)|[contentControls](/javascript/api/word/word.range#word-word-range-contentcontrols-member)|Gets the collection of content control objects in the range.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#word-word-range-select-member(1))|Selects and navigates the Word UI to the range.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-split-member(1))|Splits the range into child ranges by using delimiters.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirst-member(1))|Gets the first range in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirstornullobject-member(1))|Gets the first range in this collection.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#word-word-requestcontext-application-member)|[Api set: WordApi 1.3] *|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorepunct-member)|Gets or sets a value that indicates whether to ignore all punctuation characters between words.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorespace-member)|Gets or sets a value that indicates whether to ignore all whitespace between words.|
||[matchCase](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchcase-member)|Gets or sets a value that indicates whether to perform a case sensitive search.|
||[matchPrefix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchprefix-member)|Gets or sets a value that indicates whether to match words that begin with the search string.|
||[matchSuffix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchsuffix-member)|Gets or sets a value that indicates whether to match words that end with the search string.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwholeword-member)|Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word.|
||[matchWildcards](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwildcards-member)|Gets or sets a value that indicates whether the search will be performed using special search operators.|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#word-word-section-getnext-member(1))|Gets the next section.|
||[getNextOrNullObject()](/javascript/api/word/word.section#word-word-section-getnextornullobject-member(1))|Gets the next section.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirst-member(1))|Gets the first section in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirstornullobject-member(1))|Gets the first section in this collection.|
|[Table](/javascript/api/word/word.table)|[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#word-word-table-select-member(1))|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#word-word-table-setcellpadding-member(1))|Sets cell padding in points.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#word-word-tableborder-color-member)|Gets or sets the table border color.|
||[type](/javascript/api/word/word.tableborder#word-word-tableborder-type-member)|Gets or sets the type of the table border.|
||[width](/javascript/api/word/word.tableborder#word-word-tableborder-width-member)|Gets or sets the width, in points, of the table border.|
|[TableCell](/javascript/api/word/word.tablecell)|[body](/javascript/api/word/word.tablecell#word-word-tablecell-body-member)|Gets the body object of the cell.|
||[cellIndex](/javascript/api/word/word.tablecell#word-word-tablecell-cellindex-member)|Gets the index of the cell in its row.|
||[columnWidth](/javascript/api/word/word.tablecell#word-word-tablecell-columnwidth-member)|Gets and sets the width of the cell's column in points.|
||[deleteColumn()](/javascript/api/word/word.tablecell#word-word-tablecell-deletecolumn-member(1))|Deletes the column containing this cell.|
||[deleteRow()](/javascript/api/word/word.tablecell#word-word-tablecell-deleterow-member(1))|Deletes the row containing this cell.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getborder-member(1))|Gets the border style for the specified border.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getcellpadding-member(1))|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablecell#word-word-tablecell-getnext-member(1))|Gets the next cell.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#word-word-tablecell-getnextornullobject-member(1))|Gets the next cell.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-horizontalalignment-member)|Gets and sets the horizontal alignment of the cell.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertcolumns-member(1))|Adds columns to the left or right of the cell, using the cell's column as a template.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertrows-member(1))|Inserts rows above or below the cell, using the cell's row as a template.|
||[parentRow](/javascript/api/word/word.tablecell#word-word-tablecell-parentrow-member)|Gets the parent row of the cell.|
||[parentTable](/javascript/api/word/word.tablecell#word-word-tablecell-parenttable-member)|Gets the parent table of the cell.|
||[rowIndex](/javascript/api/word/word.tablecell#word-word-tablecell-rowindex-member)|Gets the index of the cell's row in the table.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#word-word-tablecell-setcellpadding-member(1))|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablecell#word-word-tablecell-shadingcolor-member)|Gets or sets the shading color of the cell.|
||[value](/javascript/api/word/word.tablecell#word-word-tablecell-value-member)|Gets and sets the text of the cell.|
||[verticalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-verticalalignment-member)|Gets and sets the vertical alignment of the cell.|
||[width](/javascript/api/word/word.tablecell#word-word-tablecell-width-member)|Gets the width of the cell in points.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirst-member(1))|Gets the first table cell in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirstornullobject-member(1))|Gets the first table cell in this collection.|
||[items](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-items-member)|Gets the loaded child items in this collection.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirst-member(1))|Gets the first table in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirstornullobject-member(1))|Gets the first table in this collection.|
||[items](/javascript/api/word/word.tablecollection#word-word-tablecollection-items-member)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/word/word.tablerow)|[cells](/javascript/api/word/word.tablerow#word-word-tablerow-cells-member)|Gets cells.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#word-word-tablerow-select-member(1))|Selects the row and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#word-word-tablerow-setcellpadding-member(1))|Sets cell padding in points.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirst-member(1))|Gets the first row in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirstornullobject-member(1))|Gets the first row in this collection.|
||[items](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-items-member)|Gets the loaded child items in this collection.|
