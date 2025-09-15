| Class | Fields | Description |
|:---|:---|:---|
|[Application](/.application)|[createDocument(base64File?: string)](/.application#word-javascript/api/word/-application-createdocument-member(1))|Creates a new document by using an optional Base64-encoded .docx file.|
|[Body](/.body)|[getRange(rangeLocation?: Word.RangeLocation.whole \| Word.RangeLocation.start \| Word.RangeLocation.end \| Word.RangeLocation.after \| Word.RangeLocation.content \| "Whole" \| "Start" \| "End" \| "After" \| "Content")](/.body#word-javascript/api/word/-body-getrange-member(1))|Gets the whole body, or the starting or ending point of the body, as a range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.start \| Word.InsertLocation.end \| "Start" \| "End", values?: string[][])](/.body#word-javascript/api/word/-body-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[lists](/.body#word-javascript/api/word/-body-lists-member)|Gets the collection of list objects in the body.|
||[parentBody](/.body#word-javascript/api/word/-body-parentbody-member)|Gets the parent body of the body.|
||[parentBodyOrNullObject](/.body#word-javascript/api/word/-body-parentbodyornullobject-member)|Gets the parent body of the body.|
||[parentContentControlOrNullObject](/.body#word-javascript/api/word/-body-parentcontentcontrolornullobject-member)|Gets the content control that contains the body.|
||[parentSection](/.body#word-javascript/api/word/-body-parentsection-member)|Gets the parent section of the body.|
||[parentSectionOrNullObject](/.body#word-javascript/api/word/-body-parentsectionornullobject-member)|Gets the parent section of the body.|
||[styleBuiltIn](/.body#word-javascript/api/word/-body-stylebuiltin-member)|Specifies the built-in style name for the body.|
||[tables](/.body#word-javascript/api/word/-body-tables-member)|Gets the collection of table objects in the body.|
||[type](/.body#word-javascript/api/word/-body-type-member)|Gets the type of the body.|
|[ContentControl](/.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation \| "Whole" \| "Start" \| "End" \| "Before" \| "After" \| "Content")](/.contentcontrol#word-javascript/api/word/-contentcontrol-getrange-member(1))|Gets the whole content control, or the starting or ending point of the content control, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/.contentcontrol#word-javascript/api/word/-contentcontrol-gettextranges-member(1))|Gets the text ranges in the content control by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.start \| Word.InsertLocation.end \| Word.InsertLocation.before \| Word.InsertLocation.after \| "Start" \| "End" \| "Before" \| "After", values?: string[][])](/.contentcontrol#word-javascript/api/word/-contentcontrol-inserttable-member(1))|Inserts a table with the specified number of rows and columns into, or next to, a content control.|
||[lists](/.contentcontrol#word-javascript/api/word/-contentcontrol-lists-member)|Gets the collection of list objects in the content control.|
||[parentBody](/.contentcontrol#word-javascript/api/word/-contentcontrol-parentbody-member)|Gets the parent body of the content control.|
||[parentContentControlOrNullObject](/.contentcontrol#word-javascript/api/word/-contentcontrol-parentcontentcontrolornullobject-member)|Gets the content control that contains the content control.|
||[parentTable](/.contentcontrol#word-javascript/api/word/-contentcontrol-parenttable-member)|Gets the table that contains the content control.|
||[parentTableCell](/.contentcontrol#word-javascript/api/word/-contentcontrol-parenttablecell-member)|Gets the table cell that contains the content control.|
||[parentTableCellOrNullObject](/.contentcontrol#word-javascript/api/word/-contentcontrol-parenttablecellornullobject-member)|Gets the table cell that contains the content control.|
||[parentTableOrNullObject](/.contentcontrol#word-javascript/api/word/-contentcontrol-parenttableornullobject-member)|Gets the table that contains the content control.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/.contentcontrol#word-javascript/api/word/-contentcontrol-split-member(1))|Splits the content control into child ranges by using delimiters.|
||[styleBuiltIn](/.contentcontrol#word-javascript/api/word/-contentcontrol-stylebuiltin-member)|Specifies the built-in style name for the content control.|
||[subtype](/.contentcontrol#word-javascript/api/word/-contentcontrol-subtype-member)|Gets the content control subtype.|
||[tables](/.contentcontrol#word-javascript/api/word/-contentcontrol-tables-member)|Gets the collection of table objects in the content control.|
|[ContentControlCollection](/.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/.contentcontrolcollection#word-javascript/api/word/-contentcontrolcollection-getbyidornullobject-member(1))|Gets a content control by its identifier.|
||[getByTypes(types: Word.ContentControlType[])](/.contentcontrolcollection#word-javascript/api/word/-contentcontrolcollection-getbytypes-member(1))|Gets the content controls that have the specified types.|
||[getFirst()](/.contentcontrolcollection#word-javascript/api/word/-contentcontrolcollection-getfirst-member(1))|Gets the first content control in this collection.|
||[getFirstOrNullObject()](/.contentcontrolcollection#word-javascript/api/word/-contentcontrolcollection-getfirstornullobject-member(1))|Gets the first content control in this collection.|
|[CustomProperty](/.customproperty)|[delete()](/.customproperty#word-javascript/api/word/-customproperty-delete-member(1))|Deletes the custom property.|
||[key](/.customproperty#word-javascript/api/word/-customproperty-key-member)|Gets the key of the custom property.|
||[type](/.customproperty#word-javascript/api/word/-customproperty-type-member)|Gets the value type of the custom property.|
||[value](/.customproperty#word-javascript/api/word/-customproperty-value-member)|Specifies the value of the custom property.|
|[CustomPropertyCollection](/.custompropertycollection)|[add(key: string, value: any)](/.custompropertycollection#word-javascript/api/word/-custompropertycollection-add-member(1))|Creates a new or sets an existing custom property.|
||[deleteAll()](/.custompropertycollection#word-javascript/api/word/-custompropertycollection-deleteall-member(1))|Deletes all custom properties in this collection.|
||[getCount()](/.custompropertycollection#word-javascript/api/word/-custompropertycollection-getcount-member(1))|Gets the count of custom properties.|
||[getItem(key: string)](/.custompropertycollection#word-javascript/api/word/-custompropertycollection-getitem-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/.custompropertycollection#word-javascript/api/word/-custompropertycollection-getitemornullobject-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[items](/.custompropertycollection#word-javascript/api/word/-custompropertycollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/.document)|[properties](/.document#word-javascript/api/word/-document-properties-member)|Gets the properties of the document.|
|[DocumentCreated](/.documentcreated)|[open()](/.documentcreated#word-javascript/api/word/-documentcreated-open-member(1))|Opens the document.|
|[DocumentProperties](/.documentproperties)|[applicationName](/.documentproperties#word-javascript/api/word/-documentproperties-applicationname-member)|Gets the application name of the document.|
||[author](/.documentproperties#word-javascript/api/word/-documentproperties-author-member)|Specifies the author of the document.|
||[category](/.documentproperties#word-javascript/api/word/-documentproperties-category-member)|Specifies the category of the document.|
||[comments](/.documentproperties#word-javascript/api/word/-documentproperties-comments-member)|Specifies the Comments field in the metadata of the document.|
||[company](/.documentproperties#word-javascript/api/word/-documentproperties-company-member)|Specifies the company of the document.|
||[creationDate](/.documentproperties#word-javascript/api/word/-documentproperties-creationdate-member)|Gets the creation date of the document.|
||[customProperties](/.documentproperties#word-javascript/api/word/-documentproperties-customproperties-member)|Gets the collection of custom properties of the document.|
||[format](/.documentproperties#word-javascript/api/word/-documentproperties-format-member)|Specifies the format of the document.|
||[keywords](/.documentproperties#word-javascript/api/word/-documentproperties-keywords-member)|Specifies the keywords of the document.|
||[lastAuthor](/.documentproperties#word-javascript/api/word/-documentproperties-lastauthor-member)|Gets the last author of the document.|
||[lastPrintDate](/.documentproperties#word-javascript/api/word/-documentproperties-lastprintdate-member)|Gets the last print date of the document.|
||[lastSaveTime](/.documentproperties#word-javascript/api/word/-documentproperties-lastsavetime-member)|Gets the last save time of the document.|
||[manager](/.documentproperties#word-javascript/api/word/-documentproperties-manager-member)|Specifies the manager of the document.|
||[revisionNumber](/.documentproperties#word-javascript/api/word/-documentproperties-revisionnumber-member)|Gets the revision number of the document.|
||[security](/.documentproperties#word-javascript/api/word/-documentproperties-security-member)|Gets security settings of the document.|
||[subject](/.documentproperties#word-javascript/api/word/-documentproperties-subject-member)|Specifies the subject of the document.|
||[template](/.documentproperties#word-javascript/api/word/-documentproperties-template-member)|Gets the template of the document.|
||[title](/.documentproperties#word-javascript/api/word/-documentproperties-title-member)|Specifies the title of the document.|
|[InlinePicture](/.inlinepicture)|[getNext()](/.inlinepicture#word-javascript/api/word/-inlinepicture-getnext-member(1))|Gets the next inline image.|
||[getNextOrNullObject()](/.inlinepicture#word-javascript/api/word/-inlinepicture-getnextornullobject-member(1))|Gets the next inline image.|
||[getRange(rangeLocation?: Word.RangeLocation.whole \| Word.RangeLocation.start \| Word.RangeLocation.end \| "Whole" \| "Start" \| "End")](/.inlinepicture#word-javascript/api/word/-inlinepicture-getrange-member(1))|Gets the picture, or the starting or ending point of the picture, as a range.|
||[parentContentControlOrNullObject](/.inlinepicture#word-javascript/api/word/-inlinepicture-parentcontentcontrolornullobject-member)|Gets the content control that contains the inline image.|
||[parentTable](/.inlinepicture#word-javascript/api/word/-inlinepicture-parenttable-member)|Gets the table that contains the inline image.|
||[parentTableCell](/.inlinepicture#word-javascript/api/word/-inlinepicture-parenttablecell-member)|Gets the table cell that contains the inline image.|
||[parentTableCellOrNullObject](/.inlinepicture#word-javascript/api/word/-inlinepicture-parenttablecellornullobject-member)|Gets the table cell that contains the inline image.|
||[parentTableOrNullObject](/.inlinepicture#word-javascript/api/word/-inlinepicture-parenttableornullobject-member)|Gets the table that contains the inline image.|
|[InlinePictureCollection](/.inlinepicturecollection)|[getFirst()](/.inlinepicturecollection#word-javascript/api/word/-inlinepicturecollection-getfirst-member(1))|Gets the first inline image in this collection.|
||[getFirstOrNullObject()](/.inlinepicturecollection#word-javascript/api/word/-inlinepicturecollection-getfirstornullobject-member(1))|Gets the first inline image in this collection.|
|[List](/.list)|[getLevelParagraphs(level: number)](/.list#word-javascript/api/word/-list-getlevelparagraphs-member(1))|Gets the paragraphs that occur at the specified level in the list.|
||[getLevelString(level: number)](/.list#word-javascript/api/word/-list-getlevelstring-member(1))|Gets the bullet, number, or picture at the specified level as a string.|
||[id](/.list#word-javascript/api/word/-list-id-member)|Gets the list's id.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.start \| Word.InsertLocation.end \| Word.InsertLocation.before \| Word.InsertLocation.after \| "Start" \| "End" \| "Before" \| "After")](/.list#word-javascript/api/word/-list-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[levelExistences](/.list#word-javascript/api/word/-list-levelexistences-member)|Checks whether each of the 9 levels exists in the list.|
||[levelTypes](/.list#word-javascript/api/word/-list-leveltypes-member)|Gets all 9 level types in the list.|
||[paragraphs](/.list#word-javascript/api/word/-list-paragraphs-member)|Gets paragraphs in the list.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/.list#word-javascript/api/word/-list-setlevelalignment-member(1))|Sets the alignment of the bullet, number, or picture at the specified level in the list.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/.list#word-javascript/api/word/-list-setlevelbullet-member(1))|Sets the bullet format at the specified level in the list.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/.list#word-javascript/api/word/-list-setlevelindents-member(1))|Sets the two indents of the specified level in the list.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/.list#word-javascript/api/word/-list-setlevelnumbering-member(1))|Sets the numbering format at the specified level in the list.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/.list#word-javascript/api/word/-list-setlevelstartingnumber-member(1))|Sets the starting number at the specified level in the list.|
|[ListCollection](/.listcollection)|[getById(id: number)](/.listcollection#word-javascript/api/word/-listcollection-getbyid-member(1))|Gets a list by its identifier.|
||[getByIdOrNullObject(id: number)](/.listcollection#word-javascript/api/word/-listcollection-getbyidornullobject-member(1))|Gets a list by its identifier.|
||[getFirst()](/.listcollection#word-javascript/api/word/-listcollection-getfirst-member(1))|Gets the first list in this collection.|
||[getFirstOrNullObject()](/.listcollection#word-javascript/api/word/-listcollection-getfirstornullobject-member(1))|Gets the first list in this collection.|
||[getItem(id: number)](/.listcollection#word-javascript/api/word/-listcollection-getitem-member(1))|Gets a list object by its ID.|
||[items](/.listcollection#word-javascript/api/word/-listcollection-items-member)|Gets the loaded child items in this collection.|
|[ListItem](/.listitem)|[getAncestor(parentOnly?: boolean)](/.listitem#word-javascript/api/word/-listitem-getancestor-member(1))|Gets the list item parent, or the closest ancestor if the parent doesn't exist.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/.listitem#word-javascript/api/word/-listitem-getancestorornullobject-member(1))|Gets the list item parent, or the closest ancestor if the parent doesn't exist.|
||[getDescendants(directChildrenOnly?: boolean)](/.listitem#word-javascript/api/word/-listitem-getdescendants-member(1))|Gets all descendant list items of the list item.|
||[level](/.listitem#word-javascript/api/word/-listitem-level-member)|Specifies the level of the item in the list.|
||[listString](/.listitem#word-javascript/api/word/-listitem-liststring-member)|Gets the list item bullet, number, or picture as a string.|
||[siblingIndex](/.listitem#word-javascript/api/word/-listitem-siblingindex-member)|Gets the list item order number in relation to its siblings.|
|[Paragraph](/.paragraph)|[attachToList(listId: number, level: number)](/.paragraph#word-javascript/api/word/-paragraph-attachtolist-member(1))|Lets the paragraph join an existing list at the specified level.|
||[detachFromList()](/.paragraph#word-javascript/api/word/-paragraph-detachfromlist-member(1))|Moves this paragraph out of its list, if the paragraph is a list item.|
||[getNext()](/.paragraph#word-javascript/api/word/-paragraph-getnext-member(1))|Gets the next paragraph.|
||[getNextOrNullObject()](/.paragraph#word-javascript/api/word/-paragraph-getnextornullobject-member(1))|Gets the next paragraph.|
||[getPrevious()](/.paragraph#word-javascript/api/word/-paragraph-getprevious-member(1))|Gets the previous paragraph.|
||[getPreviousOrNullObject()](/.paragraph#word-javascript/api/word/-paragraph-getpreviousornullobject-member(1))|Gets the previous paragraph.|
||[getRange(rangeLocation?: Word.RangeLocation.whole \| Word.RangeLocation.start \| Word.RangeLocation.end \| Word.RangeLocation.after \| Word.RangeLocation.content \| "Whole" \| "Start" \| "End" \| "After" \| "Content")](/.paragraph#word-javascript/api/word/-paragraph-getrange-member(1))|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/.paragraph#word-javascript/api/word/-paragraph-gettextranges-member(1))|Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After", values?: string[][])](/.paragraph#word-javascript/api/word/-paragraph-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[isLastParagraph](/.paragraph#word-javascript/api/word/-paragraph-islastparagraph-member)|Indicates the paragraph is the last one inside its parent body.|
||[isListItem](/.paragraph#word-javascript/api/word/-paragraph-islistitem-member)|Checks whether the paragraph is a list item.|
||[list](/.paragraph#word-javascript/api/word/-paragraph-list-member)|Gets the List to which this paragraph belongs.|
||[listItem](/.paragraph#word-javascript/api/word/-paragraph-listitem-member)|Gets the ListItem for the paragraph.|
||[listItemOrNullObject](/.paragraph#word-javascript/api/word/-paragraph-listitemornullobject-member)|Gets the ListItem for the paragraph.|
||[listOrNullObject](/.paragraph#word-javascript/api/word/-paragraph-listornullobject-member)|Gets the List to which this paragraph belongs.|
||[parentBody](/.paragraph#word-javascript/api/word/-paragraph-parentbody-member)|Gets the parent body of the paragraph.|
||[parentContentControlOrNullObject](/.paragraph#word-javascript/api/word/-paragraph-parentcontentcontrolornullobject-member)|Gets the content control that contains the paragraph.|
||[parentTable](/.paragraph#word-javascript/api/word/-paragraph-parenttable-member)|Gets the table that contains the paragraph.|
||[parentTableCell](/.paragraph#word-javascript/api/word/-paragraph-parenttablecell-member)|Gets the table cell that contains the paragraph.|
||[parentTableCellOrNullObject](/.paragraph#word-javascript/api/word/-paragraph-parenttablecellornullobject-member)|Gets the table cell that contains the paragraph.|
||[parentTableOrNullObject](/.paragraph#word-javascript/api/word/-paragraph-parenttableornullobject-member)|Gets the table that contains the paragraph.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/.paragraph#word-javascript/api/word/-paragraph-split-member(1))|Splits the paragraph into child ranges by using delimiters.|
||[startNewList()](/.paragraph#word-javascript/api/word/-paragraph-startnewlist-member(1))|Starts a new list with this paragraph.|
||[styleBuiltIn](/.paragraph#word-javascript/api/word/-paragraph-stylebuiltin-member)|Specifies the built-in style name for the paragraph.|
||[tableNestingLevel](/.paragraph#word-javascript/api/word/-paragraph-tablenestinglevel-member)|Gets the level of the paragraph's table.|
|[ParagraphCollection](/.paragraphcollection)|[getFirst()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-getfirst-member(1))|Gets the first paragraph in this collection.|
||[getFirstOrNullObject()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-getfirstornullobject-member(1))|Gets the first paragraph in this collection.|
||[getLast()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-getlast-member(1))|Gets the last paragraph in this collection.|
||[getLastOrNullObject()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-getlastornullobject-member(1))|Gets the last paragraph in this collection.|
|[Range](/.range)|[compareLocationWith(range: Word.Range)](/.range#word-javascript/api/word/-range-comparelocationwith-member(1))|Compares this range's location with another range's location.|
||[expandTo(range: Word.Range)](/.range#word-javascript/api/word/-range-expandto-member(1))|Returns a new range that extends from this range in either direction to cover another range.|
||[expandToOrNullObject(range: Word.Range)](/.range#word-javascript/api/word/-range-expandtoornullobject-member(1))|Returns a new range that extends from this range in either direction to cover another range.|
||[getHyperlinkRanges()](/.range#word-javascript/api/word/-range-gethyperlinkranges-member(1))|Gets hyperlink child ranges within the range.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/.range#word-javascript/api/word/-range-getnexttextrange-member(1))|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/.range#word-javascript/api/word/-range-getnexttextrangeornullobject-member(1))|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getRange(rangeLocation?: Word.RangeLocation.whole \| Word.RangeLocation.start \| Word.RangeLocation.end \| Word.RangeLocation.after \| Word.RangeLocation.content \| "Whole" \| "Start" \| "End" \| "After" \| "Content")](/.range#word-javascript/api/word/-range-getrange-member(1))|Clones the range, or gets the starting or ending point of the range as a new range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/.range#word-javascript/api/word/-range-gettextranges-member(1))|Gets the text child ranges in the range by using punctuation marks and/or other ending marks.|
||[hyperlink](/.range#word-javascript/api/word/-range-hyperlink-member)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After", values?: string[][])](/.range#word-javascript/api/word/-range-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[intersectWith(range: Word.Range)](/.range#word-javascript/api/word/-range-intersectwith-member(1))|Returns a new range as the intersection of this range with another range.|
||[intersectWithOrNullObject(range: Word.Range)](/.range#word-javascript/api/word/-range-intersectwithornullobject-member(1))|Returns a new range as the intersection of this range with another range.|
||[isEmpty](/.range#word-javascript/api/word/-range-isempty-member)|Checks whether the range length is zero.|
||[lists](/.range#word-javascript/api/word/-range-lists-member)|Gets the collection of list objects in the range.|
||[parentBody](/.range#word-javascript/api/word/-range-parentbody-member)|Gets the parent body of the range.|
||[parentContentControlOrNullObject](/.range#word-javascript/api/word/-range-parentcontentcontrolornullobject-member)|Gets the currently supported content control that contains the range.|
||[parentTable](/.range#word-javascript/api/word/-range-parenttable-member)|Gets the table that contains the range.|
||[parentTableCell](/.range#word-javascript/api/word/-range-parenttablecell-member)|Gets the table cell that contains the range.|
||[parentTableCellOrNullObject](/.range#word-javascript/api/word/-range-parenttablecellornullobject-member)|Gets the table cell that contains the range.|
||[parentTableOrNullObject](/.range#word-javascript/api/word/-range-parenttableornullobject-member)|Gets the table that contains the range.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/.range#word-javascript/api/word/-range-split-member(1))|Splits the range into child ranges by using delimiters.|
||[styleBuiltIn](/.range#word-javascript/api/word/-range-stylebuiltin-member)|Specifies the built-in style name for the range.|
||[tables](/.range#word-javascript/api/word/-range-tables-member)|Gets the collection of table objects in the range.|
|[RangeCollection](/.rangecollection)|[getFirst()](/.rangecollection#word-javascript/api/word/-rangecollection-getfirst-member(1))|Gets the first range in this collection.|
||[getFirstOrNullObject()](/.rangecollection#word-javascript/api/word/-rangecollection-getfirstornullobject-member(1))|Gets the first range in this collection.|
|[RequestContext](/.requestcontext)|[application](/.requestcontext#word-javascript/api/word/-requestcontext-application-member)|[Api set: WordApi 1.3] *|
|[Section](/.section)|[getNext()](/.section#word-javascript/api/word/-section-getnext-member(1))|Gets the next section.|
||[getNextOrNullObject()](/.section#word-javascript/api/word/-section-getnextornullobject-member(1))|Gets the next section.|
|[SectionCollection](/.sectioncollection)|[getFirst()](/.sectioncollection#word-javascript/api/word/-sectioncollection-getfirst-member(1))|Gets the first section in this collection.|
||[getFirstOrNullObject()](/.sectioncollection#word-javascript/api/word/-sectioncollection-getfirstornullobject-member(1))|Gets the first section in this collection.|
|[Style](/.style)|||
|[Table](/.table)|[addColumns(insertLocation: Word.InsertLocation.start \| Word.InsertLocation.end \| "Start" \| "End", columnCount: number, values?: string[][])](/.table#word-javascript/api/word/-table-addcolumns-member(1))|Adds columns to the start or end of the table, using the first or last existing column as a template.|
||[addRows(insertLocation: Word.InsertLocation.start \| Word.InsertLocation.end \| "Start" \| "End", rowCount: number, values?: string[][])](/.table#word-javascript/api/word/-table-addrows-member(1))|Adds rows to the start or end of the table, using the first or last existing row as a template.|
||[alignment](/.table#word-javascript/api/word/-table-alignment-member)|Specifies the alignment of the table against the page column.|
||[autoFitWindow()](/.table#word-javascript/api/word/-table-autofitwindow-member(1))|Autofits the table columns to the width of the window.|
||[clear()](/.table#word-javascript/api/word/-table-clear-member(1))|Clears the contents of the table.|
||[delete()](/.table#word-javascript/api/word/-table-delete-member(1))|Deletes the entire table.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/.table#word-javascript/api/word/-table-deletecolumns-member(1))|Deletes specific columns.|
||[deleteRows(rowIndex: number, rowCount?: number)](/.table#word-javascript/api/word/-table-deleterows-member(1))|Deletes specific rows.|
||[distributeColumns()](/.table#word-javascript/api/word/-table-distributecolumns-member(1))|Distributes the column widths evenly.|
||[font](/.table#word-javascript/api/word/-table-font-member)|Gets the font.|
||[getBorder(borderLocation: Word.BorderLocation)](/.table#word-javascript/api/word/-table-getborder-member(1))|Gets the border style for the specified border.|
||[getCell(rowIndex: number, cellIndex: number)](/.table#word-javascript/api/word/-table-getcell-member(1))|Gets the table cell at a specified row and column.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/.table#word-javascript/api/word/-table-getcellornullobject-member(1))|Gets the table cell at a specified row and column.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/.table#word-javascript/api/word/-table-getcellpadding-member(1))|Gets cell padding in points.|
||[getNext()](/.table#word-javascript/api/word/-table-getnext-member(1))|Gets the next table.|
||[getNextOrNullObject()](/.table#word-javascript/api/word/-table-getnextornullobject-member(1))|Gets the next table.|
||[getParagraphAfter()](/.table#word-javascript/api/word/-table-getparagraphafter-member(1))|Gets the paragraph after the table.|
||[getParagraphAfterOrNullObject()](/.table#word-javascript/api/word/-table-getparagraphafterornullobject-member(1))|Gets the paragraph after the table.|
||[getParagraphBefore()](/.table#word-javascript/api/word/-table-getparagraphbefore-member(1))|Gets the paragraph before the table.|
||[getParagraphBeforeOrNullObject()](/.table#word-javascript/api/word/-table-getparagraphbeforeornullobject-member(1))|Gets the paragraph before the table.|
||[getRange(rangeLocation?: Word.RangeLocation.whole \| Word.RangeLocation.start \| Word.RangeLocation.end \| Word.RangeLocation.after \| "Whole" \| "Start" \| "End" \| "After")](/.table#word-javascript/api/word/-table-getrange-member(1))|Gets the range that contains this table, or the range at the start or end of the table.|
||[headerRowCount](/.table#word-javascript/api/word/-table-headerrowcount-member)|Specifies the number of header rows.|
||[horizontalAlignment](/.table#word-javascript/api/word/-table-horizontalalignment-member)|Specifies the horizontal alignment of every cell in the table.|
||[insertContentControl()](/.table#word-javascript/api/word/-table-insertcontentcontrol-member(1))|Inserts a content control on the table.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After")](/.table#word-javascript/api/word/-table-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After", values?: string[][])](/.table#word-javascript/api/word/-table-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[isUniform](/.table#word-javascript/api/word/-table-isuniform-member)|Indicates whether all of the table rows are uniform.|
||[nestingLevel](/.table#word-javascript/api/word/-table-nestinglevel-member)|Gets the nesting level of the table.|
||[parentBody](/.table#word-javascript/api/word/-table-parentbody-member)|Gets the parent body of the table.|
||[parentContentControl](/.table#word-javascript/api/word/-table-parentcontentcontrol-member)|Gets the content control that contains the table.|
||[parentContentControlOrNullObject](/.table#word-javascript/api/word/-table-parentcontentcontrolornullobject-member)|Gets the content control that contains the table.|
||[parentTable](/.table#word-javascript/api/word/-table-parenttable-member)|Gets the table that contains this table.|
||[parentTableCell](/.table#word-javascript/api/word/-table-parenttablecell-member)|Gets the table cell that contains this table.|
||[parentTableCellOrNullObject](/.table#word-javascript/api/word/-table-parenttablecellornullobject-member)|Gets the table cell that contains this table.|
||[parentTableOrNullObject](/.table#word-javascript/api/word/-table-parenttableornullobject-member)|Gets the table that contains this table.|
||[rowCount](/.table#word-javascript/api/word/-table-rowcount-member)|Gets the number of rows in the table.|
||[rows](/.table#word-javascript/api/word/-table-rows-member)|Gets all of the table rows.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/.table#word-javascript/api/word/-table-search-member(1))|Performs a search with the specified SearchOptions on the scope of the table object.|
||[select(selectionMode?: Word.SelectionMode)](/.table#word-javascript/api/word/-table-select-member(1))|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/.table#word-javascript/api/word/-table-setcellpadding-member(1))|Sets cell padding in points.|
||[shadingColor](/.table#word-javascript/api/word/-table-shadingcolor-member)|Specifies the shading color.|
||[style](/.table#word-javascript/api/word/-table-style-member)|Specifies the style name for the table.|
||[styleBandedColumns](/.table#word-javascript/api/word/-table-stylebandedcolumns-member)|Specifies whether the table has banded columns.|
||[styleBandedRows](/.table#word-javascript/api/word/-table-stylebandedrows-member)|Specifies whether the table has banded rows.|
||[styleBuiltIn](/.table#word-javascript/api/word/-table-stylebuiltin-member)|Specifies the built-in style name for the table.|
||[styleFirstColumn](/.table#word-javascript/api/word/-table-stylefirstcolumn-member)|Specifies whether the table has a first column with a special style.|
||[styleLastColumn](/.table#word-javascript/api/word/-table-stylelastcolumn-member)|Specifies whether the table has a last column with a special style.|
||[styleTotalRow](/.table#word-javascript/api/word/-table-styletotalrow-member)|Specifies whether the table has a total (last) row with a special style.|
||[tables](/.table#word-javascript/api/word/-table-tables-member)|Gets the child tables nested one level deeper.|
||[values](/.table#word-javascript/api/word/-table-values-member)|Specifies the text values in the table, as a 2D JavaScript array.|
||[verticalAlignment](/.table#word-javascript/api/word/-table-verticalalignment-member)|Specifies the vertical alignment of every cell in the table.|
||[width](/.table#word-javascript/api/word/-table-width-member)|Specifies the width of the table in points.|
|[TableBorder](/.tableborder)|[color](/.tableborder#word-javascript/api/word/-tableborder-color-member)|Specifies the table border color.|
||[type](/.tableborder#word-javascript/api/word/-tableborder-type-member)|Specifies the type of the table border.|
||[width](/.tableborder#word-javascript/api/word/-tableborder-width-member)|Specifies the width, in points, of the table border.|
|[TableCell](/.tablecell)|[body](/.tablecell#word-javascript/api/word/-tablecell-body-member)|Gets the body object of the cell.|
||[cellIndex](/.tablecell#word-javascript/api/word/-tablecell-cellindex-member)|Gets the index of the cell in its row.|
||[columnWidth](/.tablecell#word-javascript/api/word/-tablecell-columnwidth-member)|Specifies the width of the cell's column in points.|
||[deleteColumn()](/.tablecell#word-javascript/api/word/-tablecell-deletecolumn-member(1))|Deletes the column containing this cell.|
||[deleteRow()](/.tablecell#word-javascript/api/word/-tablecell-deleterow-member(1))|Deletes the row containing this cell.|
||[getBorder(borderLocation: Word.BorderLocation)](/.tablecell#word-javascript/api/word/-tablecell-getborder-member(1))|Gets the border style for the specified border.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/.tablecell#word-javascript/api/word/-tablecell-getcellpadding-member(1))|Gets cell padding in points.|
||[getNext()](/.tablecell#word-javascript/api/word/-tablecell-getnext-member(1))|Gets the next cell.|
||[getNextOrNullObject()](/.tablecell#word-javascript/api/word/-tablecell-getnextornullobject-member(1))|Gets the next cell.|
||[horizontalAlignment](/.tablecell#word-javascript/api/word/-tablecell-horizontalalignment-member)|Specifies the horizontal alignment of the cell.|
||[insertColumns(insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After", columnCount: number, values?: string[][])](/.tablecell#word-javascript/api/word/-tablecell-insertcolumns-member(1))|Adds columns to the left or right of the cell, using the cell's column as a template.|
||[insertRows(insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After", rowCount: number, values?: string[][])](/.tablecell#word-javascript/api/word/-tablecell-insertrows-member(1))|Inserts rows above or below the cell, using the cell's row as a template.|
||[parentRow](/.tablecell#word-javascript/api/word/-tablecell-parentrow-member)|Gets the parent row of the cell.|
||[parentTable](/.tablecell#word-javascript/api/word/-tablecell-parenttable-member)|Gets the parent table of the cell.|
||[rowIndex](/.tablecell#word-javascript/api/word/-tablecell-rowindex-member)|Gets the index of the cell's row in the table.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/.tablecell#word-javascript/api/word/-tablecell-setcellpadding-member(1))|Sets cell padding in points.|
||[shadingColor](/.tablecell#word-javascript/api/word/-tablecell-shadingcolor-member)|Specifies the shading color of the cell.|
||[value](/.tablecell#word-javascript/api/word/-tablecell-value-member)|Specifies the text of the cell.|
||[verticalAlignment](/.tablecell#word-javascript/api/word/-tablecell-verticalalignment-member)|Specifies the vertical alignment of the cell.|
||[width](/.tablecell#word-javascript/api/word/-tablecell-width-member)|Gets the width of the cell in points.|
|[TableCellCollection](/.tablecellcollection)|[getFirst()](/.tablecellcollection#word-javascript/api/word/-tablecellcollection-getfirst-member(1))|Gets the first table cell in this collection.|
||[getFirstOrNullObject()](/.tablecellcollection#word-javascript/api/word/-tablecellcollection-getfirstornullobject-member(1))|Gets the first table cell in this collection.|
||[items](/.tablecellcollection#word-javascript/api/word/-tablecellcollection-items-member)|Gets the loaded child items in this collection.|
|[TableCollection](/.tablecollection)|[getFirst()](/.tablecollection#word-javascript/api/word/-tablecollection-getfirst-member(1))|Gets the first table in this collection.|
||[getFirstOrNullObject()](/.tablecollection#word-javascript/api/word/-tablecollection-getfirstornullobject-member(1))|Gets the first table in this collection.|
||[items](/.tablecollection#word-javascript/api/word/-tablecollection-items-member)|Gets the loaded child items in this collection.|
|[TableRow](/.tablerow)|[cellCount](/.tablerow#word-javascript/api/word/-tablerow-cellcount-member)|Gets the number of cells in the row.|
||[cells](/.tablerow#word-javascript/api/word/-tablerow-cells-member)|Gets cells.|
||[clear()](/.tablerow#word-javascript/api/word/-tablerow-clear-member(1))|Clears the contents of the row.|
||[delete()](/.tablerow#word-javascript/api/word/-tablerow-delete-member(1))|Deletes the entire row.|
||[font](/.tablerow#word-javascript/api/word/-tablerow-font-member)|Gets the font.|
||[getBorder(borderLocation: Word.BorderLocation)](/.tablerow#word-javascript/api/word/-tablerow-getborder-member(1))|Gets the border style of the cells in the row.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/.tablerow#word-javascript/api/word/-tablerow-getcellpadding-member(1))|Gets cell padding in points.|
||[getNext()](/.tablerow#word-javascript/api/word/-tablerow-getnext-member(1))|Gets the next row.|
||[getNextOrNullObject()](/.tablerow#word-javascript/api/word/-tablerow-getnextornullobject-member(1))|Gets the next row.|
||[horizontalAlignment](/.tablerow#word-javascript/api/word/-tablerow-horizontalalignment-member)|Specifies the horizontal alignment of every cell in the row.|
||[insertRows(insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After", rowCount: number, values?: string[][])](/.tablerow#word-javascript/api/word/-tablerow-insertrows-member(1))|Inserts rows using this row as a template.|
||[isHeader](/.tablerow#word-javascript/api/word/-tablerow-isheader-member)|Checks whether the row is a header row.|
||[parentTable](/.tablerow#word-javascript/api/word/-tablerow-parenttable-member)|Gets parent table.|
||[preferredHeight](/.tablerow#word-javascript/api/word/-tablerow-preferredheight-member)|Specifies the preferred height of the row in points.|
||[rowIndex](/.tablerow#word-javascript/api/word/-tablerow-rowindex-member)|Gets the index of the row in its parent table.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/.tablerow#word-javascript/api/word/-tablerow-search-member(1))|Performs a search with the specified SearchOptions on the scope of the row.|
||[select(selectionMode?: Word.SelectionMode)](/.tablerow#word-javascript/api/word/-tablerow-select-member(1))|Selects the row and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/.tablerow#word-javascript/api/word/-tablerow-setcellpadding-member(1))|Sets cell padding in points.|
||[shadingColor](/.tablerow#word-javascript/api/word/-tablerow-shadingcolor-member)|Specifies the shading color.|
||[values](/.tablerow#word-javascript/api/word/-tablerow-values-member)|Specifies the text values in the row, as a 2D JavaScript array.|
||[verticalAlignment](/.tablerow#word-javascript/api/word/-tablerow-verticalalignment-member)|Specifies the vertical alignment of the cells in the row.|
|[TableRowCollection](/.tablerowcollection)|[getFirst()](/.tablerowcollection#word-javascript/api/word/-tablerowcollection-getfirst-member(1))|Gets the first row in this collection.|
||[getFirstOrNullObject()](/.tablerowcollection#word-javascript/api/word/-tablerowcollection-getfirstornullobject-member(1))|Gets the first row in this collection.|
||[items](/.tablerowcollection#word-javascript/api/word/-tablerowcollection-items-member)|Gets the loaded child items in this collection.|
