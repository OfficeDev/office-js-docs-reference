| Class | Fields | Description |
|:---|:---|:---|
|[Annotation](/javascript/api/word/word.annotation)|||
|[Application](/javascript/api/word/word.application)|||
|[Body](/javascript/api/word/word.body)|[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.body#word-word-body-search-member(1))|Performs a search with the specified SearchOptions on the scope of the body object.|
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
|[CheckboxContentControl](/javascript/api/word/word.checkboxcontentcontrol)|||
|[Comment](/javascript/api/word/word.comment)|||
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|||
|[CommentReply](/javascript/api/word/word.commentreply)|||
|[ContentControl](/javascript/api/word/word.contentcontrol)|[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-search-member(1))|Performs a search with the specified SearchOptions on the scope of the content control object.|
|[CritiqueAnnotation](/javascript/api/word/word.critiqueannotation)|||
|[CustomProperty](/javascript/api/word/word.customproperty)|||
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[deleteAttribute(xpath: string, namespaceMappings: { [key: string]: string }, name: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteattribute-member(1))|Deletes an attribute with the given name from the element identified by xpath.|
||[deleteElement(xpath: string, namespaceMappings: { [key: string]: string })](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteelement-member(1))|Deletes the element identified by xpath.|
||[insertAttribute(xpath: string, namespaceMappings: { [key: string]: string }, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertattribute-member(1))|Inserts an attribute with the given name and value to the element identified by xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: { [key: string]: string }, index?: number)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertelement-member(1))|Inserts the given XML under the parent element identified by xpath at child position index.|
||[query(xpath: string, namespaceMappings: { [key: string]: string })](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-query-member(1))|Queries the XML content of the custom XML part.|
||[updateAttribute(xpath: string, namespaceMappings: { [key: string]: string }, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateattribute-member(1))|Updates the value of an attribute with the given name of the element identified by xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: { [key: string]: string })](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateelement-member(1))|Updates the XML of the element identified by xpath.|
|[Document](/javascript/api/word/word.document)|[compare(filePath: string, documentCompareOptions?: Word.DocumentCompareOptions)](/javascript/api/word/word.document#word-word-document-compare-member(1))|Displays revision marks that indicate where the specified document differs from another document.|
||[importStylesFromJson(stylesJson: string, importedStylesConflictBehavior?: Word.ImportedStylesConflictBehavior)](/javascript/api/word/word.document#word-word-document-importstylesfromjson-member(1))|Import styles from a JSON-formatted string.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.document#word-word-document-search-member(1))|Performs a search with the specified search options on the scope of the whole document.|
|[DocumentCompareOptions](/javascript/api/word/word.documentcompareoptions)|[addToRecentFiles](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-addtorecentfiles-member)|True adds the document to the list of recently used files on the File menu.|
||[authorName](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-authorname-member)|The reviewer name associated with the differences generated by the comparison.|
||[compareTarget](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-comparetarget-member)|The target document for the comparison.|
||[detectFormatChanges](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-detectformatchanges-member)|True (default) for the comparison to include detection of format changes.|
||[ignoreAllComparisonWarnings](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-ignoreallcomparisonwarnings-member)|True compares the documents without notifying a user of problems.|
||[removeDateAndTime](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-removedateandtime-member)|True removes date and time stamp information from tracked changes in the returned Document object.|
||[removePersonalInformation](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-removepersonalinformation-member)|True removes all user information from comments, revisions, and the properties dialog box in the returned Document object.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|||
|[DocumentProperties](/javascript/api/word/word.documentproperties)|||
|[Field](/javascript/api/word/word.field)|[showCodes](/javascript/api/word/word.field#word-word-field-showcodes-member)|Specifies whether the field codes are displayed for the specified field.|
|[Font](/javascript/api/word/word.font)|||
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|Gets the format of the inline image.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#word-word-list-getlevelfont-member(1))|Gets the font of the bullet, number, or picture at the specified level in the list.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#word-word-list-getlevelpicture-member(1))|Gets the Base64-encoded string representation of the picture at the specified level in the list.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|Resets the font of the bullet, number, or picture at the specified level in the list.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#word-word-list-setlevelpicture-member(1))|Sets the picture at the specified level in the list.|
|[ListItem](/javascript/api/word/word.listitem)|||
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
|[NoteItem](/javascript/api/word/word.noteitem)|||
|[Paragraph](/javascript/api/word/word.paragraph)|[getText(options?: Word.GetTextOptions \| { IncludeHiddenText?: boolean IncludeTextMarkedAsDeleted?: boolean })](/javascript/api/word/word.paragraph#word-word-paragraph-gettext-member(1))|Returns the text of the paragraph.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.paragraph#word-word-paragraph-search-member(1))|Performs a search with the specified SearchOptions on the scope of the paragraph object.|
|[ParagraphFormat](/javascript/api/word/word.paragraphformat)|||
|[Range](/javascript/api/word/word.range)|[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.range#word-word-range-search-member(1))|Performs a search with the specified SearchOptions on the scope of the range object.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|||
|[Section](/javascript/api/word/word.section)|||
|[Setting](/javascript/api/word/word.setting)|||
|[Shading](/javascript/api/word/word.shading)|[foregroundPatternColor](/javascript/api/word/word.shading#word-word-shading-foregroundpatterncolor-member)|Specifies the color for the foreground of the object.|
||[texture](/javascript/api/word/word.shading#word-word-shading-texture-member)|Specifies the shading texture of the object.|
|[Style](/javascript/api/word/word.style)|[borders](/javascript/api/word/word.style#word-word-style-borders-member)|Specifies a BorderCollection object that represents all the borders for the specified style.|
||[listTemplate](/javascript/api/word/word.style#word-word-style-listtemplate-member)|Gets a ListTemplate object that represents the list formatting for the specified Style object.|
|[Table](/javascript/api/word/word.table)|[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.table#word-word-table-search-member(1))|Performs a search with the specified SearchOptions on the scope of the table object.|
|[TableBorder](/javascript/api/word/word.tableborder)|||
|[TableCell](/javascript/api/word/word.tablecell)|||
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|Inserts a content control on the row.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#word-word-tablerow-search-member(1))|Performs a search with the specified SearchOptions on the scope of the row.|
|[TableStyle](/javascript/api/word/word.tablestyle)|[alignment](/javascript/api/word/word.tablestyle#word-word-tablestyle-alignment-member)|Specifies the table's alignment against the page margin.|
||[allowBreakAcrossPage](/javascript/api/word/word.tablestyle#word-word-tablestyle-allowbreakacrosspage-member)|Specifies whether lines in tables formatted with a specified style break across pages.|
|[TrackedChange](/javascript/api/word/word.trackedchange)|||
