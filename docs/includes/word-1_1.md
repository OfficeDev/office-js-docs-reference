| Class | Fields | Description |
|:---|:---|:---|
|[Body](/.body)|[clear()](/.body#word-javascript/api/word/-body-clear-member(1))|Clears the contents of the body object.|
||[contentControls](/.body#word-javascript/api/word/-body-contentcontrols-member)|Gets the collection of rich text content control objects in the body.|
||[font](/.body#word-javascript/api/word/-body-font-member)|Gets the text format of the body.|
||[getHtml()](/.body#word-javascript/api/word/-body-gethtml-member(1))|Gets an HTML representation of the body object.|
||[getOoxml()](/.body#word-javascript/api/word/-body-getooxml-member(1))|Gets the OOXML (Office Open XML) representation of the body object.|
||[inlinePictures](/.body#word-javascript/api/word/-body-inlinepictures-member)|Gets the collection of InlinePicture objects in the body.|
||[insertBreak(breakType: Word.BreakType \| "Page" \| "Next" \| "SectionNext" \| "SectionContinuous" \| "SectionEven" \| "SectionOdd" \| "Line", insertLocation: Word.InsertLocation.start \| Word.InsertLocation.end \| "Start" \| "End")](/.body#word-javascript/api/word/-body-insertbreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| Word.ContentControlType.checkBox \| Word.ContentControlType.dropDownList \| Word.ContentControlType.comboBox \| "RichText" \| "PlainText" \| "CheckBox" \| "DropDownList" \| "ComboBox")](/.body#word-javascript/api/word/-body-insertcontentcontrol-member(1))|Wraps the Body object with a content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.body#word-javascript/api/word/-body-insertfilefrombase64-member(1))|Inserts a document into the body at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.body#word-javascript/api/word/-body-inserthtml-member(1))|Inserts HTML at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.body#word-javascript/api/word/-body-insertooxml-member(1))|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.start \| Word.InsertLocation.end \| "Start" \| "End")](/.body#word-javascript/api/word/-body-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.body#word-javascript/api/word/-body-inserttext-member(1))|Inserts text into the body at the specified location.|
||[paragraphs](/.body#word-javascript/api/word/-body-paragraphs-member)|Gets the collection of paragraph objects in the body.|
||[parentContentControl](/.body#word-javascript/api/word/-body-parentcontentcontrol-member)|Gets the content control that contains the body.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/.body#word-javascript/api/word/-body-search-member(1))|Performs a search with the specified SearchOptions on the scope of the body object.|
||[select(selectionMode?: Word.SelectionMode)](/.body#word-javascript/api/word/-body-select-member(1))|Selects the body and navigates the Word UI to it.|
||[style](/.body#word-javascript/api/word/-body-style-member)|Specifies the style name for the body.|
||[text](/.body#word-javascript/api/word/-body-text-member)|Gets the text of the body.|
|[ContentControl](/.contentcontrol)|[appearance](/.contentcontrol#word-javascript/api/word/-contentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[cannotDelete](/.contentcontrol#word-javascript/api/word/-contentcontrol-cannotdelete-member)|Specifies a value that indicates whether the user can delete the content control.|
||[cannotEdit](/.contentcontrol#word-javascript/api/word/-contentcontrol-cannotedit-member)|Specifies a value that indicates whether the user can edit the contents of the content control.|
||[clear()](/.contentcontrol#word-javascript/api/word/-contentcontrol-clear-member(1))|Clears the contents of the content control.|
||[color](/.contentcontrol#word-javascript/api/word/-contentcontrol-color-member)|Specifies the color of the content control.|
||[contentControls](/.contentcontrol#word-javascript/api/word/-contentcontrol-contentcontrols-member)|Gets the collection of content control objects in the content control.|
||[delete(keepContent: boolean)](/.contentcontrol#word-javascript/api/word/-contentcontrol-delete-member(1))|Deletes the content control and its content.|
||[font](/.contentcontrol#word-javascript/api/word/-contentcontrol-font-member)|Gets the text format of the content control.|
||[getHtml()](/.contentcontrol#word-javascript/api/word/-contentcontrol-gethtml-member(1))|Gets an HTML representation of the content control object.|
||[getOoxml()](/.contentcontrol#word-javascript/api/word/-contentcontrol-getooxml-member(1))|Gets the Office Open XML (OOXML) representation of the content control object.|
||[id](/.contentcontrol#word-javascript/api/word/-contentcontrol-id-member)|Gets an integer that represents the content control identifier.|
||[inlinePictures](/.contentcontrol#word-javascript/api/word/-contentcontrol-inlinepictures-member)|Gets the collection of InlinePicture objects in the content control.|
||[insertBreak(breakType: Word.BreakType \| "Page" \| "Next" \| "SectionNext" \| "SectionContinuous" \| "SectionEven" \| "SectionOdd" \| "Line", insertLocation: Word.InsertLocation.start \| Word.InsertLocation.end \| Word.InsertLocation.before \| Word.InsertLocation.after \| "Start" \| "End" \| "Before" \| "After")](/.contentcontrol#word-javascript/api/word/-contentcontrol-insertbreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.contentcontrol#word-javascript/api/word/-contentcontrol-insertfilefrombase64-member(1))|Inserts a document into the content control at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.contentcontrol#word-javascript/api/word/-contentcontrol-inserthtml-member(1))|Inserts HTML into the content control at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.contentcontrol#word-javascript/api/word/-contentcontrol-insertooxml-member(1))|Inserts OOXML into the content control at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.start \| Word.InsertLocation.end \| Word.InsertLocation.before \| Word.InsertLocation.after \| "Start" \| "End" \| "Before" \| "After")](/.contentcontrol#word-javascript/api/word/-contentcontrol-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.contentcontrol#word-javascript/api/word/-contentcontrol-inserttext-member(1))|Inserts text into the content control at the specified location.|
||[paragraphs](/.contentcontrol#word-javascript/api/word/-contentcontrol-paragraphs-member)|Gets the collection of paragraph objects in the content control.|
||[parentContentControl](/.contentcontrol#word-javascript/api/word/-contentcontrol-parentcontentcontrol-member)|Gets the content control that contains the content control.|
||[placeholderText](/.contentcontrol#word-javascript/api/word/-contentcontrol-placeholdertext-member)|Specifies the placeholder text of the content control.|
||[removeWhenEdited](/.contentcontrol#word-javascript/api/word/-contentcontrol-removewhenedited-member)|Specifies a value that indicates whether the content control is removed after it is edited.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/.contentcontrol#word-javascript/api/word/-contentcontrol-search-member(1))|Performs a search with the specified SearchOptions on the scope of the content control object.|
||[select(selectionMode?: Word.SelectionMode)](/.contentcontrol#word-javascript/api/word/-contentcontrol-select-member(1))|Selects the content control.|
||[style](/.contentcontrol#word-javascript/api/word/-contentcontrol-style-member)|Specifies the style name for the content control.|
||[tag](/.contentcontrol#word-javascript/api/word/-contentcontrol-tag-member)|Specifies a tag to identify a content control.|
||[text](/.contentcontrol#word-javascript/api/word/-contentcontrol-text-member)|Gets the text of the content control.|
||[title](/.contentcontrol#word-javascript/api/word/-contentcontrol-title-member)|Specifies the title for a content control.|
||[type](/.contentcontrol#word-javascript/api/word/-contentcontrol-type-member)|Gets the content control type.|
|[ContentControlCollection](/.contentcontrolcollection)|[getById(id: number)](/.contentcontrolcollection#word-javascript/api/word/-contentcontrolcollection-getbyid-member(1))|Gets a content control by its identifier.|
||[getByTag(tag: string)](/.contentcontrolcollection#word-javascript/api/word/-contentcontrolcollection-getbytag-member(1))|Gets the content controls that have the specified tag.|
||[getByTitle(title: string)](/.contentcontrolcollection#word-javascript/api/word/-contentcontrolcollection-getbytitle-member(1))|Gets the content controls that have the specified title.|
||[getItem(id: number)](/.contentcontrolcollection#word-javascript/api/word/-contentcontrolcollection-getitem-member(1))|Gets a content control by its ID.|
||[items](/.contentcontrolcollection#word-javascript/api/word/-contentcontrolcollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/.document)|[body](/.document#word-javascript/api/word/-document-body-member)|Gets the body object of the main document.|
||[contentControls](/.document#word-javascript/api/word/-document-contentcontrols-member)|Gets the collection of content control objects in the document.|
||[getSelection()](/.document#word-javascript/api/word/-document-getselection-member(1))|Gets the current selection of the document.|
||[save(saveBehavior?: Word.SaveBehavior, fileName?: string)](/.document#word-javascript/api/word/-document-save-member(1))|Saves the document.|
||[saved](/.document#word-javascript/api/word/-document-saved-member)|Indicates whether the changes in the document have been saved.|
||[sections](/.document#word-javascript/api/word/-document-sections-member)|Gets the collection of section objects in the document.|
|[Font](/.font)|[bold](/.font#word-javascript/api/word/-font-bold-member)|Specifies a value that indicates whether the font is bold.|
||[color](/.font#word-javascript/api/word/-font-color-member)|Specifies the color for the specified font.|
||[doubleStrikeThrough](/.font#word-javascript/api/word/-font-doublestrikethrough-member)|Specifies a value that indicates whether the font has a double strikethrough.|
||[highlightColor](/.font#word-javascript/api/word/-font-highlightcolor-member)|Specifies the highlight color.|
||[italic](/.font#word-javascript/api/word/-font-italic-member)|Specifies a value that indicates whether the font is italicized.|
||[name](/.font#word-javascript/api/word/-font-name-member)|Specifies a value that represents the name of the font.|
||[size](/.font#word-javascript/api/word/-font-size-member)|Specifies a value that represents the font size in points.|
||[strikeThrough](/.font#word-javascript/api/word/-font-strikethrough-member)|Specifies a value that indicates whether the font has a strikethrough.|
||[subscript](/.font#word-javascript/api/word/-font-subscript-member)|Specifies a value that indicates whether the font is a subscript.|
||[superscript](/.font#word-javascript/api/word/-font-superscript-member)|Specifies a value that indicates whether the font is a superscript.|
||[underline](/.font#word-javascript/api/word/-font-underline-member)|Specifies a value that indicates the font's underline type.|
|[InlinePicture](/.inlinepicture)|[altTextDescription](/.inlinepicture#word-javascript/api/word/-inlinepicture-alttextdescription-member)|Specifies a string that represents the alternative text associated with the inline image.|
||[altTextTitle](/.inlinepicture#word-javascript/api/word/-inlinepicture-alttexttitle-member)|Specifies a string that contains the title for the inline image.|
||[getBase64ImageSrc()](/.inlinepicture#word-javascript/api/word/-inlinepicture-getbase64imagesrc-member(1))|Gets the Base64-encoded string representation of the inline image.|
||[height](/.inlinepicture#word-javascript/api/word/-inlinepicture-height-member)|Specifies a number that describes the height of the inline image.|
||[hyperlink](/.inlinepicture#word-javascript/api/word/-inlinepicture-hyperlink-member)|Specifies a hyperlink on the image.|
||[insertContentControl()](/.inlinepicture#word-javascript/api/word/-inlinepicture-insertcontentcontrol-member(1))|Wraps the inline picture with a rich text content control.|
||[lockAspectRatio](/.inlinepicture#word-javascript/api/word/-inlinepicture-lockaspectratio-member)|Specifies a value that indicates whether the inline image retains its original proportions when you resize it.|
||[parentContentControl](/.inlinepicture#word-javascript/api/word/-inlinepicture-parentcontentcontrol-member)|Gets the content control that contains the inline image.|
||[width](/.inlinepicture#word-javascript/api/word/-inlinepicture-width-member)|Specifies a number that describes the width of the inline image.|
|[InlinePictureCollection](/.inlinepicturecollection)|[items](/.inlinepicturecollection#word-javascript/api/word/-inlinepicturecollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/.paragraph)|[alignment](/.paragraph#word-javascript/api/word/-paragraph-alignment-member)|Specifies the alignment for a paragraph.|
||[clear()](/.paragraph#word-javascript/api/word/-paragraph-clear-member(1))|Clears the contents of the paragraph object.|
||[contentControls](/.paragraph#word-javascript/api/word/-paragraph-contentcontrols-member)|Gets the collection of content control objects in the paragraph.|
||[delete()](/.paragraph#word-javascript/api/word/-paragraph-delete-member(1))|Deletes the paragraph and its content from the document.|
||[firstLineIndent](/.paragraph#word-javascript/api/word/-paragraph-firstlineindent-member)|Specifies the value, in points, for a first line or hanging indent.|
||[font](/.paragraph#word-javascript/api/word/-paragraph-font-member)|Gets the text format of the paragraph.|
||[getHtml()](/.paragraph#word-javascript/api/word/-paragraph-gethtml-member(1))|Gets an HTML representation of the paragraph object.|
||[getOoxml()](/.paragraph#word-javascript/api/word/-paragraph-getooxml-member(1))|Gets the Office Open XML (OOXML) representation of the paragraph object.|
||[inlinePictures](/.paragraph#word-javascript/api/word/-paragraph-inlinepictures-member)|Gets the collection of InlinePicture objects in the paragraph.|
||[insertBreak(breakType: Word.BreakType \| "Page" \| "Next" \| "SectionNext" \| "SectionContinuous" \| "SectionEven" \| "SectionOdd" \| "Line", insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After")](/.paragraph#word-javascript/api/word/-paragraph-insertbreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| Word.ContentControlType.checkBox \| Word.ContentControlType.dropDownList \| Word.ContentControlType.comboBox \| "RichText" \| "PlainText" \| "CheckBox" \| "DropDownList" \| "ComboBox")](/.paragraph#word-javascript/api/word/-paragraph-insertcontentcontrol-member(1))|Wraps the Paragraph object with a content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.paragraph#word-javascript/api/word/-paragraph-insertfilefrombase64-member(1))|Inserts a document into the paragraph at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.paragraph#word-javascript/api/word/-paragraph-inserthtml-member(1))|Inserts HTML into the paragraph at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.paragraph#word-javascript/api/word/-paragraph-insertinlinepicturefrombase64-member(1))|Inserts a picture into the paragraph at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.paragraph#word-javascript/api/word/-paragraph-insertooxml-member(1))|Inserts OOXML into the paragraph at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After")](/.paragraph#word-javascript/api/word/-paragraph-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End")](/.paragraph#word-javascript/api/word/-paragraph-inserttext-member(1))|Inserts text into the paragraph at the specified location.|
||[leftIndent](/.paragraph#word-javascript/api/word/-paragraph-leftindent-member)|Specifies the left indent value, in points, for the paragraph.|
||[lineSpacing](/.paragraph#word-javascript/api/word/-paragraph-linespacing-member)|Specifies the line spacing, in points, for the specified paragraph.|
||[lineUnitAfter](/.paragraph#word-javascript/api/word/-paragraph-lineunitafter-member)|Specifies the amount of spacing, in grid lines, after the paragraph.|
||[lineUnitBefore](/.paragraph#word-javascript/api/word/-paragraph-lineunitbefore-member)|Specifies the amount of spacing, in grid lines, before the paragraph.|
||[outlineLevel](/.paragraph#word-javascript/api/word/-paragraph-outlinelevel-member)|Specifies the outline level for the paragraph.|
||[parentContentControl](/.paragraph#word-javascript/api/word/-paragraph-parentcontentcontrol-member)|Gets the content control that contains the paragraph.|
||[rightIndent](/.paragraph#word-javascript/api/word/-paragraph-rightindent-member)|Specifies the right indent value, in points, for the paragraph.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/.paragraph#word-javascript/api/word/-paragraph-search-member(1))|Performs a search with the specified SearchOptions on the scope of the paragraph object.|
||[select(selectionMode?: Word.SelectionMode)](/.paragraph#word-javascript/api/word/-paragraph-select-member(1))|Selects and navigates the Word UI to the paragraph.|
||[spaceAfter](/.paragraph#word-javascript/api/word/-paragraph-spaceafter-member)|Specifies the spacing, in points, after the paragraph.|
||[spaceBefore](/.paragraph#word-javascript/api/word/-paragraph-spacebefore-member)|Specifies the spacing, in points, before the paragraph.|
||[style](/.paragraph#word-javascript/api/word/-paragraph-style-member)|Specifies the style name for the paragraph.|
||[text](/.paragraph#word-javascript/api/word/-paragraph-text-member)|Gets the text of the paragraph.|
|[ParagraphCollection](/.paragraphcollection)|[items](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-items-member)|Gets the loaded child items in this collection.|
|[Range](/.range)|[clear()](/.range#word-javascript/api/word/-range-clear-member(1))|Clears the contents of the range object.|
||[contentControls](/.range#word-javascript/api/word/-range-contentcontrols-member)|Gets the collection of content control objects in the range.|
||[delete()](/.range#word-javascript/api/word/-range-delete-member(1))|Deletes the range and its content from the document.|
||[font](/.range#word-javascript/api/word/-range-font-member)|Gets the text format of the range.|
||[getHtml()](/.range#word-javascript/api/word/-range-gethtml-member(1))|Gets an HTML representation of the range object.|
||[getOoxml()](/.range#word-javascript/api/word/-range-getooxml-member(1))|Gets the OOXML representation of the range object.|
||[insertBreak(breakType: Word.BreakType \| "Page" \| "Next" \| "SectionNext" \| "SectionContinuous" \| "SectionEven" \| "SectionOdd" \| "Line", insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After")](/.range#word-javascript/api/word/-range-insertbreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| Word.ContentControlType.checkBox \| Word.ContentControlType.dropDownList \| Word.ContentControlType.comboBox \| "RichText" \| "PlainText" \| "CheckBox" \| "DropDownList" \| "ComboBox")](/.range#word-javascript/api/word/-range-insertcontentcontrol-member(1))|Wraps the Range object with a content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation \| "Replace" \| "Start" \| "End" \| "Before" \| "After")](/.range#word-javascript/api/word/-range-insertfilefrombase64-member(1))|Inserts a document at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation \| "Replace" \| "Start" \| "End" \| "Before" \| "After")](/.range#word-javascript/api/word/-range-inserthtml-member(1))|Inserts HTML at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation \| "Replace" \| "Start" \| "End" \| "Before" \| "After")](/.range#word-javascript/api/word/-range-insertooxml-member(1))|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before \| Word.InsertLocation.after \| "Before" \| "After")](/.range#word-javascript/api/word/-range-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation \| "Replace" \| "Start" \| "End" \| "Before" \| "After")](/.range#word-javascript/api/word/-range-inserttext-member(1))|Inserts text at the specified location.|
||[paragraphs](/.range#word-javascript/api/word/-range-paragraphs-member)|Gets the collection of paragraph objects in the range.|
||[parentContentControl](/.range#word-javascript/api/word/-range-parentcontentcontrol-member)|Gets the currently supported content control that contains the range.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/.range#word-javascript/api/word/-range-search-member(1))|Performs a search with the specified SearchOptions on the scope of the range object.|
||[select(selectionMode?: Word.SelectionMode)](/.range#word-javascript/api/word/-range-select-member(1))|Selects and navigates the Word UI to the range.|
||[style](/.range#word-javascript/api/word/-range-style-member)|Specifies the style name for the range.|
||[text](/.range#word-javascript/api/word/-range-text-member)|Gets the text of the range.|
|[RangeCollection](/.rangecollection)|[items](/.rangecollection#word-javascript/api/word/-rangecollection-items-member)|Gets the loaded child items in this collection.|
|[SearchOptions](/.searchoptions)|[ignorePunct](/.searchoptions#word-javascript/api/word/-searchoptions-ignorepunct-member)|Specifies a value that indicates whether to ignore all punctuation characters between words.|
||[ignoreSpace](/.searchoptions#word-javascript/api/word/-searchoptions-ignorespace-member)|Specifies a value that indicates whether to ignore all whitespace between words.|
||[matchCase](/.searchoptions#word-javascript/api/word/-searchoptions-matchcase-member)|Specifies a value that indicates whether to perform a case sensitive search.|
||[matchPrefix](/.searchoptions#word-javascript/api/word/-searchoptions-matchprefix-member)|Specifies a value that indicates whether to match words that begin with the search string.|
||[matchSuffix](/.searchoptions#word-javascript/api/word/-searchoptions-matchsuffix-member)|Specifies a value that indicates whether to match words that end with the search string.|
||[matchWholeWord](/.searchoptions#word-javascript/api/word/-searchoptions-matchwholeword-member)|Specifies a value that indicates whether to find operation only entire words, not text that is part of a larger word.|
||[matchWildcards](/.searchoptions#word-javascript/api/word/-searchoptions-matchwildcards-member)|Specifies a value that indicates whether the search will be performed using special search operators.|
|[Section](/.section)|[body](/.section#word-javascript/api/word/-section-body-member)|Gets the body object of the section.|
||[getFooter(type: Word.HeaderFooterType)](/.section#word-javascript/api/word/-section-getfooter-member(1))|Gets one of the section's footers.|
||[getHeader(type: Word.HeaderFooterType)](/.section#word-javascript/api/word/-section-getheader-member(1))|Gets one of the section's headers.|
|[SectionCollection](/.sectioncollection)|[items](/.sectioncollection#word-javascript/api/word/-sectioncollection-items-member)|Gets the loaded child items in this collection.|
