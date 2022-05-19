| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[contentControls](/javascript/api/word/word.body#word-word-body-contentcontrols-member)|Gets the collection of rich text content control objects in the body.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#word-word-body-select-member(1))|Selects the body and navigates the Word UI to it.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[contentControls](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-contentcontrols-member)|Gets the collection of content control objects in the content control.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-select-member(1))|Selects the content control.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyid-member(1))|Gets a content control by its identifier.|
||[items](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[body](/javascript/api/word/word.document#word-word-document-body-member)|Gets the body object of the main document.|
||[contentControls](/javascript/api/word/word.document#word-word-document-contentcontrols-member)|Gets the collection of content control objects in the document.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#word-word-font-bold-member)|Gets or sets a value that indicates whether the font is bold.|
||[color](/javascript/api/word/word.font#word-word-font-color-member)|Gets or sets the color for the specified font.|
||[doubleStrikeThrough](/javascript/api/word/word.font#word-word-font-doublestrikethrough-member)|Gets or sets a value that indicates whether the font has a double strikethrough.|
||[highlightColor](/javascript/api/word/word.font#word-word-font-highlightcolor-member)|Gets or sets the highlight color.|
||[italic](/javascript/api/word/word.font#word-word-font-italic-member)|Gets or sets a value that indicates whether the font is italicized.|
||[name](/javascript/api/word/word.font#word-word-font-name-member)|Gets or sets a value that represents the name of the font.|
||[size](/javascript/api/word/word.font#word-word-font-size-member)|Gets or sets a value that represents the font size in points.|
||[strikeThrough](/javascript/api/word/word.font#word-word-font-strikethrough-member)|Gets or sets a value that indicates whether the font has a strikethrough.|
||[subscript](/javascript/api/word/word.font#word-word-font-subscript-member)|Gets or sets a value that indicates whether the font is a subscript.|
||[superscript](/javascript/api/word/word.font#word-word-font-superscript-member)|Gets or sets a value that indicates whether the font is a superscript.|
||[underline](/javascript/api/word/word.font#word-word-font-underline-member)|Gets or sets a value that indicates the font's underline type.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[items](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[contentControls](/javascript/api/word/word.paragraph#word-word-paragraph-contentcontrols-member)|Gets the collection of content control objects in the paragraph.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#word-word-paragraph-select-member(1))|Selects and navigates the Word UI to the paragraph.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[contentControls](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-contentcontrols-member)|Gets the collection of content control objects in the range.|
||[items](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-items-member)|Gets the loaded child items in this collection.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-select-member(1))|Selects and navigates the Word UI to the range.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[ignorePunct](/javascript/api/word/word.rangecollection#word-word-rangecollection-ignorepunct-member)|Gets or sets a value that indicates whether to ignore all punctuation characters between words.|
||[ignoreSpace](/javascript/api/word/word.rangecollection#word-word-rangecollection-ignorespace-member)|Gets or sets a value that indicates whether to ignore all whitespace between words.|
||[items](/javascript/api/word/word.rangecollection#word-word-rangecollection-items-member)|Gets the loaded child items in this collection.|
||[matchCase](/javascript/api/word/word.rangecollection#word-word-rangecollection-matchcase-member)|Gets or sets a value that indicates whether to perform a case sensitive search.|
||[matchPrefix](/javascript/api/word/word.rangecollection#word-word-rangecollection-matchprefix-member)|Gets or sets a value that indicates whether to match words that begin with the search string.|
||[matchSuffix](/javascript/api/word/word.rangecollection#word-word-rangecollection-matchsuffix-member)|Gets or sets a value that indicates whether to match words that end with the search string.|
||[matchWholeWord](/javascript/api/word/word.rangecollection#word-word-rangecollection-matchwholeword-member)|Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word.|
||[matchWildcards](/javascript/api/word/word.rangecollection#word-word-rangecollection-matchwildcards-member)|Gets or sets a value that indicates whether the search will be performed using special search operators.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[document](/javascript/api/word/word.requestcontext#word-word-requestcontext-document-member)|Executes a batch script that performs actions on the Word object model, using the RequestContext of previously created API objects.|
|[Section](/javascript/api/word/word.section)|[body](/javascript/api/word/word.section#word-word-section-body-member)|Gets the body object of the section.|
||[getFooter(type: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getfooter-member(1))|Gets one of the section's footers.|
||[getHeader(type: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getheader-member(1))|Gets one of the section's headers.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-items-member)|Gets the loaded child items in this collection.|
