| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[retrieveStylesFromBase64(base64File: string)](/javascript/api/word/word.application#word-word-application-retrievestylesfrombase64-member(1))|Parse styles from template Base64 file and return JSON format of retrieved styles as a string.|
|[Body](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#word-word-body-endnotes-member)|Gets the collection of endnotes in the body.|
||[footnotes](/javascript/api/word/word.body#word-word-body-footnotes-member)|Gets the collection of footnotes in the body.|
||[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.body#word-word-body-getcontentcontrols-member(1))|Gets the currently supported content controls in the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|Gets the collection of endnotes in the content control.|
||[footnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|Gets the collection of footnotes in the content control.|
||[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcontentcontrols-member(1))|Gets the currently supported child content controls in this content control.|
||[onDataChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member)|Occurs when data within the content control are changed.|
||[onDeleted](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member)|Occurs when the content control is deleted.|
||[onEntered](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onentered-member)|Occurs when the content control is entered.|
||[onExited](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onexited-member)|Occurs when the content control is exited, for example, when the cursor leaves the content control.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member)|Occurs when selection within the content control is changed.|
|[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs)|[eventType](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-source-member)|The source of the event.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByChangeTrackingStates(changeTrackingStates: Word.ChangeTrackingState[])](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbychangetrackingstates-member(1))|Gets the content controls that have the specified tracking state.|
|[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs)|[eventType](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-source-member)|The source of the event.|
|[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs)|[eventType](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-source-member)|The source of the event.|
|[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs)|[eventType](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-source-member)|The source of the event.|
|[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs)|[eventType](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-source-member)|The source of the event.|
|[ContentControlOptions](/javascript/api/word/word.contentcontroloptions)|[types](/javascript/api/word/word.contentcontroloptions#word-word-contentcontroloptions-types-member)|An array of content control types, item must be 'RichText', 'PlainText', 'CheckBox', 'DropDownList', or 'ComboBox'.|
|[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs)|[eventType](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-source-member)|The source of the event.|
|[Document](/javascript/api/word/word.document)|[addStyle(name: string, type: Word.StyleType)](/javascript/api/word/word.document#word-word-document-addstyle-member(1))|Adds a style into the document by name and type.|
||[close(closeBehavior?: Word.CloseBehavior)](/javascript/api/word/word.document#word-word-document-close-member(1))|Closes the current document.|
||[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.document#word-word-document-getcontentcontrols-member(1))|Gets the currently supported content controls in the document.|
||[getEndnoteBody()](/javascript/api/word/word.document#word-word-document-getendnotebody-member(1))|Gets the document's endnotes in a single body.|
||[getFootnoteBody()](/javascript/api/word/word.document#word-word-document-getfootnotebody-member(1))|Gets the document's footnotes in a single body.|
||[getStyles()](/javascript/api/word/word.document#word-word-document-getstyles-member(1))|Gets a StyleCollection object that represents the whole style set of the document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End", insertFileOptions?: Word.InsertFileOptions)](/javascript/api/word/word.document#word-word-document-insertfilefrombase64-member(1))|Inserts a document into the target document at a specific location with additional properties.|
||[onContentControlAdded](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member)|Occurs when a content control is added.|
|[Field](/javascript/api/word/word.field)|[data](/javascript/api/word/word.field#word-word-field-data-member)|Specifies data in an "Addin" field.|
||[delete()](/javascript/api/word/word.field#word-word-field-delete-member(1))|Deletes the field.|
||[kind](/javascript/api/word/word.field#word-word-field-kind-member)|Gets the field's kind.|
||[locked](/javascript/api/word/word.field#word-word-field-locked-member)|Specifies whether the field is locked.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.field#word-word-field-select-member(1))|Selects the field.|
||[type](/javascript/api/word/word.field#word-word-field-type-member)|Gets the field's type.|
||[updateResult()](/javascript/api/word/word.field#word-word-field-updateresult-member(1))|Updates the field.|
|[FieldCollection](/javascript/api/word/word.fieldcollection)|[getByTypes(types: Word.FieldType[])](/javascript/api/word/word.fieldcollection#word-word-fieldcollection-getbytypes-member(1))|Gets the Field object collection including the specified types of fields.|
|[InsertFileOptions](/javascript/api/word/word.insertfileoptions)|[importChangeTrackingMode](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importchangetrackingmode-member)|Represents whether the change tracking mode status from the source document should be imported.|
||[importPageColor](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importpagecolor-member)|Represents whether the page color and other background information from the source document should be imported.|
||[importParagraphSpacing](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importparagraphspacing-member)|Represents whether the paragraph spacing from the source document should be imported.|
||[importStyles](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importstyles-member)|Represents whether the styles from the source document should be imported.|
||[importTheme](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importtheme-member)|Represents whether the theme from the source document should be imported.|
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
||[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-getcontentcontrols-member(1))|Gets the currently supported content controls in the paragraph.|
|[ParagraphFormat](/javascript/api/word/word.paragraphformat)|[alignment](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-alignment-member)|Specifies the alignment for the specified paragraphs.|
||[firstLineIndent](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-firstlineindent-member)|Specifies the value (in points) for a first line or hanging indent.|
||[keepTogether](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-keeptogether-member)|Specifies whether all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.|
||[keepWithNext](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-keepwithnext-member)|Specifies whether the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.|
||[leftIndent](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-leftindent-member)|Specifies the left indent.|
||[lineSpacing](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-linespacing-member)|Specifies the line spacing (in points) for the specified paragraphs.|
||[lineUnitAfter](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-lineunitafter-member)|Specifies the amount of spacing (in gridlines) after the specified paragraphs.|
||[lineUnitBefore](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-lineunitbefore-member)|Specifies the amount of spacing (in gridlines) before the specified paragraphs.|
||[mirrorIndents](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-mirrorindents-member)|Specifies whether left and right indents are the same width.|
||[outlineLevel](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-outlinelevel-member)|Specifies the outline level for the specified paragraphs.|
||[rightIndent](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-rightindent-member)|Specifies the right indent (in points) for the specified paragraphs.|
||[spaceAfter](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-spaceafter-member)|Specifies the amount of spacing (in points) after the specified paragraph or text column.|
||[spaceBefore](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-spacebefore-member)|Specifies the spacing (in points) before the specified paragraphs.|
||[widowControl](/javascript/api/word/word.paragraphformat#word-word-paragraphformat-widowcontrol-member)|Specifies whether the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Microsoft Word repaginates the document.|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#word-word-range-endnotes-member)|Gets the collection of endnotes in the range.|
||[footnotes](/javascript/api/word/word.range#word-word-range-footnotes-member)|Gets the collection of footnotes in the range.|
||[getContentControls(options?: Word.ContentControlOptions)](/javascript/api/word/word.range#word-word-range-getcontentcontrols-member(1))|Gets the currently supported content controls in the range.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertendnote-member(1))|Inserts an endnote.|
||[insertField(insertLocation: Word.InsertLocation \| "Replace" \| "Start" \| "End" \| "Before" \| "After", fieldType?: Word.FieldType, text?: string, removeFormatting?: boolean)](/javascript/api/word/word.range#word-word-range-insertfield-member(1))|Inserts a field at the specified location.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertfootnote-member(1))|Inserts a footnote.|
|[Style](/javascript/api/word/word.style)|[baseStyle](/javascript/api/word/word.style#word-word-style-basestyle-member)|Specifies the name of an existing style to use as the base formatting of another style.|
||[builtIn](/javascript/api/word/word.style#word-word-style-builtin-member)|Gets whether the specified style is a built-in style.|
||[delete()](/javascript/api/word/word.style#word-word-style-delete-member(1))|Deletes the style.|
||[font](/javascript/api/word/word.style#word-word-style-font-member)|Gets a font object that represents the character formatting of the specified style.|
||[inUse](/javascript/api/word/word.style#word-word-style-inuse-member)|Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.|
||[linked](/javascript/api/word/word.style#word-word-style-linked-member)|Gets whether a style is a linked style that can be used for both paragraph and character formatting.|
||[nameLocal](/javascript/api/word/word.style#word-word-style-namelocal-member)|Gets the name of a style in the language of the user.|
||[nextParagraphStyle](/javascript/api/word/word.style#word-word-style-nextparagraphstyle-member)|Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.|
||[paragraphFormat](/javascript/api/word/word.style#word-word-style-paragraphformat-member)|Gets a ParagraphFormat object that represents the paragraph settings for the specified style.|
||[priority](/javascript/api/word/word.style#word-word-style-priority-member)|Specifies the priority.|
||[quickStyle](/javascript/api/word/word.style#word-word-style-quickstyle-member)|Specifies whether the style corresponds to an available quick style.|
||[type](/javascript/api/word/word.style#word-word-style-type-member)|Gets the style type.|
||[unhideWhenUsed](/javascript/api/word/word.style#word-word-style-unhidewhenused-member)|Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.|
||[visibility](/javascript/api/word/word.style#word-word-style-visibility-member)|Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.|
|[StyleCollection](/javascript/api/word/word.stylecollection)|[getByName(name: string)](/javascript/api/word/word.stylecollection#word-word-stylecollection-getbyname-member(1))|Get the style object by its name.|
||[getByNameOrNullObject(name: string)](/javascript/api/word/word.stylecollection#word-word-stylecollection-getbynameornullobject-member(1))|If the corresponding style doesn't exist, then this method returns an object with its `isNullObject` property set to `true`.|
||[getCount()](/javascript/api/word/word.stylecollection#word-word-stylecollection-getcount-member(1))|Gets the number of the styles in the collection.|
||[getItem(index: number)](/javascript/api/word/word.stylecollection#word-word-stylecollection-getitem-member(1))|Gets a style object by its index in the collection.|
||[items](/javascript/api/word/word.stylecollection#word-word-stylecollection-items-member)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#word-word-table-endnotes-member)|Gets the collection of endnotes in the table.|
||[footnotes](/javascript/api/word/word.table#word-word-table-footnotes-member)|Gets the collection of footnotes in the table.|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|Gets the collection of endnotes in the table row.|
||[footnotes](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|Gets the collection of footnotes in the table row.|
