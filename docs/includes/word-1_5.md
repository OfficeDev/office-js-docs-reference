| Class | Fields | Description |
|:---|:---|:---|
|[Application](/.application)|[retrieveStylesFromBase64(base64File: string)](/.application#word-javascript/api/word/-application-retrievestylesfrombase64-member(1))|Parse styles from template Base64 file and return JSON format of retrieved styles as a string.|
|[Body](/.body)|[endnotes](/.body#word-javascript/api/word/-body-endnotes-member)|Gets the collection of endnotes in the body.|
||[footnotes](/.body#word-javascript/api/word/-body-footnotes-member)|Gets the collection of footnotes in the body.|
||[getContentControls(options?: Word.ContentControlOptions)](/.body#word-javascript/api/word/-body-getcontentcontrols-member(1))|Gets the currently supported content controls in the body.|
|[ContentControl](/.contentcontrol)|[endnotes](/.contentcontrol#word-javascript/api/word/-contentcontrol-endnotes-member)|Gets the collection of endnotes in the content control.|
||[footnotes](/.contentcontrol#word-javascript/api/word/-contentcontrol-footnotes-member)|Gets the collection of footnotes in the content control.|
||[getContentControls(options?: Word.ContentControlOptions)](/.contentcontrol#word-javascript/api/word/-contentcontrol-getcontentcontrols-member(1))|Gets the currently supported child content controls in this content control.|
||[onDataChanged](/.contentcontrol#word-javascript/api/word/-contentcontrol-ondatachanged-member)|Occurs when data within the content control are changed.|
||[onDeleted](/.contentcontrol#word-javascript/api/word/-contentcontrol-ondeleted-member)|Occurs when the content control is deleted.|
||[onEntered](/.contentcontrol#word-javascript/api/word/-contentcontrol-onentered-member)|Occurs when the content control is entered.|
||[onExited](/.contentcontrol#word-javascript/api/word/-contentcontrol-onexited-member)|Occurs when the content control is exited, for example, when the cursor leaves the content control.|
||[onSelectionChanged](/.contentcontrol#word-javascript/api/word/-contentcontrol-onselectionchanged-member)|Occurs when selection within the content control is changed.|
|[ContentControlAddedEventArgs](/.contentcontroladdedeventargs)|[eventType](/.contentcontroladdedeventargs#word-javascript/api/word/-contentcontroladdedeventargs-eventtype-member)|The event type.|
||[ids](/.contentcontroladdedeventargs#word-javascript/api/word/-contentcontroladdedeventargs-ids-member)|Gets the content control IDs.|
||[source](/.contentcontroladdedeventargs#word-javascript/api/word/-contentcontroladdedeventargs-source-member)|The source of the event.|
|[ContentControlCollection](/.contentcontrolcollection)|[getByChangeTrackingStates(changeTrackingStates: Word.ChangeTrackingState[])](/.contentcontrolcollection#word-javascript/api/word/-contentcontrolcollection-getbychangetrackingstates-member(1))|Gets the content controls that have the specified tracking state.|
|[ContentControlDataChangedEventArgs](/.contentcontroldatachangedeventargs)|[eventType](/.contentcontroldatachangedeventargs#word-javascript/api/word/-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
||[ids](/.contentcontroldatachangedeventargs#word-javascript/api/word/-contentcontroldatachangedeventargs-ids-member)|Gets the content control IDs.|
||[source](/.contentcontroldatachangedeventargs#word-javascript/api/word/-contentcontroldatachangedeventargs-source-member)|The source of the event.|
|[ContentControlDeletedEventArgs](/.contentcontroldeletedeventargs)|[eventType](/.contentcontroldeletedeventargs#word-javascript/api/word/-contentcontroldeletedeventargs-eventtype-member)|The event type.|
||[ids](/.contentcontroldeletedeventargs#word-javascript/api/word/-contentcontroldeletedeventargs-ids-member)|Gets the content control IDs.|
||[source](/.contentcontroldeletedeventargs#word-javascript/api/word/-contentcontroldeletedeventargs-source-member)|The source of the event.|
|[ContentControlEnteredEventArgs](/.contentcontrolenteredeventargs)|[eventType](/.contentcontrolenteredeventargs#word-javascript/api/word/-contentcontrolenteredeventargs-eventtype-member)|The event type.|
||[ids](/.contentcontrolenteredeventargs#word-javascript/api/word/-contentcontrolenteredeventargs-ids-member)|Gets the content control IDs.|
||[source](/.contentcontrolenteredeventargs#word-javascript/api/word/-contentcontrolenteredeventargs-source-member)|The source of the event.|
|[ContentControlExitedEventArgs](/.contentcontrolexitedeventargs)|[eventType](/.contentcontrolexitedeventargs#word-javascript/api/word/-contentcontrolexitedeventargs-eventtype-member)|The event type.|
||[ids](/.contentcontrolexitedeventargs#word-javascript/api/word/-contentcontrolexitedeventargs-ids-member)|Gets the content control IDs.|
||[source](/.contentcontrolexitedeventargs#word-javascript/api/word/-contentcontrolexitedeventargs-source-member)|The source of the event.|
|[ContentControlOptions](/.contentcontroloptions)|[types](/.contentcontroloptions#word-javascript/api/word/-contentcontroloptions-types-member)|An array of content control types, item must be 'RichText', 'PlainText', 'CheckBox', 'DropDownList', or 'ComboBox'.|
|[ContentControlSelectionChangedEventArgs](/.contentcontrolselectionchangedeventargs)|[eventType](/.contentcontrolselectionchangedeventargs#word-javascript/api/word/-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
||[ids](/.contentcontrolselectionchangedeventargs#word-javascript/api/word/-contentcontrolselectionchangedeventargs-ids-member)|Gets the content control IDs.|
||[source](/.contentcontrolselectionchangedeventargs#word-javascript/api/word/-contentcontrolselectionchangedeventargs-source-member)|The source of the event.|
|[Document](/.document)|[addStyle(name: string, type: Word.StyleType)](/.document#word-javascript/api/word/-document-addstyle-member(1))|Adds a style into the document by name and type.|
||[close(closeBehavior?: Word.CloseBehavior)](/.document#word-javascript/api/word/-document-close-member(1))|Closes the current document.|
||[getContentControls(options?: Word.ContentControlOptions)](/.document#word-javascript/api/word/-document-getcontentcontrols-member(1))|Gets the currently supported content controls in the document.|
||[getEndnoteBody()](/.document#word-javascript/api/word/-document-getendnotebody-member(1))|Gets the document's endnotes in a single body.|
||[getFootnoteBody()](/.document#word-javascript/api/word/-document-getfootnotebody-member(1))|Gets the document's footnotes in a single body.|
||[getStyles()](/.document#word-javascript/api/word/-document-getstyles-member(1))|Gets a StyleCollection object that represents the whole style set of the document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace \| Word.InsertLocation.start \| Word.InsertLocation.end \| "Replace" \| "Start" \| "End", insertFileOptions?: Word.InsertFileOptions)](/.document#word-javascript/api/word/-document-insertfilefrombase64-member(1))|Inserts a document into the target document at a specific location with additional properties.|
||[onContentControlAdded](/.document#word-javascript/api/word/-document-oncontentcontroladded-member)|Occurs when a content control is added.|
|[Field](/.field)|[data](/.field#word-javascript/api/word/-field-data-member)|Specifies data in an "Addin" field.|
||[delete()](/.field#word-javascript/api/word/-field-delete-member(1))|Deletes the field.|
||[kind](/.field#word-javascript/api/word/-field-kind-member)|Gets the field's kind.|
||[locked](/.field#word-javascript/api/word/-field-locked-member)|Specifies whether the field is locked.|
||[select(selectionMode?: Word.SelectionMode)](/.field#word-javascript/api/word/-field-select-member(1))|Selects the field.|
||[type](/.field#word-javascript/api/word/-field-type-member)|Gets the field's type.|
||[updateResult()](/.field#word-javascript/api/word/-field-updateresult-member(1))|Updates the field.|
|[FieldCollection](/.fieldcollection)|[getByTypes(types: Word.FieldType[])](/.fieldcollection#word-javascript/api/word/-fieldcollection-getbytypes-member(1))|Gets the Field object collection including the specified types of fields.|
|[InsertFileOptions](/.insertfileoptions)|[importChangeTrackingMode](/.insertfileoptions#word-javascript/api/word/-insertfileoptions-importchangetrackingmode-member)|Represents whether the change tracking mode status from the source document should be imported.|
||[importPageColor](/.insertfileoptions#word-javascript/api/word/-insertfileoptions-importpagecolor-member)|Represents whether the page color and other background information from the source document should be imported.|
||[importParagraphSpacing](/.insertfileoptions#word-javascript/api/word/-insertfileoptions-importparagraphspacing-member)|Represents whether the paragraph spacing from the source document should be imported.|
||[importStyles](/.insertfileoptions#word-javascript/api/word/-insertfileoptions-importstyles-member)|Represents whether the styles from the source document should be imported.|
||[importTheme](/.insertfileoptions#word-javascript/api/word/-insertfileoptions-importtheme-member)|Represents whether the theme from the source document should be imported.|
|[NoteItem](/.noteitem)|[body](/.noteitem#word-javascript/api/word/-noteitem-body-member)|Represents the body object of the note item.|
||[delete()](/.noteitem#word-javascript/api/word/-noteitem-delete-member(1))|Deletes the note item.|
||[getNext()](/.noteitem#word-javascript/api/word/-noteitem-getnext-member(1))|Gets the next note item of the same type.|
||[getNextOrNullObject()](/.noteitem#word-javascript/api/word/-noteitem-getnextornullobject-member(1))|Gets the next note item of the same type.|
||[reference](/.noteitem#word-javascript/api/word/-noteitem-reference-member)|Represents a footnote or endnote reference in the main document.|
||[type](/.noteitem#word-javascript/api/word/-noteitem-type-member)|Represents the note item type: footnote or endnote.|
|[NoteItemCollection](/.noteitemcollection)|[getFirst()](/.noteitemcollection#word-javascript/api/word/-noteitemcollection-getfirst-member(1))|Gets the first note item in this collection.|
||[getFirstOrNullObject()](/.noteitemcollection#word-javascript/api/word/-noteitemcollection-getfirstornullobject-member(1))|Gets the first note item in this collection.|
||[items](/.noteitemcollection#word-javascript/api/word/-noteitemcollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/.paragraph)|[endnotes](/.paragraph#word-javascript/api/word/-paragraph-endnotes-member)|Gets the collection of endnotes in the paragraph.|
||[footnotes](/.paragraph#word-javascript/api/word/-paragraph-footnotes-member)|Gets the collection of footnotes in the paragraph.|
||[getContentControls(options?: Word.ContentControlOptions)](/.paragraph#word-javascript/api/word/-paragraph-getcontentcontrols-member(1))|Gets the currently supported content controls in the paragraph.|
|[ParagraphFormat](/.paragraphformat)|[alignment](/.paragraphformat#word-javascript/api/word/-paragraphformat-alignment-member)|Specifies the alignment for the specified paragraphs.|
||[firstLineIndent](/.paragraphformat#word-javascript/api/word/-paragraphformat-firstlineindent-member)|Specifies the value (in points) for a first line or hanging indent.|
||[keepTogether](/.paragraphformat#word-javascript/api/word/-paragraphformat-keeptogether-member)|Specifies whether all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.|
||[keepWithNext](/.paragraphformat#word-javascript/api/word/-paragraphformat-keepwithnext-member)|Specifies whether the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.|
||[leftIndent](/.paragraphformat#word-javascript/api/word/-paragraphformat-leftindent-member)|Specifies the left indent.|
||[lineSpacing](/.paragraphformat#word-javascript/api/word/-paragraphformat-linespacing-member)|Specifies the line spacing (in points) for the specified paragraphs.|
||[lineUnitAfter](/.paragraphformat#word-javascript/api/word/-paragraphformat-lineunitafter-member)|Specifies the amount of spacing (in gridlines) after the specified paragraphs.|
||[lineUnitBefore](/.paragraphformat#word-javascript/api/word/-paragraphformat-lineunitbefore-member)|Specifies the amount of spacing (in gridlines) before the specified paragraphs.|
||[mirrorIndents](/.paragraphformat#word-javascript/api/word/-paragraphformat-mirrorindents-member)|Specifies whether left and right indents are the same width.|
||[outlineLevel](/.paragraphformat#word-javascript/api/word/-paragraphformat-outlinelevel-member)|Specifies the outline level for the specified paragraphs.|
||[rightIndent](/.paragraphformat#word-javascript/api/word/-paragraphformat-rightindent-member)|Specifies the right indent (in points) for the specified paragraphs.|
||[spaceAfter](/.paragraphformat#word-javascript/api/word/-paragraphformat-spaceafter-member)|Specifies the amount of spacing (in points) after the specified paragraph or text column.|
||[spaceBefore](/.paragraphformat#word-javascript/api/word/-paragraphformat-spacebefore-member)|Specifies the spacing (in points) before the specified paragraphs.|
||[widowControl](/.paragraphformat#word-javascript/api/word/-paragraphformat-widowcontrol-member)|Specifies whether the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Microsoft Word repaginates the document.|
|[Range](/.range)|[endnotes](/.range#word-javascript/api/word/-range-endnotes-member)|Gets the collection of endnotes in the range.|
||[footnotes](/.range#word-javascript/api/word/-range-footnotes-member)|Gets the collection of footnotes in the range.|
||[getContentControls(options?: Word.ContentControlOptions)](/.range#word-javascript/api/word/-range-getcontentcontrols-member(1))|Gets the currently supported content controls in the range.|
||[insertEndnote(insertText?: string)](/.range#word-javascript/api/word/-range-insertendnote-member(1))|Inserts an endnote.|
||[insertField(insertLocation: Word.InsertLocation \| "Replace" \| "Start" \| "End" \| "Before" \| "After", fieldType?: Word.FieldType, text?: string, removeFormatting?: boolean)](/.range#word-javascript/api/word/-range-insertfield-member(1))|Inserts a field at the specified location.|
||[insertFootnote(insertText?: string)](/.range#word-javascript/api/word/-range-insertfootnote-member(1))|Inserts a footnote.|
|[Style](/.style)|[baseStyle](/.style#word-javascript/api/word/-style-basestyle-member)|Specifies the name of an existing style to use as the base formatting of another style.|
||[builtIn](/.style#word-javascript/api/word/-style-builtin-member)|Gets whether the specified style is a built-in style.|
||[delete()](/.style#word-javascript/api/word/-style-delete-member(1))|Deletes the style.|
||[font](/.style#word-javascript/api/word/-style-font-member)|Gets a font object that represents the character formatting of the specified style.|
||[inUse](/.style#word-javascript/api/word/-style-inuse-member)|Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.|
||[linked](/.style#word-javascript/api/word/-style-linked-member)|Gets whether a style is a linked style that can be used for both paragraph and character formatting.|
||[nameLocal](/.style#word-javascript/api/word/-style-namelocal-member)|Gets the name of a style in the language of the user.|
||[nextParagraphStyle](/.style#word-javascript/api/word/-style-nextparagraphstyle-member)|Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.|
||[paragraphFormat](/.style#word-javascript/api/word/-style-paragraphformat-member)|Gets a ParagraphFormat object that represents the paragraph settings for the specified style.|
||[priority](/.style#word-javascript/api/word/-style-priority-member)|Specifies the priority.|
||[quickStyle](/.style#word-javascript/api/word/-style-quickstyle-member)|Specifies whether the style corresponds to an available quick style.|
||[type](/.style#word-javascript/api/word/-style-type-member)|Gets the style type.|
||[unhideWhenUsed](/.style#word-javascript/api/word/-style-unhidewhenused-member)|Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.|
||[visibility](/.style#word-javascript/api/word/-style-visibility-member)|Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.|
|[StyleCollection](/.stylecollection)|[getByName(name: string)](/.stylecollection#word-javascript/api/word/-stylecollection-getbyname-member(1))|Get the style object by its name.|
||[getByNameOrNullObject(name: string)](/.stylecollection#word-javascript/api/word/-stylecollection-getbynameornullobject-member(1))|If the corresponding style doesn't exist, then this method returns an object with its `isNullObject` property set to `true`.|
||[getCount()](/.stylecollection#word-javascript/api/word/-stylecollection-getcount-member(1))|Gets the number of the styles in the collection.|
||[getItem(index: number)](/.stylecollection#word-javascript/api/word/-stylecollection-getitem-member(1))|Gets a style object by its index in the collection.|
||[items](/.stylecollection#word-javascript/api/word/-stylecollection-items-member)|Gets the loaded child items in this collection.|
|[Table](/.table)|[endnotes](/.table#word-javascript/api/word/-table-endnotes-member)|Gets the collection of endnotes in the table.|
||[footnotes](/.table#word-javascript/api/word/-table-footnotes-member)|Gets the collection of footnotes in the table.|
|[TableRow](/.tablerow)|[endnotes](/.tablerow#word-javascript/api/word/-tablerow-endnotes-member)|Gets the collection of endnotes in the table row.|
||[footnotes](/.tablerow#word-javascript/api/word/-tablerow-footnotes-member)|Gets the collection of footnotes in the table row.|
