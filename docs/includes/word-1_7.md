| Class | Fields | Description |
|:---|:---|:---|
|[Annotation](/javascript/api/word/word.annotation)|[critiqueAnnotation](/javascript/api/word/word.annotation#word-word-annotation-critiqueannotation-member)|Gets the critique annotation object.|
||[delete()](/javascript/api/word/word.annotation#word-word-annotation-delete-member(1))|Deletes the annotation.|
||[id](/javascript/api/word/word.annotation#word-word-annotation-id-member)|Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.|
||[state](/javascript/api/word/word.annotation#word-word-annotation-state-member)|Gets the state of the annotation.|
|[AnnotationClickedEventArgs](/javascript/api/word/word.annotationclickedeventargs)|[id](/javascript/api/word/word.annotationclickedeventargs#word-word-annotationclickedeventargs-id-member)|Specifies the annotation ID for which the event was fired.|
|[AnnotationCollection](/javascript/api/word/word.annotationcollection)|[getFirst()](/javascript/api/word/word.annotationcollection#word-word-annotationcollection-getfirst-member(1))|Gets the first annotation in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.annotationcollection#word-word-annotationcollection-getfirstornullobject-member(1))|Gets the first annotation in this collection.|
||[items](/javascript/api/word/word.annotationcollection#word-word-annotationcollection-items-member)|Gets the loaded child items in this collection.|
|[AnnotationHoveredEventArgs](/javascript/api/word/word.annotationhoveredeventargs)|[id](/javascript/api/word/word.annotationhoveredeventargs#word-word-annotationhoveredeventargs-id-member)|Specifies the annotation ID for which the event was fired.|
|[AnnotationInsertedEventArgs](/javascript/api/word/word.annotationinsertedeventargs)|[ids](/javascript/api/word/word.annotationinsertedeventargs#word-word-annotationinsertedeventargs-ids-member)|Specifies the annotation IDs for which the event was fired.|
|[AnnotationRemovedEventArgs](/javascript/api/word/word.annotationremovedeventargs)|[ids](/javascript/api/word/word.annotationremovedeventargs#word-word-annotationremovedeventargs-ids-member)|Specifies the annotation IDs for which the event was fired.|
|[AnnotationSet](/javascript/api/word/word.annotationset)|[critiques](/javascript/api/word/word.annotationset#word-word-annotationset-critiques-member)|Critiques.|
|[CheckboxContentControl](/javascript/api/word/word.checkboxcontentcontrol)|[isChecked](/javascript/api/word/word.checkboxcontentcontrol#word-word-checkboxcontentcontrol-ischecked-member)|Specifies the current state of the checkbox.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[checkboxContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-checkboxcontentcontrol-member)|Gets the data of the content control when its type is 'CheckBox'.|
|[Critique](/javascript/api/word/word.critique)|[colorScheme](/javascript/api/word/word.critique#word-word-critique-colorscheme-member)|Specifies the color scheme of the critique.|
||[length](/javascript/api/word/word.critique#word-word-critique-length-member)|Specifies the length of the critique inside paragraph.|
||[start](/javascript/api/word/word.critique#word-word-critique-start-member)|Specifies the start index of the critique inside paragraph.|
|[CritiqueAnnotation](/javascript/api/word/word.critiqueannotation)|[accept()](/javascript/api/word/word.critiqueannotation#word-word-critiqueannotation-accept-member(1))|Accepts the critique.|
||[critique](/javascript/api/word/word.critiqueannotation#word-word-critiqueannotation-critique-member)|Gets the critique that was passed when the annotation was inserted.|
||[range](/javascript/api/word/word.critiqueannotation#word-word-critiqueannotation-range-member)|Gets the range of text that is annotated.|
||[reject()](/javascript/api/word/word.critiqueannotation#word-word-critiqueannotation-reject-member(1))|Rejects the critique.|
|[Document](/javascript/api/word/word.document)|[getAnnotationById(id: string)](/javascript/api/word/word.document#word-word-document-getannotationbyid-member(1))|Gets the annotation by ID.|
||[onAnnotationClicked](/javascript/api/word/word.document#word-word-document-onannotationclicked-member)|Occurs when the user clicks an annotation (or selects it using **Alt+Down**).|
||[onAnnotationHovered](/javascript/api/word/word.document#word-word-document-onannotationhovered-member)|Occurs when the user hovers the cursor over an annotation.|
||[onAnnotationInserted](/javascript/api/word/word.document#word-word-document-onannotationinserted-member)|Occurs when the user adds one or more annotations.|
||[onAnnotationRemoved](/javascript/api/word/word.document#word-word-document-onannotationremoved-member)|Occurs when the user deletes one or more annotations.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.document#word-word-document-search-member(1))|Performs a search with the specified search options on the scope of the whole document.|
|[GetTextOptions](/javascript/api/word/word.gettextoptions)|[includeHiddenText](/javascript/api/word/word.gettextoptions#word-word-gettextoptions-includehiddentext-member)|Specifies a value that indicates whether to include hidden text in the result of the GetText method.|
||[includeTextMarkedAsDeleted](/javascript/api/word/word.gettextoptions#word-word-gettextoptions-includetextmarkedasdeleted-member)|Specifies a value that indicates whether to include text marked as deleted in the result of the GetText method.|
|[InsertFileOptions](/javascript/api/word/word.insertfileoptions)|[importDifferentOddEvenPages](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importdifferentoddevenpages-member)|Represents whether to import the Different Odd and Even Pages setting for the header and footer from the source document.|
|[Paragraph](/javascript/api/word/word.paragraph)|[getAnnotations()](/javascript/api/word/word.paragraph#word-word-paragraph-getannotations-member(1))|Gets annotations set on this Paragraph object.|
||[getText(options?: Word.GetTextOptions \| { IncludeHiddenText?: boolean IncludeTextMarkedAsDeleted?: boolean })](/javascript/api/word/word.paragraph#word-word-paragraph-gettext-member(1))|Returns the text of the paragraph.|
||[insertAnnotations(annotations: Word.AnnotationSet)](/javascript/api/word/word.paragraph#word-word-paragraph-insertannotations-member(1))|Inserts annotations on this Paragraph object.|
