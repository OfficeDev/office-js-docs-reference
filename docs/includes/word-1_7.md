| Class | Fields | Description |
|:---|:---|:---|
|[Annotation](/.annotation)|[critiqueAnnotation](/.annotation#word-javascript/api/word/-annotation-critiqueannotation-member)|Gets the critique annotation object.|
||[delete()](/.annotation#word-javascript/api/word/-annotation-delete-member(1))|Deletes the annotation.|
||[id](/.annotation#word-javascript/api/word/-annotation-id-member)|Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.|
||[state](/.annotation#word-javascript/api/word/-annotation-state-member)|Gets the state of the annotation.|
|[AnnotationClickedEventArgs](/.annotationclickedeventargs)|[id](/.annotationclickedeventargs#word-javascript/api/word/-annotationclickedeventargs-id-member)|Specifies the annotation ID for which the event was fired.|
|[AnnotationCollection](/.annotationcollection)|[getFirst()](/.annotationcollection#word-javascript/api/word/-annotationcollection-getfirst-member(1))|Gets the first annotation in this collection.|
||[getFirstOrNullObject()](/.annotationcollection#word-javascript/api/word/-annotationcollection-getfirstornullobject-member(1))|Gets the first annotation in this collection.|
||[items](/.annotationcollection#word-javascript/api/word/-annotationcollection-items-member)|Gets the loaded child items in this collection.|
|[AnnotationHoveredEventArgs](/.annotationhoveredeventargs)|[id](/.annotationhoveredeventargs#word-javascript/api/word/-annotationhoveredeventargs-id-member)|Specifies the annotation ID for which the event was fired.|
|[AnnotationInsertedEventArgs](/.annotationinsertedeventargs)|[ids](/.annotationinsertedeventargs#word-javascript/api/word/-annotationinsertedeventargs-ids-member)|Specifies the annotation IDs for which the event was fired.|
|[AnnotationRemovedEventArgs](/.annotationremovedeventargs)|[ids](/.annotationremovedeventargs#word-javascript/api/word/-annotationremovedeventargs-ids-member)|Specifies the annotation IDs for which the event was fired.|
|[AnnotationSet](/.annotationset)|[critiques](/.annotationset#word-javascript/api/word/-annotationset-critiques-member)|Critiques.|
|[CheckboxContentControl](/.checkboxcontentcontrol)|[isChecked](/.checkboxcontentcontrol#word-javascript/api/word/-checkboxcontentcontrol-ischecked-member)|Specifies the current state of the checkbox.|
|[ContentControl](/.contentcontrol)|[checkboxContentControl](/.contentcontrol#word-javascript/api/word/-contentcontrol-checkboxcontentcontrol-member)|Gets the data of the content control when its type is `CheckBox`.|
|[Critique](/.critique)|[colorScheme](/.critique#word-javascript/api/word/-critique-colorscheme-member)|Specifies the color scheme of the critique.|
||[length](/.critique#word-javascript/api/word/-critique-length-member)|Specifies the length of the critique inside paragraph.|
||[start](/.critique#word-javascript/api/word/-critique-start-member)|Specifies the start index of the critique inside paragraph.|
|[CritiqueAnnotation](/.critiqueannotation)|[accept()](/.critiqueannotation#word-javascript/api/word/-critiqueannotation-accept-member(1))|Accepts the critique.|
||[critique](/.critiqueannotation#word-javascript/api/word/-critiqueannotation-critique-member)|Gets the critique that was passed when the annotation was inserted.|
||[range](/.critiqueannotation#word-javascript/api/word/-critiqueannotation-range-member)|Gets the range of text that is annotated.|
||[reject()](/.critiqueannotation#word-javascript/api/word/-critiqueannotation-reject-member(1))|Rejects the critique.|
|[Document](/.document)|[getAnnotationById(id: string)](/.document#word-javascript/api/word/-document-getannotationbyid-member(1))|Gets the annotation by ID.|
||[onAnnotationClicked](/.document#word-javascript/api/word/-document-onannotationclicked-member)|Occurs when the user clicks an annotation (or selects it using **Alt+Down**).|
||[onAnnotationHovered](/.document#word-javascript/api/word/-document-onannotationhovered-member)|Occurs when the user hovers the cursor over an annotation.|
||[onAnnotationInserted](/.document#word-javascript/api/word/-document-onannotationinserted-member)|Occurs when the user adds one or more annotations.|
||[onAnnotationRemoved](/.document#word-javascript/api/word/-document-onannotationremoved-member)|Occurs when the user deletes one or more annotations.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/.document#word-javascript/api/word/-document-search-member(1))|Performs a search with the specified search options on the scope of the whole document.|
|[GetTextOptions](/.gettextoptions)|[includeHiddenText](/.gettextoptions#word-javascript/api/word/-gettextoptions-includehiddentext-member)|Specifies a value that indicates whether to include hidden text in the result of the GetText method.|
||[includeTextMarkedAsDeleted](/.gettextoptions#word-javascript/api/word/-gettextoptions-includetextmarkedasdeleted-member)|Specifies a value that indicates whether to include text marked as deleted in the result of the GetText method.|
|[InsertFileOptions](/.insertfileoptions)|[importDifferentOddEvenPages](/.insertfileoptions#word-javascript/api/word/-insertfileoptions-importdifferentoddevenpages-member)|Represents whether to import the Different Odd and Even Pages setting for the header and footer from the source document.|
|[Paragraph](/.paragraph)|[getAnnotations()](/.paragraph#word-javascript/api/word/-paragraph-getannotations-member(1))|Gets annotations set on this Paragraph object.|
||[getText(options?: Word.GetTextOptions \| { IncludeHiddenText?: boolean IncludeTextMarkedAsDeleted?: boolean })](/.paragraph#word-javascript/api/word/-paragraph-gettext-member(1))|Returns the text of the paragraph.|
||[insertAnnotations(annotations: Word.AnnotationSet)](/.paragraph#word-javascript/api/word/-paragraph-insertannotations-member(1))|Inserts annotations on this Paragraph object.|
