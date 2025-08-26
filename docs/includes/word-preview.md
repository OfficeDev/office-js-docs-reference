| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[bibliography](/javascript/api/word/word.application#word-word-application-bibliography-member)|Returns a `Bibliography` object that represents the bibliography reference sources stored in Microsoft Word.|
||[checkLanguage](/javascript/api/word/word.application#word-word-application-checklanguage-member)|Specifies if Microsoft Word automatically detects the language you are using as you type.|
||[language](/javascript/api/word/word.application#word-word-application-language-member)|Gets a `LanguageId` value that represents the language selected for the Microsoft Word user interface.|
||[templates](/javascript/api/word/word.application#word-word-application-templates-member)|Returns a `TemplateCollection` object that represents all the available templates: global templates and those attached to open documents.|
|[Bibliography](/javascript/api/word/word.bibliography)|[bibliographyStyle](/javascript/api/word/word.bibliography#word-word-bibliography-bibliographystyle-member)|Specifies the name of the active style to use for the bibliography.|
||[generateUniqueTag()](/javascript/api/word/word.bibliography#word-word-bibliography-generateuniquetag-member(1))|Generates a unique identification tag for a bibliography source and returns a string that represents the tag.|
||[sources](/javascript/api/word/word.bibliography#word-word-bibliography-sources-member)|Returns a `SourceCollection` object that represents all the sources contained in the bibliography.|
|[Body](/javascript/api/word/word.body)|[onCommentAdded](/javascript/api/word/word.body#word-word-body-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.body#word-word-body-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.body#word-word-body-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.body#word-word-body-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.body#word-word-body-oncommentselected-member)|Occurs when a comment is selected.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Gets the type of the body.|
|[Bookmark](/javascript/api/word/word.bookmark)|[copyTo(name: string)](/javascript/api/word/word.bookmark#word-word-bookmark-copyto-member(1))|Copies this bookmark to the new bookmark specified in the `name` argument and returns a `Bookmark` object.|
||[delete()](/javascript/api/word/word.bookmark#word-word-bookmark-delete-member(1))|Deletes the bookmark.|
||[end](/javascript/api/word/word.bookmark#word-word-bookmark-end-member)|Specifies the ending character position of the bookmark.|
||[isColumn](/javascript/api/word/word.bookmark#word-word-bookmark-iscolumn-member)|Returns `true` if the bookmark is a table column.|
||[isEmpty](/javascript/api/word/word.bookmark#word-word-bookmark-isempty-member)|Returns `true` if the bookmark is empty.|
||[name](/javascript/api/word/word.bookmark#word-word-bookmark-name-member)|Returns the name of the `Bookmark` object.|
||[range](/javascript/api/word/word.bookmark#word-word-bookmark-range-member)|Returns a `Range` object that represents the portion of the document that's contained in the `Bookmark` object.|
||[select()](/javascript/api/word/word.bookmark#word-word-bookmark-select-member(1))|Selects the bookmark.|
||[start](/javascript/api/word/word.bookmark#word-word-bookmark-start-member)|Specifies the starting character position of the bookmark.|
||[storyType](/javascript/api/word/word.bookmark#word-word-bookmark-storytype-member)|Returns the story type for the bookmark.|
|[BookmarkCollection](/javascript/api/word/word.bookmarkcollection)|[exists(name: string)](/javascript/api/word/word.bookmarkcollection#word-word-bookmarkcollection-exists-member(1))|Determines whether the specified bookmark exists.|
||[getItem(index: number)](/javascript/api/word/word.bookmarkcollection#word-word-bookmarkcollection-getitem-member(1))|Gets a `Bookmark` object by its index in the collection.|
||[items](/javascript/api/word/word.bookmarkcollection#word-word-bookmarkcollection-items-member)|Gets the loaded child items in this collection.|
|[BorderUniversal](/javascript/api/word/word.borderuniversal)|[artStyle](/javascript/api/word/word.borderuniversal#word-word-borderuniversal-artstyle-member)|Specifies the graphical page-border design for the document.|
||[artWidth](/javascript/api/word/word.borderuniversal#word-word-borderuniversal-artwidth-member)|Specifies the width (in points) of the graphical page border specified in the `artStyle` property.|
||[color](/javascript/api/word/word.borderuniversal#word-word-borderuniversal-color-member)|Specifies the color for the `BorderUniversal` object.|
||[colorIndex](/javascript/api/word/word.borderuniversal#word-word-borderuniversal-colorindex-member)|Specifies the color for the `BorderUniversal` or Word.Font object.|
||[inside](/javascript/api/word/word.borderuniversal#word-word-borderuniversal-inside-member)|Returns `true` if an inside border can be applied to the specified object.|
||[isVisible](/javascript/api/word/word.borderuniversal#word-word-borderuniversal-isvisible-member)|Specifies whether the border is visible.|
||[lineStyle](/javascript/api/word/word.borderuniversal#word-word-borderuniversal-linestyle-member)|Specifies the line style of the border.|
||[lineWidth](/javascript/api/word/word.borderuniversal#word-word-borderuniversal-linewidth-member)|Specifies the line width of an object's border.|
|[BorderUniversalCollection](/javascript/api/word/word.borderuniversalcollection)|[applyPageBordersToAllSections()](/javascript/api/word/word.borderuniversalcollection#word-word-borderuniversalcollection-applypageborderstoallsections-member(1))|Applies the specified page-border formatting to all sections in the document.|
||[getItem(index: number)](/javascript/api/word/word.borderuniversalcollection#word-word-borderuniversalcollection-getitem-member(1))|Gets a `Border` object by its index in the collection.|
||[items](/javascript/api/word/word.borderuniversalcollection#word-word-borderuniversalcollection-items-member)|Gets the loaded child items in this collection.|
|[Break](/javascript/api/word/word.break)|[pageIndex](/javascript/api/word/word.break#word-word-break-pageindex-member)|Returns the page number on which the break occurs.|
||[range](/javascript/api/word/word.break#word-word-break-range-member)|Returns a `Range` object that represents the portion of the document that's contained in the break.|
|[BreakCollection](/javascript/api/word/word.breakcollection)|[items](/javascript/api/word/word.breakcollection#word-word-breakcollection-items-member)|Gets the loaded child items in this collection.|
|[BuildingBlock](/javascript/api/word/word.buildingblock)|[category](/javascript/api/word/word.buildingblock#word-word-buildingblock-category-member)|Returns a `BuildingBlockCategory` object that represents the category for the building block.|
||[delete()](/javascript/api/word/word.buildingblock#word-word-buildingblock-delete-member(1))|Deletes the building block.|
||[description](/javascript/api/word/word.buildingblock#word-word-buildingblock-description-member)|Specifies the description for the building block.|
||[id](/javascript/api/word/word.buildingblock#word-word-buildingblock-id-member)|Returns the internal identification number for the building block.|
||[index](/javascript/api/word/word.buildingblock#word-word-buildingblock-index-member)|Returns the position of this building block in a collection.|
||[insert(range: Word.Range, richText: boolean)](/javascript/api/word/word.buildingblock#word-word-buildingblock-insert-member(1))|Inserts the value of the building block into the document and returns a `Range` object that represents the contents of the building block within the document.|
||[insertType](/javascript/api/word/word.buildingblock#word-word-buildingblock-inserttype-member)|Specifies a `DocPartInsertType` value that represents how to insert the contents of the building block into the document.|
||[name](/javascript/api/word/word.buildingblock#word-word-buildingblock-name-member)|Specifies the name of the building block.|
||[type](/javascript/api/word/word.buildingblock#word-word-buildingblock-type-member)|Returns a `BuildingBlockTypeItem` object that represents the type for the building block.|
||[value](/javascript/api/word/word.buildingblock#word-word-buildingblock-value-member)|Specifies the contents of the building block.|
|[BuildingBlockCategory](/javascript/api/word/word.buildingblockcategory)|[buildingBlocks](/javascript/api/word/word.buildingblockcategory#word-word-buildingblockcategory-buildingblocks-member)|Returns a `BuildingBlockCollection` object that represents the building blocks for the category.|
||[index](/javascript/api/word/word.buildingblockcategory#word-word-buildingblockcategory-index-member)|Returns the position of the `BuildingBlockCategory` object in a collection.|
||[name](/javascript/api/word/word.buildingblockcategory#word-word-buildingblockcategory-name-member)|Returns the name of the `BuildingBlockCategory` object.|
||[type](/javascript/api/word/word.buildingblockcategory#word-word-buildingblockcategory-type-member)|Returns a `BuildingBlockTypeItem` object that represents the type of building block for the building block category.|
|[BuildingBlockCategoryCollection](/javascript/api/word/word.buildingblockcategorycollection)|[getCount()](/javascript/api/word/word.buildingblockcategorycollection#word-word-buildingblockcategorycollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/javascript/api/word/word.buildingblockcategorycollection#word-word-buildingblockcategorycollection-getitemat-member(1))|Returns a `BuildingBlockCategory` object that represents the specified item in the collection.|
|[BuildingBlockCollection](/javascript/api/word/word.buildingblockcollection)|[add(name: string, range: Word.Range, description: string, insertType: Word.DocPartInsertType)](/javascript/api/word/word.buildingblockcollection#word-word-buildingblockcollection-add-member(1))|Creates a new building block and returns a `BuildingBlock` object.|
||[getCount()](/javascript/api/word/word.buildingblockcollection#word-word-buildingblockcollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/javascript/api/word/word.buildingblockcollection#word-word-buildingblockcollection-getitemat-member(1))|Returns a `BuildingBlock` object that represents the specified item in the collection.|
|[BuildingBlockEntryCollection](/javascript/api/word/word.buildingblockentrycollection)|[add(name: string, type: Word.BuildingBlockType, category: string, range: Word.Range, description: string, insertType: Word.DocPartInsertType)](/javascript/api/word/word.buildingblockentrycollection#word-word-buildingblockentrycollection-add-member(1))|Creates a new building block entry in a template and returns a `BuildingBlock` object that represents the new building block entry.|
||[getCount()](/javascript/api/word/word.buildingblockentrycollection#word-word-buildingblockentrycollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/javascript/api/word/word.buildingblockentrycollection#word-word-buildingblockentrycollection-getitemat-member(1))|Returns a `BuildingBlock` object that represents the specified item in the collection.|
|[BuildingBlockGalleryContentControl](/javascript/api/word/word.buildingblockgallerycontentcontrol)|[appearance](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[buildingBlockCategory](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-buildingblockcategory-member)|Specifies the category for the building block content control.|
||[buildingBlockType](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-buildingblocktype-member)|Specifies a `BuildingBlockType` value that represents the type of building block for the building block content control.|
||[color](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-color-member)|Specifies the red-green-blue (RGB) value of the color of the content control.|
||[copy()](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-copy-member(1))|Copies the content control from the active document to the Clipboard.|
||[cut()](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-cut-member(1))|Removes the content control from the active document and moves the content control to the Clipboard.|
||[delete(deleteContents?: boolean)](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-delete-member(1))|Deletes the content control and optionally its contents.|
||[id](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-id-member)|Gets the identification for the content control.|
||[isTemporary](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-istemporary-member)|Specifies whether to remove the content control from the active document when the user edits the contents of the control.|
||[level](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-level-member)|Gets the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.|
||[lockContentControl](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-lockcontentcontrol-member)|Specifies if the content control is locked (can't be deleted).|
||[lockContents](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-lockcontents-member)|Specifies if the contents of the content control are locked (not editable).|
||[placeholderText](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-placeholdertext-member)|Returns a `BuildingBlock` object that represents the placeholder text for the content control.|
||[range](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-range-member)|Returns a `Range` object that represents the contents of the content control in the active document.|
||[setPlaceholderText(options?: Word.ContentControlPlaceholderOptions)](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-setplaceholdertext-member(1))|Sets the placeholder text that displays in the content control until a user enters their own text.|
||[showingPlaceholderText](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-showingplaceholdertext-member)|Gets if the placeholder text for the content control is being displayed.|
||[tag](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-tag-member)|Specifies a tag to identify the content control.|
||[title](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-title-member)|Specifies the title for the content control.|
||[xmlMapping](/javascript/api/word/word.buildingblockgallerycontentcontrol#word-word-buildingblockgallerycontentcontrol-xmlmapping-member)|Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[BuildingBlockTypeItem](/javascript/api/word/word.buildingblocktypeitem)|[categories](/javascript/api/word/word.buildingblocktypeitem#word-word-buildingblocktypeitem-categories-member)|Returns a `BuildingBlockCategoryCollection` object that represents the categories for a building block type.|
||[index](/javascript/api/word/word.buildingblocktypeitem#word-word-buildingblocktypeitem-index-member)|Returns the position of an item in a collection.|
||[name](/javascript/api/word/word.buildingblocktypeitem#word-word-buildingblocktypeitem-name-member)|Returns the localized name of a building block type.|
|[BuildingBlockTypeItemCollection](/javascript/api/word/word.buildingblocktypeitemcollection)|[getByType(type: Word.BuildingBlockType)](/javascript/api/word/word.buildingblocktypeitemcollection#word-word-buildingblocktypeitemcollection-getbytype-member(1))|Gets a `BuildingBlockTypeItem` object by its type in the collection.|
||[getCount()](/javascript/api/word/word.buildingblocktypeitemcollection#word-word-buildingblocktypeitemcollection-getcount-member(1))|Returns the number of items in the collection.|
|[ColorFormat](/javascript/api/word/word.colorformat)|[brightness](/javascript/api/word/word.colorformat#word-word-colorformat-brightness-member)|Specifies the brightness of a specified shape color.|
||[objectThemeColor](/javascript/api/word/word.colorformat#word-word-colorformat-objectthemecolor-member)|Specifies the theme color for a color format.|
||[rgb](/javascript/api/word/word.colorformat#word-word-colorformat-rgb-member)|Specifies the red-green-blue (RGB) value of the specified color.|
||[tintAndShade](/javascript/api/word/word.colorformat#word-word-colorformat-tintandshade-member)|Specifies the lightening or darkening of a specified shape's color.|
||[type](/javascript/api/word/word.colorformat#word-word-colorformat-type-member)|Returns the shape color type.|
|[CommentDetail](/javascript/api/word/word.commentdetail)|[id](/javascript/api/word/word.commentdetail#word-word-commentdetail-id-member)|Represents the ID of this comment.|
||[replyIds](/javascript/api/word/word.commentdetail#word-word-commentdetail-replyids-member)|Represents the IDs of the replies to this comment.|
|[CommentEventArgs](/javascript/api/word/word.commenteventargs)|[changeType](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-changetype-member)|Represents how the comment changed event is triggered.|
||[commentDetails](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-commentdetails-member)|Gets the CommentDetail array which contains the IDs and reply IDs of the involved comments.|
||[source](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-type-member)|The event type.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[buildingBlockGalleryContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-buildingblockgallerycontentcontrol-member)|Gets the building block gallery-related data if the content control's Word.ContentControlType is `BuildingBlockGallery`.|
||[datePickerContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-datepickercontentcontrol-member)|Gets the date picker-related data if the content control's Word.ContentControlType is `DatePicker`.|
||[groupContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-groupcontentcontrol-member)|Gets the group-related data if the content control's Word.ContentControlType is `Group`.|
||[onCommentAdded](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentselected-member)|Occurs when a comment is selected.|
||[pictureContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-picturecontentcontrol-member)|Gets the picture-related data if the content control's Word.ContentControlType is `Picture`.|
||[repeatingSectionContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-repeatingsectioncontentcontrol-member)|Gets the repeating section-related data if the content control's Word.ContentControlType is `RepeatingSection`.|
||[resetState()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-resetstate-member(1))|Resets the state of the content control.|
||[setState(contentControlState: Word.ContentControlState)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-setstate-member(1))|Sets the state of the content control.|
||[xmlMapping](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-xmlmapping-member)|Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs)|[eventType](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-eventtype-member)|The event type.|
|[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs)|[eventType](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
|[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs)|[eventType](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-eventtype-member)|The event type.|
|[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs)|[eventType](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-eventtype-member)|The event type.|
|[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs)|[eventType](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-eventtype-member)|The event type.|
|[ContentControlPlaceholderOptions](/javascript/api/word/word.contentcontrolplaceholderoptions)|[buildingBlock](/javascript/api/word/word.contentcontrolplaceholderoptions#word-word-contentcontrolplaceholderoptions-buildingblock-member)|If provided, specifies the `BuildingBlock` object to use as placeholder.|
||[range](/javascript/api/word/word.contentcontrolplaceholderoptions#word-word-contentcontrolplaceholderoptions-range-member)|If provided, specifies the `Range` object to use as placeholder.|
||[text](/javascript/api/word/word.contentcontrolplaceholderoptions#word-word-contentcontrolplaceholderoptions-text-member)|If provided, specifies the text to use as placeholder.|
|[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs)|[eventType](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
|[CustomXmlAddNodeOptions](/javascript/api/word/word.customxmladdnodeoptions)|[name](/javascript/api/word/word.customxmladdnodeoptions#word-word-customxmladdnodeoptions-name-member)|If provided, specifies the base name of the element to be added.|
||[namespaceUri](/javascript/api/word/word.customxmladdnodeoptions#word-word-customxmladdnodeoptions-namespaceuri-member)|If provided, specifies the namespace of the element to be appended.|
||[nextSibling](/javascript/api/word/word.customxmladdnodeoptions#word-word-customxmladdnodeoptions-nextsibling-member)|If provided, specifies the node which should become the next sibling of the new node.|
||[nodeType](/javascript/api/word/word.customxmladdnodeoptions#word-word-customxmladdnodeoptions-nodetype-member)|If provided, specifies the type of node to add.|
||[nodeValue](/javascript/api/word/word.customxmladdnodeoptions#word-word-customxmladdnodeoptions-nodevalue-member)|If provided, specifies the value of the added node for those nodes that allow text.|
|[CustomXmlAddSchemaOptions](/javascript/api/word/word.customxmladdschemaoptions)|[alias](/javascript/api/word/word.customxmladdschemaoptions#word-word-customxmladdschemaoptions-alias-member)|If provided, specifies the alias of the schema to be added to the collection.|
||[fileName](/javascript/api/word/word.customxmladdschemaoptions#word-word-customxmladdschemaoptions-filename-member)|If provided, specifies the location of the schema on a disk.|
||[installForAllUsers](/javascript/api/word/word.customxmladdschemaoptions#word-word-customxmladdschemaoptions-installforallusers-member)|If provided, specifies whether, in the case where the schema is being added to the Schema Library, the Schema Library keys should be written to the registry (`HKEY_LOCAL_MACHINE` for all users or `HKEY_CURRENT_USER` for just the current user).|
||[namespaceUri](/javascript/api/word/word.customxmladdschemaoptions#word-word-customxmladdschemaoptions-namespaceuri-member)|If provided, specifies the namespace of the schema to be added to the collection.|
|[CustomXmlAddValidationErrorOptions](/javascript/api/word/word.customxmladdvalidationerroroptions)|[clearedOnUpdate](/javascript/api/word/word.customxmladdvalidationerroroptions#word-word-customxmladdvalidationerroroptions-clearedonupdate-member)|If provided, specifies whether the error is to be cleared from the Word.CustomXmlValidationErrorCollection when the XML is corrected and updated.|
||[errorText](/javascript/api/word/word.customxmladdvalidationerroroptions#word-word-customxmladdvalidationerroroptions-errortext-member)|If provided, specifies the descriptive error text.|
|[CustomXmlAppendChildNodeOptions](/javascript/api/word/word.customxmlappendchildnodeoptions)|[name](/javascript/api/word/word.customxmlappendchildnodeoptions#word-word-customxmlappendchildnodeoptions-name-member)|If provided, specifies the base name of the element to be appended.|
||[namespaceUri](/javascript/api/word/word.customxmlappendchildnodeoptions#word-word-customxmlappendchildnodeoptions-namespaceuri-member)|If provided, specifies the namespace of the element to be appended.|
||[nodeType](/javascript/api/word/word.customxmlappendchildnodeoptions#word-word-customxmlappendchildnodeoptions-nodetype-member)|If provided, specifies the type of node to append.|
||[nodeValue](/javascript/api/word/word.customxmlappendchildnodeoptions#word-word-customxmlappendchildnodeoptions-nodevalue-member)|If provided, specifies the value of the appended node for those nodes that allow text.|
|[CustomXmlInsertNodeBeforeOptions](/javascript/api/word/word.customxmlinsertnodebeforeoptions)|[name](/javascript/api/word/word.customxmlinsertnodebeforeoptions#word-word-customxmlinsertnodebeforeoptions-name-member)|If provided, specifies the base name of the element to be inserted.|
||[namespaceUri](/javascript/api/word/word.customxmlinsertnodebeforeoptions#word-word-customxmlinsertnodebeforeoptions-namespaceuri-member)|If provided, specifies the namespace of the element to be inserted.|
||[nextSibling](/javascript/api/word/word.customxmlinsertnodebeforeoptions#word-word-customxmlinsertnodebeforeoptions-nextsibling-member)|If provided, specifies the context node.|
||[nodeType](/javascript/api/word/word.customxmlinsertnodebeforeoptions#word-word-customxmlinsertnodebeforeoptions-nodetype-member)|If provided, specifies the type of node to append.|
||[nodeValue](/javascript/api/word/word.customxmlinsertnodebeforeoptions#word-word-customxmlinsertnodebeforeoptions-nodevalue-member)|If provided, specifies the value of the inserted node for those nodes that allow text.|
|[CustomXmlInsertSubtreeBeforeOptions](/javascript/api/word/word.customxmlinsertsubtreebeforeoptions)|[nextSibling](/javascript/api/word/word.customxmlinsertsubtreebeforeoptions#word-word-customxmlinsertsubtreebeforeoptions-nextsibling-member)|If provided, specifies the context node.|
|[CustomXmlNode](/javascript/api/word/word.customxmlnode)|[appendChildNode(options?: Word.CustomXmlAppendChildNodeOptions)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-appendchildnode-member(1))|Appends a single node as the last child under the context element node in the tree.|
||[appendChildSubtree(xml: string)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-appendchildsubtree-member(1))|Adds a subtree as the last child under the context element node in the tree.|
||[attributes](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-attributes-member)|Gets a `CustomXmlNodeCollection` object representing the attributes of the current element in the current node.|
||[baseName](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-basename-member)|Gets the base name of the node without the namespace prefix, if one exists.|
||[childNodes](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-childnodes-member)|Gets a `CustomXmlNodeCollection` object containing all of the child elements of the current node.|
||[delete()](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-delete-member(1))|Deletes the current node from the tree (including all of its children, if any exist).|
||[firstChild](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-firstchild-member)|Gets a `CustomXmlNode` object corresponding to the first child element of the current node.|
||[hasChildNodes()](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-haschildnodes-member(1))|Specifies if the current element node has child element nodes.|
||[insertNodeBefore(options?: Word.CustomXmlInsertNodeBeforeOptions)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-insertnodebefore-member(1))|Inserts a new node just before the context node in the tree.|
||[insertSubtreeBefore(xml: string, options?: Word.CustomXmlInsertSubtreeBeforeOptions)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-insertsubtreebefore-member(1))|Inserts the specified subtree into the location just before the context node.|
||[lastChild](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-lastchild-member)|Gets a `CustomXmlNode` object corresponding to the last child element of the current node.|
||[namespaceUri](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-namespaceuri-member)|Gets the unique address identifier for the namespace of the node.|
||[nextSibling](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-nextsibling-member)|Gets the next sibling node (element, comment, or processing instruction) of the current node.|
||[nodeType](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-nodetype-member)|Gets the type of the current node.|
||[nodeValue](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-nodevalue-member)|Specifies the value of the current node.|
||[ownerPart](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-ownerpart-member)|Gets the object representing the part associated with this node.|
||[parentNode](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-parentnode-member)|Gets the parent element node of the current node.|
||[previousSibling](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-previoussibling-member)|Gets the previous sibling node (element, comment, or processing instruction) of the current node.|
||[removeChild(child: Word.CustomXmlNode)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-removechild-member(1))|Removes the specified child node from the tree.|
||[replaceChildNode(oldNode: Word.CustomXmlNode, options?: Word.CustomXmlReplaceChildNodeOptions)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-replacechildnode-member(1))|Removes the specified child node and replaces it with a different node in the same location.|
||[replaceChildSubtree(xml: string, oldNode: Word.CustomXmlNode)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-replacechildsubtree-member(1))|Removes the specified node and replaces it with a different subtree in the same location.|
||[selectNodes(xPath: string)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-selectnodes-member(1))|Selects a collection of nodes matching an XPath expression.|
||[selectSingleNode(xPath: string)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-selectsinglenode-member(1))|Selects a single node from a collection matching an XPath expression.|
||[text](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-text-member)|Specifies the text for the current node.|
||[xml](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-xml-member)|Gets the XML representation of the current node and its children.|
||[xpath](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-xpath-member)|Gets a string with the canonicalized XPath for the current node.|
|[CustomXmlNodeCollection](/javascript/api/word/word.customxmlnodecollection)|[getCount()](/javascript/api/word/word.customxmlnodecollection#word-word-customxmlnodecollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/javascript/api/word/word.customxmlnodecollection#word-word-customxmlnodecollection-getitem-member(1))|Returns a `CustomXmlNode` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.customxmlnodecollection#word-word-customxmlnodecollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[addNode(parent: Word.CustomXmlNode, options?: Word.CustomXmlAddNodeOptions)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-addnode-member(1))|Adds a node to the XML tree.|
||[builtIn](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-builtin-member)|Gets a value that indicates whether the `CustomXmlPart` is built-in.|
||[documentElement](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-documentelement-member)|Gets the root element of a bound region of data in the document.|
||[errors](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-errors-member)|Gets a `CustomXmlValidationErrorCollection` object that provides access to any XML validation errors.|
||[loadXml(xml: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-loadxml-member(1))|Populates the `CustomXmlPart` object from an XML string.|
||[namespaceManager](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-namespacemanager-member)|Gets the set of namespace prefix mappings used against the current `CustomXmlPart` object.|
||[schemaCollection](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-schemacollection-member)|Specifies a `CustomXmlSchemaCollection` object representing the set of schemas attached to a bound region of data in the document.|
||[selectNodes(xPath: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-selectnodes-member(1))|Selects a collection of nodes from a custom XML part.|
||[selectSingleNode(xPath: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-selectsinglenode-member(1))|Selects a single node within a custom XML part matching an XPath expression.|
||[xml](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-xml-member)|Gets the XML representation of the current `CustomXmlPart` object.|
|[CustomXmlPrefixMapping](/javascript/api/word/word.customxmlprefixmapping)|[namespaceUri](/javascript/api/word/word.customxmlprefixmapping#word-word-customxmlprefixmapping-namespaceuri-member)|Gets the unique address identifier for the namespace of the `CustomXmlPrefixMapping` object.|
||[prefix](/javascript/api/word/word.customxmlprefixmapping#word-word-customxmlprefixmapping-prefix-member)|Gets the prefix for the `CustomXmlPrefixMapping` object.|
|[CustomXmlPrefixMappingCollection](/javascript/api/word/word.customxmlprefixmappingcollection)|[addNamespace(prefix: string, namespaceUri: string)](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-addnamespace-member(1))|Adds a custom namespace/prefix mapping to use when querying an item.|
||[getCount()](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-getitem-member(1))|Returns a `CustomXmlPrefixMapping` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-items-member)|Gets the loaded child items in this collection.|
||[lookupNamespace(prefix: string)](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-lookupnamespace-member(1))|Gets the namespace corresponding to the specified prefix.|
||[lookupPrefix(namespaceUri: string)](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-lookupprefix-member(1))|Gets the prefix corresponding to the specified namespace.|
|[CustomXmlReplaceChildNodeOptions](/javascript/api/word/word.customxmlreplacechildnodeoptions)|[name](/javascript/api/word/word.customxmlreplacechildnodeoptions#word-word-customxmlreplacechildnodeoptions-name-member)|If provided, specifies the base name of the replacement element.|
||[namespaceUri](/javascript/api/word/word.customxmlreplacechildnodeoptions#word-word-customxmlreplacechildnodeoptions-namespaceuri-member)|If provided, specifies the namespace of the replacement element.|
||[nodeType](/javascript/api/word/word.customxmlreplacechildnodeoptions#word-word-customxmlreplacechildnodeoptions-nodetype-member)|If provided, specifies the type of the replacement node.|
||[nodeValue](/javascript/api/word/word.customxmlreplacechildnodeoptions#word-word-customxmlreplacechildnodeoptions-nodevalue-member)|If provided, specifies the value of the replacement node for those nodes that allow text.|
|[CustomXmlSchema](/javascript/api/word/word.customxmlschema)|[delete()](/javascript/api/word/word.customxmlschema#word-word-customxmlschema-delete-member(1))|Deletes this schema from the Word.CustomXmlSchemaCollection object.|
||[location](/javascript/api/word/word.customxmlschema#word-word-customxmlschema-location-member)|Gets the location of the schema on a computer.|
||[namespaceUri](/javascript/api/word/word.customxmlschema#word-word-customxmlschema-namespaceuri-member)|Gets the unique address identifier for the namespace of the `CustomXmlSchema` object.|
||[reload()](/javascript/api/word/word.customxmlschema#word-word-customxmlschema-reload-member(1))|Reloads the schema from a file.|
|[CustomXmlSchemaCollection](/javascript/api/word/word.customxmlschemacollection)|[add(options?: Word.CustomXmlAddSchemaOptions)](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-add-member(1))|Adds one or more schemas to the schema collection that can then be added to a stream in the data store and to the schema library.|
||[addCollection(schemaCollection: Word.CustomXmlSchemaCollection)](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-addcollection-member(1))|Adds an existing schema collection to the current schema collection.|
||[getCount()](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-getitem-member(1))|Returns a `CustomXmlSchema` object that represents the specified item in the collection.|
||[getNamespaceUri()](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-getnamespaceuri-member(1))|Returns the number of items in the collection.|
||[items](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-items-member)|Gets the loaded child items in this collection.|
||[validate()](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-validate-member(1))|Specifies whether the schemas in the schema collection are valid (conforms to the syntactic rules of XML and the rules for a specified vocabulary).|
|[CustomXmlValidationError](/javascript/api/word/word.customxmlvalidationerror)|[delete()](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-delete-member(1))|Deletes this `CustomXmlValidationError` object.|
||[errorCode](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-errorcode-member)|Gets an integer representing the validation error in the `CustomXmlValidationError` object.|
||[name](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-name-member)|Gets the name of the error in the `CustomXmlValidationError` object.If no errors exist, the property returns `Nothing`|
||[node](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-node-member)|Gets the node associated with this `CustomXmlValidationError` object, if any exist.If no nodes exist, the property returns `Nothing`.|
||[text](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-text-member)|Gets the text in the `CustomXmlValidationError` object.|
||[type](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-type-member)|Gets the type of error generated from the `CustomXmlValidationError` object.|
|[CustomXmlValidationErrorCollection](/javascript/api/word/word.customxmlvalidationerrorcollection)|[add(node: Word.CustomXmlNode, errorName: string, options?: Word.CustomXmlAddValidationErrorOptions)](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-add-member(1))|Adds a `CustomXmlValidationError` object containing an XML validation error to the `CustomXmlValidationErrorCollection` object.|
||[getCount()](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-getitem-member(1))|Returns a `CustomXmlValidationError` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-items-member)|Gets the loaded child items in this collection.|
|[DatePickerContentControl](/javascript/api/word/word.datepickercontentcontrol)|[appearance](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[color](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-color-member)|Specifies the red-green-blue (RGB) value of the color of the content control.|
||[copy()](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-copy-member(1))|Copies the content control from the active document to the Clipboard.|
||[cut()](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-cut-member(1))|Removes the content control from the active document and moves the content control to the Clipboard.|
||[dateCalendarType](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-datecalendartype-member)|Specifies a `CalendarType` value that represents the calendar type for the date picker content control.|
||[dateDisplayFormat](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-datedisplayformat-member)|Specifies the format in which dates are displayed.|
||[dateDisplayLocale](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-datedisplaylocale-member)|Specifies a `LanguageId` that represents the language format for the date displayed in the date picker content control.|
||[dateStorageFormat](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-datestorageformat-member)|Specifies a `ContentControlDateStorageFormat` value that represents the format for storage and retrieval of dates when the date picker content control is bound to the XML data store of the active document.|
||[delete(deleteContents?: boolean)](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-delete-member(1))|Deletes this content control and the contents of the content control.|
||[id](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-id-member)|Gets the identification for the content control.|
||[isTemporary](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-istemporary-member)|Specifies whether to remove the content control from the active document when the user edits the contents of the control.|
||[level](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-level-member)|Specifies the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.|
||[lockContentControl](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-lockcontentcontrol-member)|Specifies if the content control is locked (can't be deleted).|
||[lockContents](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-lockcontents-member)|Specifies if the contents of the content control are locked (not editable).|
||[placeholderText](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-placeholdertext-member)|Returns a `BuildingBlock` object that represents the placeholder text for the content control.|
||[range](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-range-member)|Gets a `Range` object that represents the contents of the content control in the active document.|
||[setPlaceholderText(options?: Word.ContentControlPlaceholderOptions)](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-setplaceholdertext-member(1))|Sets the placeholder text that displays in the content control until a user enters their own text.|
||[showingPlaceholderText](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-showingplaceholdertext-member)|Gets whether the placeholder text for the content control is being displayed.|
||[tag](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-tag-member)|Specifies a tag to identify the content control.|
||[title](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-title-member)|Specifies the title for the content control.|
||[xmlMapping](/javascript/api/word/word.datepickercontentcontrol#word-word-datepickercontentcontrol-xmlmapping-member)|Gets an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[Document](/javascript/api/word/word.document)|[attachedTemplate](/javascript/api/word/word.document#word-word-document-attachedtemplate-member)|Specifies a `Template` object that represents the template attached to the document.|
||[autoHyphenation](/javascript/api/word/word.document#word-word-document-autohyphenation-member)|Specifies if automatic hyphenation is turned on for the document.|
||[autoSaveOn](/javascript/api/word/word.document#word-word-document-autosaveon-member)|Specifies if the edits in the document are automatically saved.|
||[bibliography](/javascript/api/word/word.document#word-word-document-bibliography-member)|Returns a `Bibliography` object that represents the bibliography references contained within the document.|
||[bookmarks](/javascript/api/word/word.document#word-word-document-bookmarks-member)|Returns a `BookmarkCollection` object that represents all the bookmarks in the document.|
||[consecutiveHyphensLimit](/javascript/api/word/word.document#word-word-document-consecutivehyphenslimit-member)|Specifies the maximum number of consecutive lines that can end with hyphens.|
||[detectLanguage()](/javascript/api/word/word.document#word-word-document-detectlanguage-member(1))|Analyzes the document text to determine the language.|
||[documentLibraryVersions](/javascript/api/word/word.document#word-word-document-documentlibraryversions-member)|Returns a `DocumentLibraryVersionCollection` object that represents the collection of versions of a shared document that has versioning enabled and that's stored in a document library on a server.|
||[frames](/javascript/api/word/word.document#word-word-document-frames-member)|Returns a `FrameCollection` object that represents all the frames in the document.|
||[hyperlinks](/javascript/api/word/word.document#word-word-document-hyperlinks-member)|Returns a `HyperlinkCollection` object that represents all the hyperlinks in the document.|
||[hyphenateCaps](/javascript/api/word/word.document#word-word-document-hyphenatecaps-member)|Specifies whether words in all capital letters can be hyphenated.|
||[indexes](/javascript/api/word/word.document#word-word-document-indexes-member)|Returns an `IndexCollection` object that represents all the indexes in the document.|
||[languageDetected](/javascript/api/word/word.document#word-word-document-languagedetected-member)|Specifies whether Microsoft Word has detected the language of the document text.|
||[manualHyphenation()](/javascript/api/word/word.document#word-word-document-manualhyphenation-member(1))|Initiates manual hyphenation of a document, one line at a time.|
||[pageSetup](/javascript/api/word/word.document#word-word-document-pagesetup-member)|Returns a `PageSetup` object that's associated with the document.|
|[DocumentLibraryVersion](/javascript/api/word/word.documentlibraryversion)|[comments](/javascript/api/word/word.documentlibraryversion#word-word-documentlibraryversion-comments-member)|Gets any optional comments associated with this version of the shared document.|
||[modified](/javascript/api/word/word.documentlibraryversion#word-word-documentlibraryversion-modified-member)|Gets the date and time at which this version of the shared document was last saved to the server.|
||[modifiedBy](/javascript/api/word/word.documentlibraryversion#word-word-documentlibraryversion-modifiedby-member)|Gets the name of the user who last saved this version of the shared document to the server.|
|[DocumentLibraryVersionCollection](/javascript/api/word/word.documentlibraryversioncollection)|[getItem(index: number)](/javascript/api/word/word.documentlibraryversioncollection#word-word-documentlibraryversioncollection-getitem-member(1))|Gets a `DocumentLibraryVersion` object by its index in the collection.|
||[isVersioningEnabled()](/javascript/api/word/word.documentlibraryversioncollection#word-word-documentlibraryversioncollection-isversioningenabled-member(1))|Returns whether the document library in which the active document is saved on the server is configured to create a backup copy, or version, each time the file is edited on the website.|
||[items](/javascript/api/word/word.documentlibraryversioncollection#word-word-documentlibraryversioncollection-items-member)|Gets the loaded child items in this collection.|
|[DropCap](/javascript/api/word/word.dropcap)|[clear()](/javascript/api/word/word.dropcap#word-word-dropcap-clear-member(1))|Removes the dropped capital letter formatting.|
||[distanceFromText](/javascript/api/word/word.dropcap#word-word-dropcap-distancefromtext-member)|Gets the distance (in points) between the dropped capital letter and the paragraph text.|
||[enable()](/javascript/api/word/word.dropcap#word-word-dropcap-enable-member(1))|Formats the first character in the specified paragraph as a dropped capital letter.|
||[fontName](/javascript/api/word/word.dropcap#word-word-dropcap-fontname-member)|Gets the name of the font for the dropped capital letter.|
||[linesToDrop](/javascript/api/word/word.dropcap#word-word-dropcap-linestodrop-member)|Gets the height (in lines) of the dropped capital letter.|
||[position](/javascript/api/word/word.dropcap#word-word-dropcap-position-member)|Gets the position of the dropped capital letter.|
|[Field](/javascript/api/word/word.field)|[copyToClipboard()](/javascript/api/word/word.field#word-word-field-copytoclipboard-member(1))|Copies the field to the Clipboard.|
||[cut()](/javascript/api/word/word.field#word-word-field-cut-member(1))|Removes the field from the document and places it on the Clipboard.|
||[doClick()](/javascript/api/word/word.field#word-word-field-doclick-member(1))|Clicks the field.|
||[linkFormat](/javascript/api/word/word.field#word-word-field-linkformat-member)|Gets a `LinkFormat` object that represents the link options of the field.|
||[oleFormat](/javascript/api/word/word.field#word-word-field-oleformat-member)|Gets an `OleFormat` object that represents the OLE characteristics (other than linking) for the field.|
||[unlink()](/javascript/api/word/word.field#word-word-field-unlink-member(1))|Replaces the field with its most recent result.|
||[updateSource()](/javascript/api/word/word.field#word-word-field-updatesource-member(1))|Saves the changes made to the results of an {@link https://support.microsoft.com/office/1c34d6d6-0de3-4b5c-916a-2ff950fb629e | INCLUDETEXT field} back to the source document.|
|[FillFormat](/javascript/api/word/word.fillformat)|[backgroundColor](/javascript/api/word/word.fillformat#word-word-fillformat-backgroundcolor-member)|Returns a `ColorFormat` object that represents the background color for the fill.|
||[foregroundColor](/javascript/api/word/word.fillformat#word-word-fillformat-foregroundcolor-member)|Returns a `ColorFormat` object that represents the foreground color for the fill.|
||[gradientAngle](/javascript/api/word/word.fillformat#word-word-fillformat-gradientangle-member)|Specifies the angle of the gradient fill.|
||[gradientColorType](/javascript/api/word/word.fillformat#word-word-fillformat-gradientcolortype-member)|Gets the gradient color type.|
||[gradientDegree](/javascript/api/word/word.fillformat#word-word-fillformat-gradientdegree-member)|Returns how dark or light a one-color gradient fill is.|
||[gradientStyle](/javascript/api/word/word.fillformat#word-word-fillformat-gradientstyle-member)|Returns the gradient style for the fill.|
||[gradientVariant](/javascript/api/word/word.fillformat#word-word-fillformat-gradientvariant-member)|Returns the gradient variant for the fill as an integer value from 1 to 4 for most gradient fills.|
||[isVisible](/javascript/api/word/word.fillformat#word-word-fillformat-isvisible-member)|Specifies if the object, or the formatting applied to it, is visible.|
||[pattern](/javascript/api/word/word.fillformat#word-word-fillformat-pattern-member)|Returns a `PatternType` value that represents the pattern applied to the fill or line.|
||[presetGradientType](/javascript/api/word/word.fillformat#word-word-fillformat-presetgradienttype-member)|Returns the preset gradient type for the fill.|
||[presetTexture](/javascript/api/word/word.fillformat#word-word-fillformat-presettexture-member)|Gets the preset texture.|
||[rotateWithObject](/javascript/api/word/word.fillformat#word-word-fillformat-rotatewithobject-member)|Specifies whether the fill rotates with the shape.|
||[setOneColorGradient(style: Word.GradientStyle, variant: number, degree: number)](/javascript/api/word/word.fillformat#word-word-fillformat-setonecolorgradient-member(1))|Sets the fill to a one-color gradient.|
||[setPatterned(pattern: Word.PatternType)](/javascript/api/word/word.fillformat#word-word-fillformat-setpatterned-member(1))|Sets the fill to a pattern.|
||[setPresetGradient(style: Word.GradientStyle, variant: number, presetGradientType: Word.PresetGradientType)](/javascript/api/word/word.fillformat#word-word-fillformat-setpresetgradient-member(1))|Sets the fill to a preset gradient.|
||[setPresetTextured(presetTexture: Word.PresetTexture)](/javascript/api/word/word.fillformat#word-word-fillformat-setpresettextured-member(1))|Sets the fill to a preset texture.|
||[setTwoColorGradient(style: Word.GradientStyle, variant: number)](/javascript/api/word/word.fillformat#word-word-fillformat-settwocolorgradient-member(1))|Sets the fill to a two-color gradient.|
||[solid()](/javascript/api/word/word.fillformat#word-word-fillformat-solid-member(1))|Sets the fill to a uniform color.|
||[textureAlignment](/javascript/api/word/word.fillformat#word-word-fillformat-texturealignment-member)|Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.|
||[textureHorizontalScale](/javascript/api/word/word.fillformat#word-word-fillformat-texturehorizontalscale-member)|Specifies the horizontal scaling factor for the texture fill.|
||[textureName](/javascript/api/word/word.fillformat#word-word-fillformat-texturename-member)|Returns the name of the custom texture file for the fill.|
||[textureOffsetX](/javascript/api/word/word.fillformat#word-word-fillformat-textureoffsetx-member)|Specifies the horizontal offset of the texture from the origin in points.|
||[textureOffsetY](/javascript/api/word/word.fillformat#word-word-fillformat-textureoffsety-member)|Specifies the vertical offset of the texture.|
||[textureTile](/javascript/api/word/word.fillformat#word-word-fillformat-texturetile-member)|Specifies whether the texture is tiled.|
||[textureType](/javascript/api/word/word.fillformat#word-word-fillformat-texturetype-member)|Returns the texture type for the fill.|
||[textureVerticalScale](/javascript/api/word/word.fillformat#word-word-fillformat-textureverticalscale-member)|Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.|
||[transparency](/javascript/api/word/word.fillformat#word-word-fillformat-transparency-member)|Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).|
||[type](/javascript/api/word/word.fillformat#word-word-fillformat-type-member)|Gets the fill format type.|
|[Font](/javascript/api/word/word.font)|[allCaps](/javascript/api/word/word.font#word-word-font-allcaps-member)|Specifies whether the font is formatted as all capital letters, which makes lowercase letters appear as uppercase letters.|
||[boldBidirectional](/javascript/api/word/word.font#word-word-font-boldbidirectional-member)|Specifies whether the font is formatted as bold in a right-to-left language document.|
||[borders](/javascript/api/word/word.font#word-word-font-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the font.|
||[colorIndex](/javascript/api/word/word.font#word-word-font-colorindex-member)|Specifies a `ColorIndex` value that represents the color for the font.|
||[colorIndexBidirectional](/javascript/api/word/word.font#word-word-font-colorindexbidirectional-member)|Specifies the color for the `Font` object in a right-to-left language document.|
||[contextualAlternates](/javascript/api/word/word.font#word-word-font-contextualalternates-member)|Specifies whether contextual alternates are enabled for the font.|
||[decreaseFontSize()](/javascript/api/word/word.font#word-word-font-decreasefontsize-member(1))|Decreases the font size to the next available size.|
||[diacriticColor](/javascript/api/word/word.font#word-word-font-diacriticcolor-member)|Specifies the color to be used for diacritics for the `Font` object.|
||[disableCharacterSpaceGrid](/javascript/api/word/word.font#word-word-font-disablecharacterspacegrid-member)|Specifies whether Microsoft Word ignores the number of characters per line for the corresponding `Font` object.|
||[emboss](/javascript/api/word/word.font#word-word-font-emboss-member)|Specifies whether the font is formatted as embossed.|
||[emphasisMark](/javascript/api/word/word.font#word-word-font-emphasismark-member)|Specifies an `EmphasisMark` value that represents the emphasis mark for a character or designated character string.|
||[engrave](/javascript/api/word/word.font#word-word-font-engrave-member)|Specifies whether the font is formatted as engraved.|
||[fill](/javascript/api/word/word.font#word-word-font-fill-member)|Returns a `FillFormat` object that contains fill formatting properties for the font used by the range of text.|
||[glow](/javascript/api/word/word.font#word-word-font-glow-member)|Returns a `GlowFormat` object that represents the glow formatting for the font used by the range of text.|
||[increaseFontSize()](/javascript/api/word/word.font#word-word-font-increasefontsize-member(1))|Increases the font size to the next available size.|
||[italicBidirectional](/javascript/api/word/word.font#word-word-font-italicbidirectional-member)|Specifies whether the font is italicized in a right-to-left language document.|
||[kerning](/javascript/api/word/word.font#word-word-font-kerning-member)|Specifies the minimum font size for which Microsoft Word will adjust kerning automatically.|
||[ligature](/javascript/api/word/word.font#word-word-font-ligature-member)|Specifies the ligature setting for the `Font` object.|
||[line](/javascript/api/word/word.font#word-word-font-line-member)|Returns a `LineFormat` object that specifies the formatting for a line.|
||[nameAscii](/javascript/api/word/word.font#word-word-font-nameascii-member)|Specifies the font used for Latin text (characters with character codes from 0 (zero) through 127).|
||[nameBidirectional](/javascript/api/word/word.font#word-word-font-namebidirectional-member)|Specifies the font name in a right-to-left language document.|
||[nameFarEast](/javascript/api/word/word.font#word-word-font-namefareast-member)|Specifies the East Asian font name.|
||[nameOther](/javascript/api/word/word.font#word-word-font-nameother-member)|Specifies the font used for characters with codes from 128 through 255.|
||[numberForm](/javascript/api/word/word.font#word-word-font-numberform-member)|Specifies the number form setting for an OpenType font.|
||[numberSpacing](/javascript/api/word/word.font#word-word-font-numberspacing-member)|Specifies the number spacing setting for the font.|
||[outline](/javascript/api/word/word.font#word-word-font-outline-member)|Specifies if the font is formatted as outlined.|
||[position](/javascript/api/word/word.font#word-word-font-position-member)|Specifies the position of text (in points) relative to the base line.|
||[reflection](/javascript/api/word/word.font#word-word-font-reflection-member)|Returns a `ReflectionFormat` object that represents the reflection formatting for a shape.|
||[reset()](/javascript/api/word/word.font#word-word-font-reset-member(1))|Removes manual character formatting.|
||[scaling](/javascript/api/word/word.font#word-word-font-scaling-member)|Specifies the scaling percentage applied to the font.|
||[setAsTemplateDefault()](/javascript/api/word/word.font#word-word-font-setastemplatedefault-member(1))|Sets the specified font formatting as the default for the active document and all new documents based on the active template.|
||[shadow](/javascript/api/word/word.font#word-word-font-shadow-member)|Specifies if the font is formatted as shadowed.|
||[sizeBidirectional](/javascript/api/word/word.font#word-word-font-sizebidirectional-member)|Specifies the font size in points for right-to-left text.|
||[smallCaps](/javascript/api/word/word.font#word-word-font-smallcaps-member)|Specifies whether the font is formatted as small caps, which makes lowercase letters appear as small uppercase letters.|
||[spacing](/javascript/api/word/word.font#word-word-font-spacing-member)|Specifies the spacing between characters.|
||[stylisticSet](/javascript/api/word/word.font#word-word-font-stylisticset-member)|Specifies the stylistic set for the font.|
||[textColor](/javascript/api/word/word.font#word-word-font-textcolor-member)|Returns a `ColorFormat` object that represents the color for the font.|
||[textShadow](/javascript/api/word/word.font#word-word-font-textshadow-member)|Returns a `ShadowFormat` object that specifies the shadow formatting for the font.|
||[threeDimensionalFormat](/javascript/api/word/word.font#word-word-font-threedimensionalformat-member)|Returns a `ThreeDimensionalFormat` object that contains 3-dimensional (3D) effect formatting properties for the font.|
||[underlineColor](/javascript/api/word/word.font#word-word-font-underlinecolor-member)|Specifies the color of the underline for the `Font` object.|
|[Frame](/javascript/api/word/word.frame)|[borders](/javascript/api/word/word.frame#word-word-frame-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the frame.|
||[copy()](/javascript/api/word/word.frame#word-word-frame-copy-member(1))|Copies the frame to the Clipboard.|
||[cut()](/javascript/api/word/word.frame#word-word-frame-cut-member(1))|Removes the frame from the document and places it on the Clipboard.|
||[delete()](/javascript/api/word/word.frame#word-word-frame-delete-member(1))|Deletes the frame.|
||[height](/javascript/api/word/word.frame#word-word-frame-height-member)|Specifies the height (in points) of the frame.|
||[heightRule](/javascript/api/word/word.frame#word-word-frame-heightrule-member)|Specifies a `FrameSizeRule` value that represents the rule for determining the height of the frame.|
||[horizontalDistanceFromText](/javascript/api/word/word.frame#word-word-frame-horizontaldistancefromtext-member)|Specifies the horizontal distance between the frame and the surrounding text, in points.|
||[horizontalPosition](/javascript/api/word/word.frame#word-word-frame-horizontalposition-member)|Specifies the horizontal distance between the edge of the frame and the item specified by the `relativeHorizontalPosition` property.|
||[lockAnchor](/javascript/api/word/word.frame#word-word-frame-lockanchor-member)|Specifies if the frame is locked.|
||[range](/javascript/api/word/word.frame#word-word-frame-range-member)|Returns a `Range` object that represents the portion of the document that's contained within the frame.|
||[relativeHorizontalPosition](/javascript/api/word/word.frame#word-word-frame-relativehorizontalposition-member)|Specifies the relative horizontal position of the frame.|
||[relativeVerticalPosition](/javascript/api/word/word.frame#word-word-frame-relativeverticalposition-member)|Specifies the relative vertical position of the frame.|
||[select()](/javascript/api/word/word.frame#word-word-frame-select-member(1))|Selects the frame.|
||[shading](/javascript/api/word/word.frame#word-word-frame-shading-member)|Returns a `ShadingUniversal` object that refers to the shading formatting for the frame.|
||[textWrap](/javascript/api/word/word.frame#word-word-frame-textwrap-member)|Specifies if document text wraps around the frame.|
||[verticalDistanceFromText](/javascript/api/word/word.frame#word-word-frame-verticaldistancefromtext-member)|Specifies the vertical distance (in points) between the frame and the surrounding text.|
||[verticalPosition](/javascript/api/word/word.frame#word-word-frame-verticalposition-member)|Specifies the vertical distance between the edge of the frame and the item specified by the `relativeVerticalPosition` property.|
||[width](/javascript/api/word/word.frame#word-word-frame-width-member)|Specifies the width (in points) of the frame.|
||[widthRule](/javascript/api/word/word.frame#word-word-frame-widthrule-member)|Specifies the rule used to determine the width of the frame.|
|[FrameCollection](/javascript/api/word/word.framecollection)|[add(range: Word.Range)](/javascript/api/word/word.framecollection#word-word-framecollection-add-member(1))|Returns a `Frame` object that represents a new frame added to a range, selection, or document.|
||[delete()](/javascript/api/word/word.framecollection#word-word-framecollection-delete-member(1))|Deletes the `FrameCollection` object.|
||[getItem(index: number)](/javascript/api/word/word.framecollection#word-word-framecollection-getitem-member(1))|Gets a `Frame` object by its index in the collection.|
||[items](/javascript/api/word/word.framecollection#word-word-framecollection-items-member)|Gets the loaded child items in this collection.|
|[GlowFormat](/javascript/api/word/word.glowformat)|[color](/javascript/api/word/word.glowformat#word-word-glowformat-color-member)|Returns a `ColorFormat` object that represents the color for a glow effect.|
||[radius](/javascript/api/word/word.glowformat#word-word-glowformat-radius-member)|Specifies the length of the radius for a glow effect.|
||[transparency](/javascript/api/word/word.glowformat#word-word-glowformat-transparency-member)|Specifies the degree of transparency for the glow effect as a value between 0.0 (opaque) and 1.0 (clear).|
|[GroupContentControl](/javascript/api/word/word.groupcontentcontrol)|[appearance](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[color](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-color-member)|Specifies the red-green-blue (RGB) value of the color of the content control.|
||[copy()](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-copy-member(1))|Copies the content control from the active document to the Clipboard.|
||[cut()](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-cut-member(1))|Removes the content control from the active document and moves the content control to the Clipboard.|
||[delete(deleteContents: boolean)](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-delete-member(1))|Deletes the content control and optionally its contents.|
||[id](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-id-member)|Returns the identification for the content control.|
||[isTemporary](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-istemporary-member)|Specifies whether to remove the content control from the active document when the user edits the contents of the control.|
||[level](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-level-member)|Gets the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.|
||[lockContentControl](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-lockcontentcontrol-member)|Specifies if the content control is locked (can't be deleted).|
||[lockContents](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-lockcontents-member)|Specifies if the contents of the content control are locked (not editable).|
||[placeholderText](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-placeholdertext-member)|Returns a `BuildingBlock` object that represents the placeholder text for the content control.|
||[range](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-range-member)|Gets a `Range` object that represents the contents of the content control in the active document.|
||[setPlaceholderText(options?: Word.ContentControlPlaceholderOptions)](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-setplaceholdertext-member(1))|Sets the placeholder text that displays in the content control until a user enters their own text.|
||[showingPlaceholderText](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-showingplaceholdertext-member)|Returns whether the placeholder text for the content control is being displayed.|
||[tag](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-tag-member)|Specifies a tag to identify the content control.|
||[title](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-title-member)|Specifies the title for the content control.|
||[ungroup()](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-ungroup-member(1))|Removes the group content control from the document so that its child content controls are no longer nested and can be freely edited.|
||[xmlMapping](/javascript/api/word/word.groupcontentcontrol#word-word-groupcontentcontrol-xmlmapping-member)|Gets an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[Hyperlink](/javascript/api/word/word.hyperlink)|[addToFavorites()](/javascript/api/word/word.hyperlink#word-word-hyperlink-addtofavorites-member(1))|Creates a shortcut to the document or hyperlink and adds it to the **Favorites** folder.|
||[address](/javascript/api/word/word.hyperlink#word-word-hyperlink-address-member)|Specifies the address (for example, a file name or URL) of the hyperlink.|
||[createNewDocument(fileName: string, editNow: boolean, overwrite: boolean)](/javascript/api/word/word.hyperlink#word-word-hyperlink-createnewdocument-member(1))|Creates a new document linked to the hyperlink.|
||[delete()](/javascript/api/word/word.hyperlink#word-word-hyperlink-delete-member(1))|Deletes the hyperlink.|
||[emailSubject](/javascript/api/word/word.hyperlink#word-word-hyperlink-emailsubject-member)|Specifies the text string for the hyperlink's subject line.|
||[isExtraInfoRequired](/javascript/api/word/word.hyperlink#word-word-hyperlink-isextrainforequired-member)|Returns `true` if extra information is required to resolve the hyperlink.|
||[name](/javascript/api/word/word.hyperlink#word-word-hyperlink-name-member)|Returns the name of the `Hyperlink` object.|
||[range](/javascript/api/word/word.hyperlink#word-word-hyperlink-range-member)|Returns a `Range` object that represents the portion of the document that's contained within the hyperlink.|
||[screenTip](/javascript/api/word/word.hyperlink#word-word-hyperlink-screentip-member)|Specifies the text that appears as a ScreenTip when the mouse pointer is positioned over the hyperlink.|
||[subAddress](/javascript/api/word/word.hyperlink#word-word-hyperlink-subaddress-member)|Specifies a named location in the destination of the hyperlink.|
||[target](/javascript/api/word/word.hyperlink#word-word-hyperlink-target-member)|Specifies the name of the frame or window in which to load the hyperlink.|
||[textToDisplay](/javascript/api/word/word.hyperlink#word-word-hyperlink-texttodisplay-member)|Specifies the hyperlink's visible text in the document.|
||[type](/javascript/api/word/word.hyperlink#word-word-hyperlink-type-member)|Returns the hyperlink type.|
|[HyperlinkAddOptions](/javascript/api/word/word.hyperlinkaddoptions)|[address](/javascript/api/word/word.hyperlinkaddoptions#word-word-hyperlinkaddoptions-address-member)|If provided, specifies the address (e.g., URL or file path) of the hyperlink.|
||[screenTip](/javascript/api/word/word.hyperlinkaddoptions#word-word-hyperlinkaddoptions-screentip-member)|If provided, specifies the text that appears as a tooltip.|
||[subAddress](/javascript/api/word/word.hyperlinkaddoptions#word-word-hyperlinkaddoptions-subaddress-member)|If provided, specifies the location within the file or document.|
||[target](/javascript/api/word/word.hyperlinkaddoptions#word-word-hyperlinkaddoptions-target-member)|If provided, specifies the name of the frame or window in which to load the hyperlink.|
||[textToDisplay](/javascript/api/word/word.hyperlinkaddoptions#word-word-hyperlinkaddoptions-texttodisplay-member)|If provided, specifies the visible text of the hyperlink.|
|[HyperlinkCollection](/javascript/api/word/word.hyperlinkcollection)|[add(anchor: Word.Range, options?: Word.HyperlinkAddOptions)](/javascript/api/word/word.hyperlinkcollection#word-word-hyperlinkcollection-add-member(1))|Returns a `Hyperlink` object that represents a new hyperlink added to a range, selection, or document.|
||[items](/javascript/api/word/word.hyperlinkcollection#word-word-hyperlinkcollection-items-member)|Gets the loaded child items in this collection.|
|[Index](/javascript/api/word/word.index)|[delete()](/javascript/api/word/word.index#word-word-index-delete-member(1))|Deletes this index.|
||[filter](/javascript/api/word/word.index#word-word-index-filter-member)|Gets a value that represents how Microsoft Word classifies the first character of entries in the index.|
||[headingSeparator](/javascript/api/word/word.index#word-word-index-headingseparator-member)|Gets the text between alphabetical groups (entries that start with the same letter) in the index.|
||[indexLanguage](/javascript/api/word/word.index#word-word-index-indexlanguage-member)|Gets a `LanguageId` value that represents the sorting language to use for the index.|
||[numberOfColumns](/javascript/api/word/word.index#word-word-index-numberofcolumns-member)|Gets the number of columns for each page of the index.|
||[range](/javascript/api/word/word.index#word-word-index-range-member)|Returns a `Range` object that represents the portion of the document that is contained within the index.|
||[rightAlignPageNumbers](/javascript/api/word/word.index#word-word-index-rightalignpagenumbers-member)|Specifies if page numbers are aligned with the right margin in the index.|
||[separateAccentedLetterHeadings](/javascript/api/word/word.index#word-word-index-separateaccentedletterheadings-member)|Gets if the index contains separate headings for accented letters (for example, words that begin with "À" are under|
||[sortBy](/javascript/api/word/word.index#word-word-index-sortby-member)|Specifies the sorting criteria for the index.|
||[tabLeader](/javascript/api/word/word.index#word-word-index-tableader-member)|Specifies the leader character between entries in the index and their associated page numbers.|
||[type](/javascript/api/word/word.index#word-word-index-type-member)|Gets the index type.|
|[IndexAddOptions](/javascript/api/word/word.indexaddoptions)|[headingSeparator](/javascript/api/word/word.indexaddoptions#word-word-indexaddoptions-headingseparator-member)|If provided, specifies the text between alphabetical groups (entries that start with the same letter) in the index.|
||[indexLanguage](/javascript/api/word/word.indexaddoptions#word-word-indexaddoptions-indexlanguage-member)|If provided, specifies the sorting language to be used for the index being added.|
||[numberOfColumns](/javascript/api/word/word.indexaddoptions#word-word-indexaddoptions-numberofcolumns-member)|If provided, specifies the number of columns for each page of the index.|
||[rightAlignPageNumbers](/javascript/api/word/word.indexaddoptions#word-word-indexaddoptions-rightalignpagenumbers-member)|If provided, specifies whether the page numbers in the generated index are aligned with the right margin.|
||[separateAccentedLetterHeadings](/javascript/api/word/word.indexaddoptions#word-word-indexaddoptions-separateaccentedletterheadings-member)|If provided, specifies whether to include separate headings for accented letters in the index.|
||[sortBy](/javascript/api/word/word.indexaddoptions#word-word-indexaddoptions-sortby-member)|If provided, specifies the sorting criteria to be used for the index being added.|
||[type](/javascript/api/word/word.indexaddoptions#word-word-indexaddoptions-type-member)|If provided, specifies whether subentries are on the same line (run-in) as the main entry or on a separate line (indented) from the main entry.|
|[IndexCollection](/javascript/api/word/word.indexcollection)|[add(range: Word.Range, indexAddOptions?: Word.IndexAddOptions)](/javascript/api/word/word.indexcollection#word-word-indexcollection-add-member(1))|Returns an `Index` object that represents a new index added to the document.|
||[getFormat()](/javascript/api/word/word.indexcollection#word-word-indexcollection-getformat-member(1))|Gets the `IndexFormat` value that represents the formatting for the indexes in the document.|
||[getItem(index: number)](/javascript/api/word/word.indexcollection#word-word-indexcollection-getitem-member(1))|Gets an `Index` object by its index in the collection.|
||[items](/javascript/api/word/word.indexcollection#word-word-indexcollection-items-member)|Gets the loaded child items in this collection.|
||[markAllEntries(range: Word.Range, markAllEntriesOptions?: Word.IndexMarkAllEntriesOptions)](/javascript/api/word/word.indexcollection#word-word-indexcollection-markallentries-member(1))|Inserts an {@link https://support.microsoft.com/office/abaf7c78-6e21-418d-bf8b-f8186d2e4d08 | XE (Index Entry) field} after all instances of the text in the range.|
|[IndexMarkAllEntriesOptions](/javascript/api/word/word.indexmarkallentriesoptions)|[bold](/javascript/api/word/word.indexmarkallentriesoptions#word-word-indexmarkallentriesoptions-bold-member)|If provided, specifies whether to add bold formatting to page numbers for index entries.|
||[bookmarkName](/javascript/api/word/word.indexmarkallentriesoptions#word-word-indexmarkallentriesoptions-bookmarkname-member)|If provided, specifies the bookmark name that marks the range of pages you want to appear in the index.|
||[crossReference](/javascript/api/word/word.indexmarkallentriesoptions#word-word-indexmarkallentriesoptions-crossreference-member)|If provided, specifies the cross-reference that will appear in the index.|
||[crossReferenceAutoText](/javascript/api/word/word.indexmarkallentriesoptions#word-word-indexmarkallentriesoptions-crossreferenceautotext-member)|If provided, specifies the name of the `AutoText` entry that contains the text for a cross-reference (if this property is specified, `crossReference` is ignored).|
||[entry](/javascript/api/word/word.indexmarkallentriesoptions#word-word-indexmarkallentriesoptions-entry-member)|If provided, specifies the text you want to appear in the index, in the form `MainEntry[:Subentry]`.|
||[entryAutoText](/javascript/api/word/word.indexmarkallentriesoptions#word-word-indexmarkallentriesoptions-entryautotext-member)|If provided, specifies the `AutoText` entry that contains the text you want to appear in the index (if this property is specified, `entry` is ignored).|
||[italic](/javascript/api/word/word.indexmarkallentriesoptions#word-word-indexmarkallentriesoptions-italic-member)|If provided, specifies whether to add italic formatting to page numbers for index entries.|
|[IndexMarkEntryOptions](/javascript/api/word/word.indexmarkentryoptions)|[bold](/javascript/api/word/word.indexmarkentryoptions#word-word-indexmarkentryoptions-bold-member)|If provided, specifies whether to add bold formatting to page numbers for index entries.|
||[bookmarkName](/javascript/api/word/word.indexmarkentryoptions#word-word-indexmarkentryoptions-bookmarkname-member)|If provided, specifies the bookmark name that marks the range of pages you want to appear in the index.|
||[crossReference](/javascript/api/word/word.indexmarkentryoptions#word-word-indexmarkentryoptions-crossreference-member)|If provided, specifies the cross-reference that will appear in the index.|
||[crossReferenceAutoText](/javascript/api/word/word.indexmarkentryoptions#word-word-indexmarkentryoptions-crossreferenceautotext-member)|If provided, specifies the name of the `AutoText` entry that contains the text for a cross-reference (if this property is specified, `crossReference` is ignored).|
||[entry](/javascript/api/word/word.indexmarkentryoptions#word-word-indexmarkentryoptions-entry-member)|If provided, specifies the text you want to appear in the index, in the form `MainEntry[:Subentry]`.|
||[entryAutoText](/javascript/api/word/word.indexmarkentryoptions#word-word-indexmarkentryoptions-entryautotext-member)|If provided, specifies the `AutoText` entry that contains the text you want to appear in the index (if this property is specified, `entry` is ignored).|
||[italic](/javascript/api/word/word.indexmarkentryoptions#word-word-indexmarkentryoptions-italic-member)|If provided, specifies whether to add italic formatting to page numbers for index entries.|
||[reading](/javascript/api/word/word.indexmarkentryoptions#word-word-indexmarkentryoptions-reading-member)|If provided, specifies whether to show an index entry in the right location when indexes are sorted phonetically (East Asian languages only).|
|[LineFormat](/javascript/api/word/word.lineformat)|[backgroundColor](/javascript/api/word/word.lineformat#word-word-lineformat-backgroundcolor-member)|Gets a `ColorFormat` object that represents the background color for a patterned line.|
||[beginArrowheadLength](/javascript/api/word/word.lineformat#word-word-lineformat-beginarrowheadlength-member)|Specifies the length of the arrowhead at the beginning of the line.|
||[beginArrowheadStyle](/javascript/api/word/word.lineformat#word-word-lineformat-beginarrowheadstyle-member)|Specifies the style of the arrowhead at the beginning of the line.|
||[beginArrowheadWidth](/javascript/api/word/word.lineformat#word-word-lineformat-beginarrowheadwidth-member)|Specifies the width of the arrowhead at the beginning of the line.|
||[dashStyle](/javascript/api/word/word.lineformat#word-word-lineformat-dashstyle-member)|Specifies the dash style for the line.|
||[endArrowheadLength](/javascript/api/word/word.lineformat#word-word-lineformat-endarrowheadlength-member)|Specifies the length of the arrowhead at the end of the line.|
||[endArrowheadStyle](/javascript/api/word/word.lineformat#word-word-lineformat-endarrowheadstyle-member)|Specifies the style of the arrowhead at the end of the line.|
||[endArrowheadWidth](/javascript/api/word/word.lineformat#word-word-lineformat-endarrowheadwidth-member)|Specifies the width of the arrowhead at the end of the line.|
||[foregroundColor](/javascript/api/word/word.lineformat#word-word-lineformat-foregroundcolor-member)|Gets a `ColorFormat` object that represents the foreground color for the line.|
||[insetPen](/javascript/api/word/word.lineformat#word-word-lineformat-insetpen-member)|Specifies if to draw lines inside a shape.|
||[isVisible](/javascript/api/word/word.lineformat#word-word-lineformat-isvisible-member)|Specifies if the object, or the formatting applied to it, is visible.|
||[pattern](/javascript/api/word/word.lineformat#word-word-lineformat-pattern-member)|Specifies the pattern applied to the line.|
||[style](/javascript/api/word/word.lineformat#word-word-lineformat-style-member)|Specifies the line format style.|
||[transparency](/javascript/api/word/word.lineformat#word-word-lineformat-transparency-member)|Specifies the degree of transparency of the line as a value between 0.0 (opaque) and 1.0 (clear).|
||[weight](/javascript/api/word/word.lineformat#word-word-lineformat-weight-member)|Specifies the thickness of the line in points.|
|[LineNumbering](/javascript/api/word/word.linenumbering)|[countBy](/javascript/api/word/word.linenumbering#word-word-linenumbering-countby-member)|Specifies the numeric increment for line numbers.|
||[distanceFromText](/javascript/api/word/word.linenumbering#word-word-linenumbering-distancefromtext-member)|Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.|
||[isActive](/javascript/api/word/word.linenumbering#word-word-linenumbering-isactive-member)|Specifies if line numbering is active for the specified document, section, or sections.|
||[restartMode](/javascript/api/word/word.linenumbering#word-word-linenumbering-restartmode-member)|Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.|
||[startingNumber](/javascript/api/word/word.linenumbering#word-word-linenumbering-startingnumber-member)|Specifies the starting line number.|
|[LinkFormat](/javascript/api/word/word.linkformat)|[breakLink()](/javascript/api/word/word.linkformat#word-word-linkformat-breaklink-member(1))|Breaks the link between the source file and the OLE object, picture, or linked field.|
||[isAutoUpdated](/javascript/api/word/word.linkformat#word-word-linkformat-isautoupdated-member)|Specifies if the link is updated automatically when the container file is opened or when the source file is changed.|
||[isLocked](/javascript/api/word/word.linkformat#word-word-linkformat-islocked-member)|Specifies if a `Field`, `InlineShape`, or `Shape` object is locked to prevent automatic updating.|
||[isPictureSavedWithDocument](/javascript/api/word/word.linkformat#word-word-linkformat-ispicturesavedwithdocument-member)|Specifies if the linked picture is saved with the document.|
||[sourceFullName](/javascript/api/word/word.linkformat#word-word-linkformat-sourcefullname-member)|Specifies the path and name of the source file for the linked OLE object, picture, or field.|
||[sourceName](/javascript/api/word/word.linkformat#word-word-linkformat-sourcename-member)|Gets the name of the source file for the linked OLE object, picture, or field.|
||[sourcePath](/javascript/api/word/word.linkformat#word-word-linkformat-sourcepath-member)|Gets the path of the source file for the linked OLE object, picture, or field.|
||[type](/javascript/api/word/word.linkformat#word-word-linkformat-type-member)|Gets the link type.|
|[ListFormat](/javascript/api/word/word.listformat)|[applyBulletDefault(defaultListBehavior: Word.DefaultListBehavior)](/javascript/api/word/word.listformat#word-word-listformat-applybulletdefault-member(1))|Adds bullets and formatting to the paragraphs in the range.|
||[applyListTemplateWithLevel(listTemplate: Word.ListTemplate, options?: Word.ListTemplateApplyOptions)](/javascript/api/word/word.listformat#word-word-listformat-applylisttemplatewithlevel-member(1))|Applies a list template with a specific level to the paragraphs in the range.|
||[applyNumberDefault(defaultListBehavior: Word.DefaultListBehavior)](/javascript/api/word/word.listformat#word-word-listformat-applynumberdefault-member(1))|Adds numbering and formatting to the paragraphs in the range.|
||[applyOutlineNumberDefault(defaultListBehavior: Word.DefaultListBehavior)](/javascript/api/word/word.listformat#word-word-listformat-applyoutlinenumberdefault-member(1))|Adds outline numbering and formatting to the paragraphs in the range.|
||[canContinuePreviousList(listTemplate: Word.ListTemplate)](/javascript/api/word/word.listformat#word-word-listformat-cancontinuepreviouslist-member(1))|Determines whether the `ListFormat` object can continue a previous list.|
||[convertNumbersToText(numberType: Word.NumberType)](/javascript/api/word/word.listformat#word-word-listformat-convertnumberstotext-member(1))|Converts numbers in the list to plain text.|
||[countNumberedItems(options?: Word.ListFormatCountNumberedItemsOptions)](/javascript/api/word/word.listformat#word-word-listformat-countnumbereditems-member(1))|Counts the numbered items in the list.|
||[isSingleList](/javascript/api/word/word.listformat#word-word-listformat-issinglelist-member)|Indicates whether the `ListFormat` object contains a single list.|
||[isSingleListTemplate](/javascript/api/word/word.listformat#word-word-listformat-issinglelisttemplate-member)|Indicates whether the `ListFormat` object contains a single list template.|
||[list](/javascript/api/word/word.listformat#word-word-listformat-list-member)|Returns a `List` object that represents the first formatted list contained in the `ListFormat` object.|
||[listIndent()](/javascript/api/word/word.listformat#word-word-listformat-listindent-member(1))|Indents the list by one level.|
||[listLevelNumber](/javascript/api/word/word.listformat#word-word-listformat-listlevelnumber-member)|Specifies the list level number for the first paragraph for the `ListFormat` object.|
||[listOutdent()](/javascript/api/word/word.listformat#word-word-listformat-listoutdent-member(1))|Outdents the list by one level.|
||[listString](/javascript/api/word/word.listformat#word-word-listformat-liststring-member)|Gets the string representation of the list value of the first paragraph in the range for the `ListFormat` object.|
||[listTemplate](/javascript/api/word/word.listformat#word-word-listformat-listtemplate-member)|Gets the list template associated with the `ListFormat` object.|
||[listType](/javascript/api/word/word.listformat#word-word-listformat-listtype-member)|Gets the type of the list for the `ListFormat` object.|
||[listValue](/javascript/api/word/word.listformat#word-word-listformat-listvalue-member)|Gets the numeric value of the the first paragraph in the range for the `ListFormat` object.|
||[removeNumbers(numberType: Word.NumberType)](/javascript/api/word/word.listformat#word-word-listformat-removenumbers-member(1))|Removes numbering from the list.|
|[ListFormatCountNumberedItemsOptions](/javascript/api/word/word.listformatcountnumbereditemsoptions)|[level](/javascript/api/word/word.listformatcountnumbereditemsoptions#word-word-listformatcountnumbereditemsoptions-level-member)|If provided, specifies the level to count.|
||[numberType](/javascript/api/word/word.listformatcountnumbereditemsoptions#word-word-listformatcountnumbereditemsoptions-numbertype-member)|If provided, specifies the type of number to count.|
|[ListTemplateApplyOptions](/javascript/api/word/word.listtemplateapplyoptions)|[applyLevel](/javascript/api/word/word.listtemplateapplyoptions#word-word-listtemplateapplyoptions-applylevel-member)|If provided, specifies the level to apply in the list template.|
||[applyTo](/javascript/api/word/word.listtemplateapplyoptions#word-word-listtemplateapplyoptions-applyto-member)|If provided, specifies which part of the list to apply the template to.|
||[continuePreviousList](/javascript/api/word/word.listtemplateapplyoptions#word-word-listtemplateapplyoptions-continuepreviouslist-member)|If provided, specifies whether to continue the previous list.|
||[defaultListBehavior](/javascript/api/word/word.listtemplateapplyoptions#word-word-listtemplateapplyoptions-defaultlistbehavior-member)|If provided, specifies the default list behavior.|
|[OleFormat](/javascript/api/word/word.oleformat)|[activate()](/javascript/api/word/word.oleformat#word-word-oleformat-activate-member(1))|Activates the `OleFormat` object.|
||[activateAs(classType: string)](/javascript/api/word/word.oleformat#word-word-oleformat-activateas-member(1))|Sets the Windows registry value that determines the default application used to activate the specified OLE object.|
||[classType](/javascript/api/word/word.oleformat#word-word-oleformat-classtype-member)|Specifies the class type for the specified OLE object, picture, or field.|
||[doVerb(verbIndex: Word.OleVerb)](/javascript/api/word/word.oleformat#word-word-oleformat-doverb-member(1))|Requests that the OLE object perform one of its available verbs.|
||[edit()](/javascript/api/word/word.oleformat#word-word-oleformat-edit-member(1))|Opens the OLE object for editing in the application it was created in.|
||[iconIndex](/javascript/api/word/word.oleformat#word-word-oleformat-iconindex-member)|Specifies the icon that is used when the `displayAsIcon` property is `true`.|
||[iconLabel](/javascript/api/word/word.oleformat#word-word-oleformat-iconlabel-member)|Specifies the text displayed below the icon for the OLE object.|
||[iconName](/javascript/api/word/word.oleformat#word-word-oleformat-iconname-member)|Specifies the program file in which the icon for the OLE object is stored.|
||[iconPath](/javascript/api/word/word.oleformat#word-word-oleformat-iconpath-member)|Gets the path of the file in which the icon for the OLE object is stored.|
||[isDisplayedAsIcon](/javascript/api/word/word.oleformat#word-word-oleformat-isdisplayedasicon-member)|Gets whether the specified object is displayed as an icon.|
||[isFormattingPreservedOnUpdate](/javascript/api/word/word.oleformat#word-word-oleformat-isformattingpreservedonupdate-member)|Specifies whether formatting done in Microsoft Word to the linked OLE object is preserved.|
||[label](/javascript/api/word/word.oleformat#word-word-oleformat-label-member)|Gets a string that's used to identify the portion of the source file that's being linked.|
||[open()](/javascript/api/word/word.oleformat#word-word-oleformat-open-member(1))|Opens the `OleFormat` object.|
||[progID](/javascript/api/word/word.oleformat#word-word-oleformat-progid-member)|Gets the programmatic identifier (`ProgId`) for the specified OLE object.|
|[Page](/javascript/api/word/word.page)|[breaks](/javascript/api/word/word.page#word-word-page-breaks-member)|Gets a `BreakCollection` object that represents the breaks on the page.|
|[PageSetup](/javascript/api/word/word.pagesetup)|[bookFoldPrinting](/javascript/api/word/word.pagesetup#word-word-pagesetup-bookfoldprinting-member)|Specifies whether Microsoft Word prints the document as a booklet.|
||[bookFoldPrintingSheets](/javascript/api/word/word.pagesetup#word-word-pagesetup-bookfoldprintingsheets-member)|Specifies the number of pages for each booklet.|
||[bookFoldReversePrinting](/javascript/api/word/word.pagesetup#word-word-pagesetup-bookfoldreverseprinting-member)|Specifies if Microsoft Word reverses the printing order for book fold printing of bidirectional or Asian language documents.|
||[bottomMargin](/javascript/api/word/word.pagesetup#word-word-pagesetup-bottommargin-member)|Specifies the distance (in points) between the bottom edge of the page and the bottom boundary of the body text.|
||[charsLine](/javascript/api/word/word.pagesetup#word-word-pagesetup-charsline-member)|Specifies the number of characters per line in the document grid.|
||[differentFirstPageHeaderFooter](/javascript/api/word/word.pagesetup#word-word-pagesetup-differentfirstpageheaderfooter-member)|Specifies whether the first page has a different header and footer.|
||[footerDistance](/javascript/api/word/word.pagesetup#word-word-pagesetup-footerdistance-member)|Specifies the distance between the footer and the bottom of the page in points.|
||[gutter](/javascript/api/word/word.pagesetup#word-word-pagesetup-gutter-member)|Specifies the amount (in points) of extra margin space added to each page in a document or section for binding.|
||[gutterPosition](/javascript/api/word/word.pagesetup#word-word-pagesetup-gutterposition-member)|Specifies on which side the gutter appears in a document.|
||[gutterStyle](/javascript/api/word/word.pagesetup#word-word-pagesetup-gutterstyle-member)|Specifies whether Microsoft Word uses gutters for the current document based on a right-to-left language or a left-to-right language.|
||[headerDistance](/javascript/api/word/word.pagesetup#word-word-pagesetup-headerdistance-member)|Specifies the distance between the header and the top of the page in points.|
||[layoutMode](/javascript/api/word/word.pagesetup#word-word-pagesetup-layoutmode-member)|Specifies the layout mode for the current document.|
||[leftMargin](/javascript/api/word/word.pagesetup#word-word-pagesetup-leftmargin-member)|Specifies the distance (in points) between the left edge of the page and the left boundary of the body text.|
||[lineNumbering](/javascript/api/word/word.pagesetup#word-word-pagesetup-linenumbering-member)|Specifies a `LineNumbering` object that represents the line numbers for the `PageSetup` object.|
||[linesPage](/javascript/api/word/word.pagesetup#word-word-pagesetup-linespage-member)|Specifies the number of lines per page in the document grid.|
||[mirrorMargins](/javascript/api/word/word.pagesetup#word-word-pagesetup-mirrormargins-member)|Specifies if the inside and outside margins of facing pages are the same width.|
||[oddAndEvenPagesHeaderFooter](/javascript/api/word/word.pagesetup#word-word-pagesetup-oddandevenpagesheaderfooter-member)|Specifies whether odd and even pages have different headers and footers.|
||[orientation](/javascript/api/word/word.pagesetup#word-word-pagesetup-orientation-member)|Specifies the orientation of the page.|
||[pageHeight](/javascript/api/word/word.pagesetup#word-word-pagesetup-pageheight-member)|Specifies the page height in points.|
||[pageWidth](/javascript/api/word/word.pagesetup#word-word-pagesetup-pagewidth-member)|Specifies the page width in points.|
||[paperSize](/javascript/api/word/word.pagesetup#word-word-pagesetup-papersize-member)|Specifies the paper size of the page.|
||[rightMargin](/javascript/api/word/word.pagesetup#word-word-pagesetup-rightmargin-member)|Specifies the distance (in points) between the right edge of the page and the right boundary of the body text.|
||[sectionDirection](/javascript/api/word/word.pagesetup#word-word-pagesetup-sectiondirection-member)|Specifies the reading order and alignment for the specified sections.|
||[sectionStart](/javascript/api/word/word.pagesetup#word-word-pagesetup-sectionstart-member)|Specifies the type of section break for the specified object.|
||[setAsTemplateDefault()](/javascript/api/word/word.pagesetup#word-word-pagesetup-setastemplatedefault-member(1))|Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.|
||[showGrid](/javascript/api/word/word.pagesetup#word-word-pagesetup-showgrid-member)|Specifies whether to show the grid.|
||[suppressEndnotes](/javascript/api/word/word.pagesetup#word-word-pagesetup-suppressendnotes-member)|Specifies if endnotes are printed at the end of the next section that doesn't suppress endnotes.|
||[textColumns](/javascript/api/word/word.pagesetup#word-word-pagesetup-textcolumns-member)|Gets a `TextColumnCollection` object that represents the set of text columns for the `PageSetup` object.|
||[togglePortrait()](/javascript/api/word/word.pagesetup#word-word-pagesetup-toggleportrait-member(1))|Switches between portrait and landscape page orientations for a document or section.|
||[topMargin](/javascript/api/word/word.pagesetup#word-word-pagesetup-topmargin-member)|Specifies the top margin of the page in points.|
||[twoPagesOnOne](/javascript/api/word/word.pagesetup#word-word-pagesetup-twopagesonone-member)|Specifies whether to print two pages per sheet.|
||[verticalAlignment](/javascript/api/word/word.pagesetup#word-word-pagesetup-verticalalignment-member)|Specifies the vertical alignment of text on each page in a document or section.|
|[Paragraph](/javascript/api/word/word.paragraph)|[borders](/javascript/api/word/word.paragraph#word-word-paragraph-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the paragraph.|
||[closeUp()](/javascript/api/word/word.paragraph#word-word-paragraph-closeup-member(1))|Removes any spacing before the paragraph.|
||[indent()](/javascript/api/word/word.paragraph#word-word-paragraph-indent-member(1))|Indents the paragraph by one level.|
||[indentCharacterWidth(count: number)](/javascript/api/word/word.paragraph#word-word-paragraph-indentcharacterwidth-member(1))|Indents the paragraph by a specified number of characters.|
||[indentFirstLineCharacterWidth(count: number)](/javascript/api/word/word.paragraph#word-word-paragraph-indentfirstlinecharacterwidth-member(1))|Indents the first line of the paragraph by the specified number of characters.|
||[joinList()](/javascript/api/word/word.paragraph#word-word-paragraph-joinlist-member(1))|Joins a list paragraph with the closest list above or below this paragraph.|
||[next(count: number)](/javascript/api/word/word.paragraph#word-word-paragraph-next-member(1))|Returns a `Paragraph` object that represents the next paragraph.|
||[onCommentAdded](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentselected-member)|Occurs when a comment is selected.|
||[openOrCloseUp()](/javascript/api/word/word.paragraph#word-word-paragraph-openorcloseup-member(1))|Toggles the spacing before the paragraph.|
||[openUp()](/javascript/api/word/word.paragraph#word-word-paragraph-openup-member(1))|Sets spacing before the paragraph to 12 points.|
||[outdent()](/javascript/api/word/word.paragraph#word-word-paragraph-outdent-member(1))|Removes one level of indent for the paragraph.|
||[outlineDemote()](/javascript/api/word/word.paragraph#word-word-paragraph-outlinedemote-member(1))|Applies the next heading level style (Heading 1 through Heading 8) to the paragraph.|
||[outlineDemoteToBody()](/javascript/api/word/word.paragraph#word-word-paragraph-outlinedemotetobody-member(1))|Demotes the paragraph to body text by applying the Normal style.|
||[outlinePromote()](/javascript/api/word/word.paragraph#word-word-paragraph-outlinepromote-member(1))|Applies the previous heading level style (Heading 1 through Heading 8) to the paragraph.|
||[previous(count: number)](/javascript/api/word/word.paragraph#word-word-paragraph-previous-member(1))|Returns the previous paragraph as a `Paragraph` object.|
||[range](/javascript/api/word/word.paragraph#word-word-paragraph-range-member)|Gets a `Range` object that represents the portion of the document that's contained within the paragraph.|
||[reset()](/javascript/api/word/word.paragraph#word-word-paragraph-reset-member(1))|Removes manual paragraph formatting (formatting not applied using a style).|
||[resetAdvanceTo()](/javascript/api/word/word.paragraph#word-word-paragraph-resetadvanceto-member(1))|Resets the paragraph that uses custom list levels to the original level settings.|
||[selectNumber()](/javascript/api/word/word.paragraph#word-word-paragraph-selectnumber-member(1))|Selects the number or bullet in a list.|
||[separateList()](/javascript/api/word/word.paragraph#word-word-paragraph-separatelist-member(1))|Separates a list into two separate lists.|
||[shading](/javascript/api/word/word.paragraph#word-word-paragraph-shading-member)|Returns a `ShadingUniversal` object that refers to the shading formatting for the paragraph.|
||[space1()](/javascript/api/word/word.paragraph#word-word-paragraph-space1-member(1))|Sets the paragraph to single spacing.|
||[space1Pt5()](/javascript/api/word/word.paragraph#word-word-paragraph-space1pt5-member(1))|Sets the paragraph to 1.5-line spacing.|
||[space2()](/javascript/api/word/word.paragraph#word-word-paragraph-space2-member(1))|Sets the paragraph to double spacing.|
||[tabHangingIndent(count: number)](/javascript/api/word/word.paragraph#word-word-paragraph-tabhangingindent-member(1))|Sets a hanging indent to a specified number of tab stops.|
||[tabIndent(count: number)](/javascript/api/word/word.paragraph#word-word-paragraph-tabindent-member(1))|Sets the left indent for the paragraph to a specified number of tab stops.|
|[ParagraphAddedEventArgs](/javascript/api/word/word.paragraphaddedeventargs)|[type](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-type-member)|The event type.|
|[ParagraphChangedEventArgs](/javascript/api/word/word.paragraphchangedeventargs)|[type](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-type-member)|The event type.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[add(range: Word.Range)](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-add-member(1))|Returns a `Paragraph` object that represents a new, blank paragraph added to the document.|
||[closeUp()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-closeup-member(1))|Removes any spacing before the specified paragraphs.|
||[decreaseSpacing()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-decreasespacing-member(1))|Decreases the spacing before and after paragraphs in six-point increments.|
||[increaseSpacing()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-increasespacing-member(1))|Increases the spacing before and after paragraphs in six-point increments.|
||[indent()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-indent-member(1))|Indents the paragraphs by one level.|
||[indentCharacterWidth(count: number)](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-indentcharacterwidth-member(1))|Indents the paragraphs in the collection by the specified number of characters.|
||[indentFirstLineCharacterWidth(count: number)](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-indentfirstlinecharacterwidth-member(1))|Indents the first line of the paragraphs in the collection by the specified number of characters.|
||[openOrCloseUp()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-openorcloseup-member(1))|Toggles spacing before paragraphs.|
||[openUp()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-openup-member(1))|Sets spacing before the specified paragraphs to 12 points.|
||[outdent()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-outdent-member(1))|Removes one level of indent for the paragraphs.|
||[outlineDemote()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-outlinedemote-member(1))|Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraphs.|
||[outlineDemoteToBody()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-outlinedemotetobody-member(1))|Demotes the specified paragraphs to body text by applying the Normal style.|
||[outlinePromote()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-outlinepromote-member(1))|Applies the previous heading level style (Heading 1 through Heading 8) to the paragraphs in the collection.|
||[space1()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-space1-member(1))|Sets the specified paragraphs to single spacing.|
||[space1Pt5()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-space1pt5-member(1))|Sets the specified paragraphs to 1.5-line spacing.|
||[space2()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-space2-member(1))|Sets the specified paragraphs to double spacing.|
||[tabHangingIndent(count: number)](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-tabhangingindent-member(1))|Sets a hanging indent to the specified number of tab stops.|
||[tabIndent(count: number)](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-tabindent-member(1))|Sets the left indent for the specified paragraphs to the specified number of tab stops.|
|[ParagraphDeletedEventArgs](/javascript/api/word/word.paragraphdeletedeventargs)|[type](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-type-member)|The event type.|
|[PictureContentControl](/javascript/api/word/word.picturecontentcontrol)|[appearance](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[color](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-color-member)|Specifies the red-green-blue (RGB) value of the color of the content control.|
||[copy()](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-copy-member(1))|Copies the content control from the active document to the Clipboard.|
||[cut()](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-cut-member(1))|Removes the content control from the active document and moves the content control to the Clipboard.|
||[delete(deleteContents?: boolean)](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-delete-member(1))|Deletes the content control and optionally its contents.|
||[id](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-id-member)|Returns the identification for the content control.|
||[isTemporary](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-istemporary-member)|Specifies whether to remove the content control from the active document when the user edits the contents of the control.|
||[level](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-level-member)|Returns the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.|
||[lockContentControl](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-lockcontentcontrol-member)|Specifies if the content control is locked (can't be deleted).|
||[lockContents](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-lockcontents-member)|Specifies if the contents of the content control are locked (not editable).|
||[placeholderText](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-placeholdertext-member)|Returns a `BuildingBlock` object that represents the placeholder text for the content control.|
||[range](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-range-member)|Returns a `Range` object that represents the contents of the content control in the active document.|
||[setPlaceholderText(options?: Word.ContentControlPlaceholderOptions)](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-setplaceholdertext-member(1))|Sets the placeholder text that displays in the content control until a user enters their own text.|
||[showingPlaceholderText](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-showingplaceholdertext-member)|Returns whether the placeholder text for the content control is being displayed.|
||[tag](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-tag-member)|Specifies a tag to identify the content control.|
||[title](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-title-member)|Specifies the title for the content control.|
||[xmlMapping](/javascript/api/word/word.picturecontentcontrol#word-word-picturecontentcontrol-xmlmapping-member)|Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[Range](/javascript/api/word/word.range)|[bold](/javascript/api/word/word.range#word-word-range-bold-member)|Specifies whether the range is formatted as bold.|
||[boldBidirectional](/javascript/api/word/word.range#word-word-range-boldbidirectional-member)|Specifies whether the range is formatted as bold in a right-to-left language document.|
||[bookmarks](/javascript/api/word/word.range#word-word-range-bookmarks-member)|Returns a `BookmarkCollection` object that represents all the bookmarks in the range.|
||[borders](/javascript/api/word/word.range#word-word-range-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the range.|
||[case](/javascript/api/word/word.range#word-word-range-case-member)|Specifies a `CharacterCase` value that represents the case of the text in the range.|
||[characterWidth](/javascript/api/word/word.range#word-word-range-characterwidth-member)|Specifies the character width of the range.|
||[combineCharacters](/javascript/api/word/word.range#word-word-range-combinecharacters-member)|Specifies if the range contains combined characters.|
||[detectLanguage()](/javascript/api/word/word.range#word-word-range-detectlanguage-member(1))|Analyzes the range text to determine the language that it's written in.|
||[disableCharacterSpaceGrid](/javascript/api/word/word.range#word-word-range-disablecharacterspacegrid-member)|Specifies if Microsoft Word ignores the number of characters per line for the corresponding `Range` object.|
||[emphasisMark](/javascript/api/word/word.range#word-word-range-emphasismark-member)|Specifies the emphasis mark for a character or designated character string.|
||[end](/javascript/api/word/word.range#word-word-range-end-member)|Specifies the ending character position of the range.|
||[fitTextWidth](/javascript/api/word/word.range#word-word-range-fittextwidth-member)|Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range.|
||[frames](/javascript/api/word/word.range#word-word-range-frames-member)|Gets a `FrameCollection` object that represents all the frames in the range.|
||[grammarChecked](/javascript/api/word/word.range#word-word-range-grammarchecked-member)|Specifies if a grammar check has been run on the range or document.|
||[hasNoProofing](/javascript/api/word/word.range#word-word-range-hasnoproofing-member)|Specifies the proofing status (spelling and grammar checking) of the range.|
||[highlightColorIndex](/javascript/api/word/word.range#word-word-range-highlightcolorindex-member)|Specifies the highlight color for the range.|
||[horizontalInVertical](/javascript/api/word/word.range#word-word-range-horizontalinvertical-member)|Specifies the formatting for horizontal text set within vertical text.|
||[hyperlinks](/javascript/api/word/word.range#word-word-range-hyperlinks-member)|Returns a `HyperlinkCollection` object that represents all the hyperlinks in the range.|
||[id](/javascript/api/word/word.range#word-word-range-id-member)|Specifies the ID for the range.|
||[isEndOfRowMark](/javascript/api/word/word.range#word-word-range-isendofrowmark-member)|Gets if the range is collapsed and is located at the end-of-row mark in a table.|
||[isTextVisibleOnScreen](/javascript/api/word/word.range#word-word-range-istextvisibleonscreen-member)|Gets whether the text in the range is visible on the screen.|
||[italic](/javascript/api/word/word.range#word-word-range-italic-member)|Specifies if the font or range is formatted as italic.|
||[italicBidirectional](/javascript/api/word/word.range#word-word-range-italicbidirectional-member)|Specifies if the font or range is formatted as italic (right-to-left languages).|
||[kana](/javascript/api/word/word.range#word-word-range-kana-member)|Specifies whether the range of Japanese language text is hiragana or katakana.|
||[languageDetected](/javascript/api/word/word.range#word-word-range-languagedetected-member)|Specifies whether Microsoft Word has detected the language of the text in the range.|
||[languageId](/javascript/api/word/word.range#word-word-range-languageid-member)|Specifies a `LanguageId` value that represents the language for the range.|
||[languageIdFarEast](/javascript/api/word/word.range#word-word-range-languageidfareast-member)|Specifies an East Asian language for the range.|
||[languageIdOther](/javascript/api/word/word.range#word-word-range-languageidother-member)|Specifies a language for the range that isn't classified as an East Asian language.|
||[listFormat](/javascript/api/word/word.range#word-word-range-listformat-member)|Returns a `ListFormat` object that represents all the list formatting characteristics of the range.|
||[onCommentAdded](/javascript/api/word/word.range#word-word-range-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.range#word-word-range-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.range#word-word-range-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.range#word-word-range-oncommentselected-member)|Occurs when a comment is selected.|
||[sections](/javascript/api/word/word.range#word-word-range-sections-member)|Gets the collection of sections in the range.|
||[shading](/javascript/api/word/word.range#word-word-range-shading-member)|Returns a `ShadingUniversal` object that refers to the shading formatting for the range.|
||[showAll](/javascript/api/word/word.range#word-word-range-showall-member)|Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed.|
||[spellingChecked](/javascript/api/word/word.range#word-word-range-spellingchecked-member)|Specifies if spelling has been checked throughout the range or document.|
||[start](/javascript/api/word/word.range#word-word-range-start-member)|Specifies the starting character position of the range.|
||[storyLength](/javascript/api/word/word.range#word-word-range-storylength-member)|Gets the number of characters in the story that contains the range.|
||[storyType](/javascript/api/word/word.range#word-word-range-storytype-member)|Gets the story type for the range.|
||[tableColumns](/javascript/api/word/word.range#word-word-range-tablecolumns-member)|Gets a `TableColumnCollection` object that represents all the table columns in the range.|
||[twoLinesInOne](/javascript/api/word/word.range#word-word-range-twolinesinone-member)|Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.|
||[underline](/javascript/api/word/word.range#word-word-range-underline-member)|Specifies the type of underline applied to the range.|
|[ReflectionFormat](/javascript/api/word/word.reflectionformat)|[blur](/javascript/api/word/word.reflectionformat#word-word-reflectionformat-blur-member)|Specifies the degree of blur effect applied to the `ReflectionFormat` object as a value between 0.0 and 100.0.|
||[offset](/javascript/api/word/word.reflectionformat#word-word-reflectionformat-offset-member)|Specifies the amount of separation, in points, of the reflected image from the shape.|
||[size](/javascript/api/word/word.reflectionformat#word-word-reflectionformat-size-member)|Specifies the size of the reflection as a percentage of the reflected shape from 0 to 100.|
||[transparency](/javascript/api/word/word.reflectionformat#word-word-reflectionformat-transparency-member)|Specifies the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear).|
||[type](/javascript/api/word/word.reflectionformat#word-word-reflectionformat-type-member)|Specifies a `ReflectionType` value that represents the type and direction of the lighting for a shape reflection.|
|[RepeatingSectionContentControl](/javascript/api/word/word.repeatingsectioncontentcontrol)|[allowInsertDeleteSection](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-allowinsertdeletesection-member)|Specifies whether users can add or remove sections from this repeating section content control by using the user interface.|
||[appearance](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[color](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-color-member)|Specifies the red-green-blue (RGB) value of the color of the content control.|
||[copy()](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-copy-member(1))|Copies the content control from the active document to the Clipboard.|
||[cut()](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-cut-member(1))|Removes the content control from the active document and moves the content control to the Clipboard.|
||[delete(deleteContents?: boolean)](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-delete-member(1))|Deletes the content control and the contents of the content control.|
||[id](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-id-member)|Returns the identification for the content control.|
||[isTemporary](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-istemporary-member)|Specifies whether to remove the content control from the active document when the user edits the contents of the control.|
||[level](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-level-member)|Returns the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.|
||[lockContentControl](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-lockcontentcontrol-member)|Specifies if the content control is locked (can't be deleted).|
||[lockContents](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-lockcontents-member)|Specifies if the contents of the content control are locked (not editable).|
||[placeholderText](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-placeholdertext-member)|Returns a `BuildingBlock` object that represents the placeholder text for the content control.|
||[range](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-range-member)|Gets a `Range` object that represents the contents of the content control in the active document.|
||[repeatingSectionItemTitle](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-repeatingsectionitemtitle-member)|Specifies the name of the repeating section items used in the context menu associated with this repeating section content control.|
||[repeatingSectionItems](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-repeatingsectionitems-member)|Returns the collection of repeating section items in this repeating section content control.|
||[setPlaceholderText(options?: Word.ContentControlPlaceholderOptions)](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-setplaceholdertext-member(1))|Sets the placeholder text that displays in the content control until a user enters their own text.|
||[showingPlaceholderText](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-showingplaceholdertext-member)|Returns whether the placeholder text for the content control is being displayed.|
||[tag](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-tag-member)|Specifies a tag to identify the content control.|
||[title](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-title-member)|Specifies the title for the content control.|
||[xmlapping](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-xmlapping-member)|Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[RepeatingSectionItem](/javascript/api/word/word.repeatingsectionitem)|[delete()](/javascript/api/word/word.repeatingsectionitem#word-word-repeatingsectionitem-delete-member(1))|Deletes this `RepeatingSectionItem` object.|
||[insertItemAfter()](/javascript/api/word/word.repeatingsectionitem#word-word-repeatingsectionitem-insertitemafter-member(1))|Adds a repeating section item after this item and returns the new item.|
||[insertItemBefore()](/javascript/api/word/word.repeatingsectionitem#word-word-repeatingsectionitem-insertitembefore-member(1))|Adds a repeating section item before this item and returns the new item.|
||[range](/javascript/api/word/word.repeatingsectionitem#word-word-repeatingsectionitem-range-member)|Returns the range of this repeating section item, excluding the start and end tags.|
|[RepeatingSectionItemCollection](/javascript/api/word/word.repeatingsectionitemcollection)|[getItemAt(index: number)](/javascript/api/word/word.repeatingsectionitemcollection#word-word-repeatingsectionitemcollection-getitemat-member(1))|Returns an individual repeating section item.|
|[Reviewer](/javascript/api/word/word.reviewer)|[isVisible](/javascript/api/word/word.reviewer#word-word-reviewer-isvisible-member)|Specifies if the `Reviewer` object is visible.|
|[ReviewerCollection](/javascript/api/word/word.reviewercollection)|[getItem(index: number)](/javascript/api/word/word.reviewercollection#word-word-reviewercollection-getitem-member(1))|Returns a `Reviewer` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.reviewercollection#word-word-reviewercollection-items-member)|Gets the loaded child items in this collection.|
|[RevisionsFilter](/javascript/api/word/word.revisionsfilter)|[markup](/javascript/api/word/word.revisionsfilter#word-word-revisionsfilter-markup-member)|Specifies a `RevisionsMarkup` value that represents the extent of reviewer markup displayed in the document.|
||[reviewers](/javascript/api/word/word.revisionsfilter#word-word-revisionsfilter-reviewers-member)|Gets the `ReviewerCollection` object that represents the collection of reviewers of one or more documents.|
||[toggleShowAllReviewers()](/javascript/api/word/word.revisionsfilter#word-word-revisionsfilter-toggleshowallreviewers-member(1))|Shows or hides all revisions in the document that contain comments and tracked changes.|
||[view](/javascript/api/word/word.revisionsfilter#word-word-revisionsfilter-view-member)|Specifies a `RevisionsView` value that represents globally whether Word displays the original version of the document or the final version, which might have revisions and formatting changes applied.|
|[Section](/javascript/api/word/word.section)|[borders](/javascript/api/word/word.section#word-word-section-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders in the section.|
||[pageSetup](/javascript/api/word/word.section#word-word-section-pagesetup-member)|Returns a `PageSetup` object that's associated with the section.|
||[protectedForForms](/javascript/api/word/word.section#word-word-section-protectedforforms-member)|Specifies if the section is protected for forms.|
|[ShadingUniversal](/javascript/api/word/word.shadinguniversal)|[backgroundPatternColor](/javascript/api/word/word.shadinguniversal#word-word-shadinguniversal-backgroundpatterncolor-member)|Specifies the color that's applied to the background of the `ShadingUniversal` object.|
||[backgroundPatternColorIndex](/javascript/api/word/word.shadinguniversal#word-word-shadinguniversal-backgroundpatterncolorindex-member)|Specifies the color that's applied to the background of the `ShadingUniversal` object.|
||[foregroundPatternColor](/javascript/api/word/word.shadinguniversal#word-word-shadinguniversal-foregroundpatterncolor-member)|Specifies the color that's applied to the foreground of the `ShadingUniversal` object.|
||[foregroundPatternColorIndex](/javascript/api/word/word.shadinguniversal#word-word-shadinguniversal-foregroundpatterncolorindex-member)|Specifies the color that's applied to the foreground of the `ShadingUniversal` object.|
||[texture](/javascript/api/word/word.shadinguniversal#word-word-shadinguniversal-texture-member)|Specifies the shading texture of the object.|
|[ShadowFormat](/javascript/api/word/word.shadowformat)|[blur](/javascript/api/word/word.shadowformat#word-word-shadowformat-blur-member)|Specifies the blur level for a shadow format as a value between 0.0 and 100.0.|
||[foregroundColor](/javascript/api/word/word.shadowformat#word-word-shadowformat-foregroundcolor-member)|Returns a `ColorFormat` object that represents the foreground color for the fill, line, or shadow.|
||[incrementOffsetX(increment: number)](/javascript/api/word/word.shadowformat#word-word-shadowformat-incrementoffsetx-member(1))|Changes the horizontal offset of the shadow by the number of points.|
||[incrementOffsetY(increment: number)](/javascript/api/word/word.shadowformat#word-word-shadowformat-incrementoffsety-member(1))|Changes the vertical offset of the shadow by the specified number of points.|
||[isVisible](/javascript/api/word/word.shadowformat#word-word-shadowformat-isvisible-member)|Specifies whether the object or the formatting applied to it is visible.|
||[obscured](/javascript/api/word/word.shadowformat#word-word-shadowformat-obscured-member)|Specifies `true` if the shadow of the shape appears filled in and is obscured by the shape, even if the shape has no fill,|
||[offsetX](/javascript/api/word/word.shadowformat#word-word-shadowformat-offsetx-member)|Specifies the horizontal offset (in points) of the shadow from the shape.|
||[offsetY](/javascript/api/word/word.shadowformat#word-word-shadowformat-offsety-member)|Specifies the vertical offset (in points) of the shadow from the shape.|
||[rotateWithShape](/javascript/api/word/word.shadowformat#word-word-shadowformat-rotatewithshape-member)|Specifies whether to rotate the shadow when rotating the shape.|
||[size](/javascript/api/word/word.shadowformat#word-word-shadowformat-size-member)|Specifies the width of the shadow.|
||[style](/javascript/api/word/word.shadowformat#word-word-shadowformat-style-member)|Specifies the type of shadow formatting to apply to a shape.|
||[transparency](/javascript/api/word/word.shadowformat#word-word-shadowformat-transparency-member)|Specifies the degree of transparency of the shadow as a value between 0.0 (opaque) and 1.0 (clear).|
||[type](/javascript/api/word/word.shadowformat#word-word-shadowformat-type-member)|Specifies the shape shadow type.|
|[Source](/javascript/api/word/word.source)|[delete()](/javascript/api/word/word.source#word-word-source-delete-member(1))|Deletes the `Source` object.|
||[getFieldByName(name: string)](/javascript/api/word/word.source#word-word-source-getfieldbyname-member(1))|Returns the value of a field in the bibliography `Source` object.|
||[isCited](/javascript/api/word/word.source#word-word-source-iscited-member)|Gets if the `Source` object has been cited in the document.|
||[tag](/javascript/api/word/word.source#word-word-source-tag-member)|Gets the tag of the source.|
||[xml](/javascript/api/word/word.source#word-word-source-xml-member)|Gets the XML representation of the source.|
|[SourceCollection](/javascript/api/word/word.sourcecollection)|[add(xml: string)](/javascript/api/word/word.sourcecollection#word-word-sourcecollection-add-member(1))|Adds a new `Source` object to the collection.|
||[getItem(index: number)](/javascript/api/word/word.sourcecollection#word-word-sourcecollection-getitem-member(1))|Gets a `Source` by its index in the collection.|
||[items](/javascript/api/word/word.sourcecollection#word-word-sourcecollection-items-member)|Gets the loaded child items in this collection.|
|[Style](/javascript/api/word/word.style)|[automaticallyUpdate](/javascript/api/word/word.style#word-word-style-automaticallyupdate-member)|Specifies whether the style is automatically redefined based on the selection.|
||[description](/javascript/api/word/word.style#word-word-style-description-member)|Gets the description of the specified style.|
||[frame](/javascript/api/word/word.style#word-word-style-frame-member)|Returns a `Frame` object that represents the frame formatting for the style.|
||[hasProofing](/javascript/api/word/word.style#word-word-style-hasproofing-member)|Specifies whether the spelling and grammar checker ignores text formatted with this style.|
||[languageId](/javascript/api/word/word.style#word-word-style-languageid-member)|Specifies a `LanguageId` value that represents the language for the style.|
||[languageIdFarEast](/javascript/api/word/word.style#word-word-style-languageidfareast-member)|Specifies an East Asian language for the style.|
||[linkStyle](/javascript/api/word/word.style#word-word-style-linkstyle-member)|Specifies a link between a paragraph and a character style.|
||[linkToListTemplate(listTemplate: Word.ListTemplate)](/javascript/api/word/word.style#word-word-style-linktolisttemplate-member(1))|Links this style to a list template so that the style's formatting can be applied to lists.|
||[listLevelNumber](/javascript/api/word/word.style#word-word-style-listlevelnumber-member)|Returns the list level for the style.|
||[locked](/javascript/api/word/word.style#word-word-style-locked-member)|Specifies whether the style cannot be changed or edited.|
||[noSpaceBetweenParagraphsOfSameStyle](/javascript/api/word/word.style#word-word-style-nospacebetweenparagraphsofsamestyle-member)|Specifies whether to remove spacing between paragraphs that are formatted using the same style.|
|[TabStop](/javascript/api/word/word.tabstop)|[alignment](/javascript/api/word/word.tabstop#word-word-tabstop-alignment-member)|Gets a `TabAlignment` value that represents the alignment for the tab stop.|
||[clear()](/javascript/api/word/word.tabstop#word-word-tabstop-clear-member(1))|Removes this custom tab stop.|
||[customTab](/javascript/api/word/word.tabstop#word-word-tabstop-customtab-member)|Gets whether this tab stop is a custom tab stop.|
||[leader](/javascript/api/word/word.tabstop#word-word-tabstop-leader-member)|Gets a `TabLeader` value that represents the leader for this `TabStop` object.|
||[next](/javascript/api/word/word.tabstop#word-word-tabstop-next-member)|Gets the next tab stop in the collection.|
||[position](/javascript/api/word/word.tabstop#word-word-tabstop-position-member)|Gets the position of the tab stop relative to the left margin.|
||[previous](/javascript/api/word/word.tabstop#word-word-tabstop-previous-member)|Gets the previous tab stop in the collection.|
|[TabStopAddOptions](/javascript/api/word/word.tabstopaddoptions)|[alignment](/javascript/api/word/word.tabstopaddoptions#word-word-tabstopaddoptions-alignment-member)|If provided, specifies the alignment of the tab stop.|
||[leader](/javascript/api/word/word.tabstopaddoptions#word-word-tabstopaddoptions-leader-member)|If provided, specifies the leader character for the tab stop.|
|[TabStopCollection](/javascript/api/word/word.tabstopcollection)|[add(position: number, options?: Word.TabStopAddOptions)](/javascript/api/word/word.tabstopcollection#word-word-tabstopcollection-add-member(1))|Returns a `TabStop` object that represents a custom tab stop added to the paragraph.|
||[after(Position: number)](/javascript/api/word/word.tabstopcollection#word-word-tabstopcollection-after-member(1))|Returns the next `TabStop` object to the right of the specified position.|
||[before(Position: number)](/javascript/api/word/word.tabstopcollection#word-word-tabstopcollection-before-member(1))|Returns the next `TabStop` object to the left of the specified position.|
||[clearAll()](/javascript/api/word/word.tabstopcollection#word-word-tabstopcollection-clearall-member(1))|Clears all the custom tab stops from the paragraph.|
||[getItem(index: number)](/javascript/api/word/word.tabstopcollection#word-word-tabstopcollection-getitem-member(1))|Gets a `TabStop` object by its index in the collection.|
||[items](/javascript/api/word/word.tabstopcollection#word-word-tabstopcollection-items-member)|Gets the loaded child items in this collection.|
|[TableColumn](/javascript/api/word/word.tablecolumn)|[autoFit()](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-autofit-member(1))|Changes the width of the table column to accommodate the width of the text without changing the way text wraps in the cells.|
||[borders](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the table column.|
||[columnIndex](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-columnindex-member)|Returns the position of this column in a collection.|
||[delete()](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-delete-member(1))|Deletes the column.|
||[isFirst](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-isfirst-member)|Returns `true` if the column or row is the first one in the table; `false` otherwise.|
||[isLast](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-islast-member)|Returns `true` if the column or row is the last one in the table; `false` otherwise.|
||[nestingLevel](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-nestinglevel-member)|Returns the nesting level of the column.|
||[preferredWidth](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-preferredwidth-member)|Specifies the preferred width (in points or as a percentage of the window width) for the column.|
||[preferredWidthType](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-preferredwidthtype-member)|Specifies the preferred unit of measurement to use for the width of the table column.|
||[select()](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-select-member(1))|Selects the table column.|
||[setWidth(columnWidth: number, rulerStyle: Word.RulerStyle)](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-setwidth-member(1))|Sets the width of the column in a table.|
||[shading](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-shading-member)|Returns a `ShadingUniversal` object that refers to the shading formatting for the column.|
||[sort()](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-sort-member(1))|Sorts the table column.|
||[width](/javascript/api/word/word.tablecolumn#word-word-tablecolumn-width-member)|Specifies the width of the column, in points.|
|[TableColumnCollection](/javascript/api/word/word.tablecolumncollection)|[add(beforeColumn?: Word.TableColumn)](/javascript/api/word/word.tablecolumncollection#word-word-tablecolumncollection-add-member(1))|Returns a `TableColumn` object that represents a column added to a table.|
||[autoFit()](/javascript/api/word/word.tablecolumncollection#word-word-tablecolumncollection-autofit-member(1))|Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.|
||[delete()](/javascript/api/word/word.tablecolumncollection#word-word-tablecolumncollection-delete-member(1))|Deletes the specified columns.|
||[distributeWidth()](/javascript/api/word/word.tablecolumncollection#word-word-tablecolumncollection-distributewidth-member(1))|Adjusts the width of the specified columns so that they are equal.|
||[items](/javascript/api/word/word.tablecolumncollection#word-word-tablecolumncollection-items-member)|Gets the loaded child items in this collection.|
||[select()](/javascript/api/word/word.tablecolumncollection#word-word-tablecolumncollection-select-member(1))|Selects the specified table columns.|
||[setWidth(columnWidth: number, rulerStyle: Word.RulerStyle)](/javascript/api/word/word.tablecolumncollection#word-word-tablecolumncollection-setwidth-member(1))|Sets the width of columns in a table.|
|[Template](/javascript/api/word/word.template)|[buildingBlockEntries](/javascript/api/word/word.template#word-word-template-buildingblockentries-member)|Returns a `BuildingBlockEntryCollection` object that represents the collection of building block entries in the template.|
||[buildingBlockTypes](/javascript/api/word/word.template#word-word-template-buildingblocktypes-member)|Returns a `BuildingBlockTypeItemCollection` object that represents the collection of building block types that are contained in the template.|
||[farEastLineBreakLanguage](/javascript/api/word/word.template#word-word-template-fareastlinebreaklanguage-member)|Specifies the East Asian language to use when breaking lines of text in the document or template.|
||[farEastLineBreakLevel](/javascript/api/word/word.template#word-word-template-fareastlinebreaklevel-member)|Specifies the line break control level for the document.|
||[fullName](/javascript/api/word/word.template#word-word-template-fullname-member)|Returns the name of the template, including the drive or Web path.|
||[hasNoProofing](/javascript/api/word/word.template#word-word-template-hasnoproofing-member)|Specifies whether the spelling and grammar checker ignores documents based on this template.|
||[justificationMode](/javascript/api/word/word.template#word-word-template-justificationmode-member)|Specifies the character spacing adjustment for the template.|
||[kerningByAlgorithm](/javascript/api/word/word.template#word-word-template-kerningbyalgorithm-member)|Specifies if Microsoft Word kerns half-width Latin characters and punctuation marks in the document.|
||[languageId](/javascript/api/word/word.template#word-word-template-languageid-member)|Specifies a `LanguageId` value that represents the language in the template.|
||[languageIdFarEast](/javascript/api/word/word.template#word-word-template-languageidfareast-member)|Specifies an East Asian language for the language in the template.|
||[name](/javascript/api/word/word.template#word-word-template-name-member)|Returns only the name of the document template (excluding any path or other location information).|
||[noLineBreakAfter](/javascript/api/word/word.template#word-word-template-nolinebreakafter-member)|Specifies the kinsoku characters after which Microsoft Word will not break a line.|
||[noLineBreakBefore](/javascript/api/word/word.template#word-word-template-nolinebreakbefore-member)|Specifies the kinsoku characters before which Microsoft Word will not break a line.|
||[path](/javascript/api/word/word.template#word-word-template-path-member)|Returns the path to the document template.|
||[save()](/javascript/api/word/word.template#word-word-template-save-member(1))|Saves the template.|
||[saved](/javascript/api/word/word.template#word-word-template-saved-member)|Specifies `true` if the template has not changed since it was last saved, `false` if Microsoft Word displays a prompt to save changes when the document is closed.|
||[type](/javascript/api/word/word.template#word-word-template-type-member)|Returns the template type.|
|[TemplateCollection](/javascript/api/word/word.templatecollection)|[getCount()](/javascript/api/word/word.templatecollection#word-word-templatecollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/javascript/api/word/word.templatecollection#word-word-templatecollection-getitemat-member(1))|Gets a `Template` object by its index in the collection.|
||[importBuildingBlocks()](/javascript/api/word/word.templatecollection#word-word-templatecollection-importbuildingblocks-member(1))|Imports the building blocks for all templates into Microsoft Word.|
||[items](/javascript/api/word/word.templatecollection#word-word-templatecollection-items-member)|Gets the loaded child items in this collection.|
|[TextColumn](/javascript/api/word/word.textcolumn)|[spaceAfter](/javascript/api/word/word.textcolumn#word-word-textcolumn-spaceafter-member)|Specifies the amount of spacing (in points) after the specified paragraph or text column.|
||[width](/javascript/api/word/word.textcolumn#word-word-textcolumn-width-member)|Specifies the width, in points, of the specified text columns.|
|[TextColumnAddOptions](/javascript/api/word/word.textcolumnaddoptions)|[isEvenlySpaced](/javascript/api/word/word.textcolumnaddoptions#word-word-textcolumnaddoptions-isevenlyspaced-member)|If provided, specifies whether to evenly space all the text columns in the document.|
||[spacing](/javascript/api/word/word.textcolumnaddoptions#word-word-textcolumnaddoptions-spacing-member)|If provided, specifies the spacing between the text columns in the document, in points.|
||[width](/javascript/api/word/word.textcolumnaddoptions#word-word-textcolumnaddoptions-width-member)|If provided, specifies the width of the new text column in the document, in points.|
|[TextColumnCollection](/javascript/api/word/word.textcolumncollection)|[add(options?: Word.TextColumnAddOptions)](/javascript/api/word/word.textcolumncollection#word-word-textcolumncollection-add-member(1))|Returns a `TextColumn` object that represents a new text column added to a section or document.|
||[getFlowDirection()](/javascript/api/word/word.textcolumncollection#word-word-textcolumncollection-getflowdirection-member(1))|Gets the direction in which text flows from one text column to the next.|
||[getHasLineBetween()](/javascript/api/word/word.textcolumncollection#word-word-textcolumncollection-gethaslinebetween-member(1))|Gets whether vertical lines appear between all the columns in the `TextColumnCollection` object.|
||[getIsEvenlySpaced()](/javascript/api/word/word.textcolumncollection#word-word-textcolumncollection-getisevenlyspaced-member(1))|Gets whether text columns are evenly spaced.|
||[getItem(index: number)](/javascript/api/word/word.textcolumncollection#word-word-textcolumncollection-getitem-member(1))|Gets a `TextColumn` by its index in the collection.|
||[items](/javascript/api/word/word.textcolumncollection#word-word-textcolumncollection-items-member)|Gets the loaded child items in this collection.|
||[setCount(numColumns: number)](/javascript/api/word/word.textcolumncollection#word-word-textcolumncollection-setcount-member(1))|Arranges text into the specified number of text columns.|
||[setFlowDirection(value: Word.FlowDirection)](/javascript/api/word/word.textcolumncollection#word-word-textcolumncollection-setflowdirection-member(1))|Sets the direction in which text flows from one text column to the next.|
||[setHasLineBetween(value: boolean)](/javascript/api/word/word.textcolumncollection#word-word-textcolumncollection-sethaslinebetween-member(1))|Sets whether vertical lines appear between all the columns in the `TextColumnCollection` object.|
||[setIsEvenlySpaced(value: boolean)](/javascript/api/word/word.textcolumncollection#word-word-textcolumncollection-setisevenlyspaced-member(1))|Sets whether text columns are evenly spaced.|
|[ThreeDimensionalFormat](/javascript/api/word/word.threedimensionalformat)|[bevelBottomDepth](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-bevelbottomdepth-member)|Specifies the depth of the bottom bevel.|
||[bevelBottomInset](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-bevelbottominset-member)|Specifies the inset size for the bottom bevel.|
||[bevelBottomType](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-bevelbottomtype-member)|Specifies a `BevelType` value that represents the bevel type for the bottom bevel.|
||[bevelTopDepth](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-beveltopdepth-member)|Specifies the depth of the top bevel.|
||[bevelTopInset](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-beveltopinset-member)|Specifies the inset size for the top bevel.|
||[bevelTopType](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-beveltoptype-member)|Specifies a `BevelType` value that represents the bevel type for the top bevel.|
||[contourColor](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-contourcolor-member)|Returns a `ColorFormat` object that represents color of the contour of a shape.|
||[contourWidth](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-contourwidth-member)|Specifies the width of the contour of a shape.|
||[depth](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-depth-member)|Specifies the depth of the shape's extrusion.|
||[extrusionColor](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-extrusioncolor-member)|Returns a `ColorFormat` object that represents the color of the shape's extrusion.|
||[extrusionColorType](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-extrusioncolortype-member)|Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion)|
||[fieldOfView](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-fieldofview-member)|Specifies the amount of perspective for a shape.|
||[incrementRotationHorizontal(increment: number)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-incrementrotationhorizontal-member(1))|Horizontally rotates a shape on the x-axis.|
||[incrementRotationVertical(increment: number)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-incrementrotationvertical-member(1))|Vertically rotates a shape on the y-axis.|
||[incrementRotationX(increment: number)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-incrementrotationx-member(1))|Changes the rotation around the x-axis.|
||[incrementRotationY(increment: number)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-incrementrotationy-member(1))|Changes the rotation around the y-axis.|
||[incrementRotationZ(increment: number)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-incrementrotationz-member(1))|Rotates a shape on the z-axis.|
||[isPerspective](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-isperspective-member)|Specifies `true` if the extrusion appears in perspective — that is, if the walls of the extrusion narrow toward a vanishing point,|
||[isVisible](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-isvisible-member)|Specifies if the specified object, or the formatting applied to it, is visible.|
||[lightAngle](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-lightangle-member)|Specifies the angle of the lighting.|
||[presetCamera](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetcamera-member)|Returns a `PresetCamera` value that represents the camera presets.|
||[presetExtrusionDirection](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetextrusiondirection-member)|Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion).|
||[presetLighting](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetlighting-member)|Specifies a `LightRigType` value that represents the lighting preset.|
||[presetLightingDirection](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetlightingdirection-member)|Specifies the position of the light source relative to the extrusion.|
||[presetLightingSoftness](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetlightingsoftness-member)|Specifies the intensity of the extrusion lighting.|
||[presetMaterial](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetmaterial-member)|Specifies the extrusion surface material.|
||[presetThreeDimensionalFormat](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetthreedimensionalformat-member)|Returns the preset extrusion format.|
||[projectText](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-projecttext-member)|Specifies whether text on a shape rotates with shape.|
||[resetRotation()](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-resetrotation-member(1))|Resets the extrusion rotation around the x-axis, y-axis, and z-axis to 0.|
||[rotationX](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-rotationx-member)|Specifies the rotation of the extruded shape around the x-axis in degrees.|
||[rotationY](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-rotationy-member)|Specifies the rotation of the extruded shape around the y-axis in degrees.|
||[rotationZ](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-rotationz-member)|Specifies the z-axis rotation of the camera.|
||[setExtrusionDirection(presetExtrusionDirection: Word.PresetExtrusionDirection)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-setextrusiondirection-member(1))|Sets the direction of the extrusion's sweep path.|
||[setPresetCamera(presetCamera: Word.PresetCamera)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-setpresetcamera-member(1))|Sets the camera preset for the shape.|
||[setThreeDimensionalFormat(presetThreeDimensionalFormat: Word.PresetThreeDimensionalFormat)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-setthreedimensionalformat-member(1))|Sets the preset extrusion format.|
||[z](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-z-member)|Specifies the position on the z-axis for the shape.|
|[View](/javascript/api/word/word.view)|[areAllNonprintingCharactersDisplayed](/javascript/api/word/word.view#word-word-view-areallnonprintingcharactersdisplayed-member)|Specifies whether all nonprinting characters are displayed.|
||[areBackgroundsDisplayed](/javascript/api/word/word.view#word-word-view-arebackgroundsdisplayed-member)|Gets whether background colors and images are shown when the document is displayed in print layout view.|
||[areBookmarksIndicated](/javascript/api/word/word.view#word-word-view-arebookmarksindicated-member)|Gets whether square brackets are displayed at the beginning and end of each bookmark.|
||[areCommentsDisplayed](/javascript/api/word/word.view#word-word-view-arecommentsdisplayed-member)|Specifies whether Microsoft Word displays the comments in the document.|
||[areConnectingLinesToRevisionsBalloonDisplayed](/javascript/api/word/word.view#word-word-view-areconnectinglinestorevisionsballoondisplayed-member)|Specifies whether Microsoft Word displays connecting lines from the text to the revision and comment balloons.|
||[areCropMarksDisplayed](/javascript/api/word/word.view#word-word-view-arecropmarksdisplayed-member)|Gets whether crop marks are shown in the corners of pages to indicate where margins are located.|
||[areDrawingsDisplayed](/javascript/api/word/word.view#word-word-view-aredrawingsdisplayed-member)|Gets whether objects created with the drawing tools are displayed in print layout view.|
||[areEditableRangesShaded](/javascript/api/word/word.view#word-word-view-areeditablerangesshaded-member)|Specifies whether shading is applied to the ranges in the document that users have permission to modify.|
||[areFieldCodesDisplayed](/javascript/api/word/word.view#word-word-view-arefieldcodesdisplayed-member)|Specifies whether field codes are displayed.|
||[areFormatChangesDisplayed](/javascript/api/word/word.view#word-word-view-areformatchangesdisplayed-member)|Specifies whether Microsoft Word displays formatting changes made to the document with Track Changes enabled.|
||[areInkAnnotationsDisplayed](/javascript/api/word/word.view#word-word-view-areinkannotationsdisplayed-member)|Specifies whether handwritten ink annotations are shown or hidden.|
||[areInsertionsAndDeletionsDisplayed](/javascript/api/word/word.view#word-word-view-areinsertionsanddeletionsdisplayed-member)|Specifies whether Microsoft Word displays insertions and deletions made to the document with Track Changes enabled.|
||[areLinesWrappedToWindow](/javascript/api/word/word.view#word-word-view-arelineswrappedtowindow-member)|Gets whether lines wrap at the right edge of the document window rather than at the right margin or the right column boundary.|
||[areObjectAnchorsDisplayed](/javascript/api/word/word.view#word-word-view-areobjectanchorsdisplayed-member)|Gets whether object anchors are displayed next to items that can be positioned in print layout view.|
||[areOptionalBreaksDisplayed](/javascript/api/word/word.view#word-word-view-areoptionalbreaksdisplayed-member)|Gets whether Microsoft Word displays optional line breaks.|
||[areOptionalHyphensDisplayed](/javascript/api/word/word.view#word-word-view-areoptionalhyphensdisplayed-member)|Gets whether optional hyphens are displayed.|
||[areOtherAuthorsVisible](/javascript/api/word/word.view#word-word-view-areotherauthorsvisible-member)|Gets whether other authors' presence should be visible in the document.|
||[arePageBoundariesDisplayed](/javascript/api/word/word.view#word-word-view-arepageboundariesdisplayed-member)|Gets whether the top and bottom margins and the gray area between pages in the document are displayed.|
||[areParagraphsMarksDisplayed](/javascript/api/word/word.view#word-word-view-areparagraphsmarksdisplayed-member)|Gets whether paragraph marks are displayed.|
||[arePicturePlaceholdersDisplayed](/javascript/api/word/word.view#word-word-view-arepictureplaceholdersdisplayed-member)|Gets whether blank boxes are displayed as placeholders for pictures.|
||[areRevisionsAndCommentsDisplayed](/javascript/api/word/word.view#word-word-view-arerevisionsandcommentsdisplayed-member)|Specifies whether Microsoft Word displays revisions and comments made to the document with Track Changes enabled.|
||[areSpacesIndicated](/javascript/api/word/word.view#word-word-view-arespacesindicated-member)|Gets whether space characters are displayed.|
||[areTableGridlinesDisplayed](/javascript/api/word/word.view#word-word-view-aretablegridlinesdisplayed-member)|Specifies whether table gridlines are displayed.|
||[areTabsDisplayed](/javascript/api/word/word.view#word-word-view-aretabsdisplayed-member)|Gets whether tab characters are displayed.|
||[areTextBoundariesDisplayed](/javascript/api/word/word.view#word-word-view-aretextboundariesdisplayed-member)|Gets whether dotted lines are displayed around page margins, text columns, objects, and frames in print layout view.|
||[collapseAllHeadings()](/javascript/api/word/word.view#word-word-view-collapseallheadings-member(1))|Collapses all the headings in the document.|
||[collapseOutline(range: Word.Range)](/javascript/api/word/word.view#word-word-view-collapseoutline-member(1))|Collapses the text under the selection or the specified range by one heading level.|
||[columnWidth](/javascript/api/word/word.view#word-word-view-columnwidth-member)|Specifies the column width in Reading mode.|
||[expandAllHeadings()](/javascript/api/word/word.view#word-word-view-expandallheadings-member(1))|Expands all the headings in the document.|
||[expandOutline(range: Word.Range)](/javascript/api/word/word.view#word-word-view-expandoutline-member(1))|Expands the text under the selection by one heading level.|
||[fieldShading](/javascript/api/word/word.view#word-word-view-fieldshading-member)|Gets on-screen shading for fields.|
||[isDraft](/javascript/api/word/word.view#word-word-view-isdraft-member)|Specifies whether all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.|
||[isFirstLineOnlyDisplayed](/javascript/api/word/word.view#word-word-view-isfirstlineonlydisplayed-member)|Specifies whether only the first line of body text is shown in outline view.|
||[isFormatDisplayed](/javascript/api/word/word.view#word-word-view-isformatdisplayed-member)|Specifies whether character formatting is visible in outline view.|
||[isFullScreen](/javascript/api/word/word.view#word-word-view-isfullscreen-member)|Specifies whether the window is in full-screen view.|
||[isHiddenTextDisplayed](/javascript/api/word/word.view#word-word-view-ishiddentextdisplayed-member)|Gets whether text formatted as hidden text is displayed.|
||[isHighlightingDisplayed](/javascript/api/word/word.view#word-word-view-ishighlightingdisplayed-member)|Gets whether highlight formatting is displayed and printed with the document.|
||[isInConflictMode](/javascript/api/word/word.view#word-word-view-isinconflictmode-member)|Specifies whether the document is in conflict mode view.|
||[isInPanning](/javascript/api/word/word.view#word-word-view-isinpanning-member)|Specifies whether Microsoft Word is in Panning mode.|
||[isInReadingLayout](/javascript/api/word/word.view#word-word-view-isinreadinglayout-member)|Specifies whether the document is being viewed in reading layout view.|
||[isMailMergeDataView](/javascript/api/word/word.view#word-word-view-ismailmergedataview-member)|Specifies whether mail merge data is displayed instead of mail merge fields.|
||[isMainTextLayerVisible](/javascript/api/word/word.view#word-word-view-ismaintextlayervisible-member)|Specifies whether the text in the document is visible when the header and footer areas are displayed.|
||[isPointerShownAsMagnifier](/javascript/api/word/word.view#word-word-view-ispointershownasmagnifier-member)|Specifies whether the pointer is displayed as a magnifying glass in print preview.|
||[isReadingLayoutActualView](/javascript/api/word/word.view#word-word-view-isreadinglayoutactualview-member)|Specifies whether pages displayed in reading layout view are displayed using the same layout as printed pages.|
||[isXmlMarkupVisible](/javascript/api/word/word.view#word-word-view-isxmlmarkupvisible-member)|Specifies whether XML tags are visible in the document.|
||[markupMode](/javascript/api/word/word.view#word-word-view-markupmode-member)|Specifies the display mode for tracked changes.|
||[nextHeaderFooter()](/javascript/api/word/word.view#word-word-view-nextheaderfooter-member(1))|Moves to the next header or footer, depending on whether a header or footer is displayed in the view.|
||[pageColor](/javascript/api/word/word.view#word-word-view-pagecolor-member)|Specifies the page color in Reading mode.|
||[pageMovementType](/javascript/api/word/word.view#word-word-view-pagemovementtype-member)|Specifies the page movement type.|
||[previousHeaderFooter()](/javascript/api/word/word.view#word-word-view-previousheaderfooter-member(1))|Moves to the previous header or footer, depending on whether a header or footer is displayed in the view.|
||[readingLayoutTruncateMargins](/javascript/api/word/word.view#word-word-view-readinglayouttruncatemargins-member)|Specifies whether margins are visible or hidden when the document is viewed in Full Screen Reading view.|
||[revisionsBalloonSide](/javascript/api/word/word.view#word-word-view-revisionsballoonside-member)|Gets whether Word displays revision balloons in the left or right margin in the document.|
||[revisionsBalloonWidth](/javascript/api/word/word.view#word-word-view-revisionsballoonwidth-member)|Specifies the width of the revision balloons.|
||[revisionsBalloonWidthType](/javascript/api/word/word.view#word-word-view-revisionsballoonwidthtype-member)|Specifies how Microsoft Word measures the width of revision balloons.|
||[revisionsFilter](/javascript/api/word/word.view#word-word-view-revisionsfilter-member)|Gets the instance of a `RevisionsFilter` object.|
||[seekView](/javascript/api/word/word.view#word-word-view-seekview-member)|Specifies the document element displayed in print layout view.|
||[showAllHeadings()](/javascript/api/word/word.view#word-word-view-showallheadings-member(1))|Switches between showing all text (headings and body text) and showing only headings.|
||[showHeading(level: number)](/javascript/api/word/word.view#word-word-view-showheading-member(1))|Shows all headings up to the specified heading level and hides subordinate headings and body text.|
||[splitSpecial](/javascript/api/word/word.view#word-word-view-splitspecial-member)|Specifies the active window pane.|
||[type](/javascript/api/word/word.view#word-word-view-type-member)|Specifies the view type.|
|[Window](/javascript/api/word/word.window)|[activate()](/javascript/api/word/word.window#word-word-window-activate-member(1))|Activates the window.|
||[areRulersDisplayed](/javascript/api/word/word.window#word-word-window-arerulersdisplayed-member)|Specifies whether rulers are displayed for the window or pane.|
||[areScreenTipsDisplayed](/javascript/api/word/word.window#word-word-window-arescreentipsdisplayed-member)|Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.|
||[areThumbnailsDisplayed](/javascript/api/word/word.window#word-word-window-arethumbnailsdisplayed-member)|Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.|
||[caption](/javascript/api/word/word.window#word-word-window-caption-member)|Specifies the caption text for the window that is displayed in the title bar of the document or application window.|
||[close(options?: Word.WindowCloseOptions)](/javascript/api/word/word.window#word-word-window-close-member(1))|Closes the window.|
||[height](/javascript/api/word/word.window#word-word-window-height-member)|Specifies the height of the window (in points).|
||[horizontalPercentScrolled](/javascript/api/word/word.window#word-word-window-horizontalpercentscrolled-member)|Specifies the horizontal scroll position as a percentage of the document width.|
||[imemode](/javascript/api/word/word.window#word-word-window-imemode-member)|Specifies the default start-up mode for the Japanese Input Method Editor (IME).|
||[index](/javascript/api/word/word.window#word-word-window-index-member)|Gets the position of an item in a collection.|
||[isActive](/javascript/api/word/word.window#word-word-window-isactive-member)|Specifies whether the window is active.|
||[isDocumentMapVisible](/javascript/api/word/word.window#word-word-window-isdocumentmapvisible-member)|Specifies whether the document map is visible.|
||[isEnvelopeVisible](/javascript/api/word/word.window#word-word-window-isenvelopevisible-member)|Specifies whether the email message header is visible in the document window.|
||[isHorizontalScrollBarDisplayed](/javascript/api/word/word.window#word-word-window-ishorizontalscrollbardisplayed-member)|Specifies whether a horizontal scroll bar is displayed for the window.|
||[isLeftScrollBarDisplayed](/javascript/api/word/word.window#word-word-window-isleftscrollbardisplayed-member)|Specifies whether the vertical scroll bar appears on the left side of the document window.|
||[isRightRulerDisplayed](/javascript/api/word/word.window#word-word-window-isrightrulerdisplayed-member)|Specifies whether the vertical ruler appears on the right side of the document window in print layout view.|
||[isSplit](/javascript/api/word/word.window#word-word-window-issplit-member)|Specifies whether the window is split into multiple panes.|
||[isVerticalRulerDisplayed](/javascript/api/word/word.window#word-word-window-isverticalrulerdisplayed-member)|Specifies whether a vertical ruler is displayed for the window or pane.|
||[isVerticalScrollBarDisplayed](/javascript/api/word/word.window#word-word-window-isverticalscrollbardisplayed-member)|Specifies whether a vertical scroll bar is displayed for the window.|
||[isVisible](/javascript/api/word/word.window#word-word-window-isvisible-member)|Specifies whether the window is visible.|
||[largeScroll(options?: Word.WindowScrollOptions)](/javascript/api/word/word.window#word-word-window-largescroll-member(1))|Scrolls the window by the specified number of screens.|
||[left](/javascript/api/word/word.window#word-word-window-left-member)|Specifies the horizontal position of the window, measured in points.|
||[next](/javascript/api/word/word.window#word-word-window-next-member)|Gets the next document window in the collection of open document windows.|
||[pageScroll(options?: Word.WindowPageScrollOptions)](/javascript/api/word/word.window#word-word-window-pagescroll-member(1))|Scrolls through the window page by page.|
||[previous](/javascript/api/word/word.window#word-word-window-previous-member)|Gets the previous document window in the collection open document windows.|
||[setFocus()](/javascript/api/word/word.window#word-word-window-setfocus-member(1))|Sets the focus of the document window to the body of an email message.|
||[showSourceDocuments](/javascript/api/word/word.window#word-word-window-showsourcedocuments-member)|Specifies how Microsoft Word displays source documents after a compare and merge process.|
||[smallScroll(options?: Word.WindowScrollOptions)](/javascript/api/word/word.window#word-word-window-smallscroll-member(1))|Scrolls the window by the specified number of lines.|
||[splitVertical](/javascript/api/word/word.window#word-word-window-splitvertical-member)|Specifies the vertical split percentage for the window.|
||[styleAreaWidth](/javascript/api/word/word.window#word-word-window-styleareawidth-member)|Specifies the width of the style area in points.|
||[toggleRibbon()](/javascript/api/word/word.window#word-word-window-toggleribbon-member(1))|Shows or hides the ribbon.|
||[top](/javascript/api/word/word.window#word-word-window-top-member)|Specifies the vertical position of the document window, in points.|
||[type](/javascript/api/word/word.window#word-word-window-type-member)|Gets the window type.|
||[usableHeight](/javascript/api/word/word.window#word-word-window-usableheight-member)|Gets the height (in points) of the active working area in the document window.|
||[usableWidth](/javascript/api/word/word.window#word-word-window-usablewidth-member)|Gets the width (in points) of the active working area in the document window.|
||[verticalPercentScrolled](/javascript/api/word/word.window#word-word-window-verticalpercentscrolled-member)|Specifies the vertical scroll position as a percentage of the document length.|
||[view](/javascript/api/word/word.window#word-word-window-view-member)|Gets the `View` object that represents the view for the window.|
||[width](/javascript/api/word/word.window#word-word-window-width-member)|Specifies the width of the document window, in points.|
||[windowNumber](/javascript/api/word/word.window#word-word-window-windownumber-member)|Gets an integer that represents the position of the window.|
||[windowState](/javascript/api/word/word.window#word-word-window-windowstate-member)|Specifies the state of the document window or task window.|
|[WindowCloseOptions](/javascript/api/word/word.windowcloseoptions)|[routeDocument](/javascript/api/word/word.windowcloseoptions#word-word-windowcloseoptions-routedocument-member)|If provided, specifies whether to route the document to the next recipient.|
||[saveChanges](/javascript/api/word/word.windowcloseoptions#word-word-windowcloseoptions-savechanges-member)|If provided, specifies the save action for the document.|
|[WindowCollection](/javascript/api/word/word.windowcollection)|||
|[WindowPageScrollOptions](/javascript/api/word/word.windowpagescrolloptions)|[down](/javascript/api/word/word.windowpagescrolloptions#word-word-windowpagescrolloptions-down-member)|If provided, specifies the number of pages to scroll the window down.|
||[up](/javascript/api/word/word.windowpagescrolloptions#word-word-windowpagescrolloptions-up-member)|If provided, specifies the number of pages to scroll the window up.|
|[WindowScrollOptions](/javascript/api/word/word.windowscrolloptions)|[down](/javascript/api/word/word.windowscrolloptions#word-word-windowscrolloptions-down-member)|If provided, specifies the number of units to scroll the window down.|
||[left](/javascript/api/word/word.windowscrolloptions#word-word-windowscrolloptions-left-member)|If provided, specifies the number of screens to scroll the window to the left.|
||[right](/javascript/api/word/word.windowscrolloptions#word-word-windowscrolloptions-right-member)|If provided, specifies the number of screens to scroll the window to the right.|
||[up](/javascript/api/word/word.windowscrolloptions#word-word-windowscrolloptions-up-member)|If provided, specifies the number of units to scroll the window up.|
|[XmlMapping](/javascript/api/word/word.xmlmapping)|[customXmlNode](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-customxmlnode-member)|Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.|
||[customXmlPart](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-customxmlpart-member)|Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.|
||[delete()](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-delete-member(1))|Deletes the XML mapping from the parent content control.|
||[isMapped](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-ismapped-member)|Returns whether the content control in the document is mapped to an XML node in the document's XML data store.|
||[prefixMappings](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-prefixmappings-member)|Returns the prefix mappings used to evaluate the XPath for the current XML mapping.|
||[setMapping(xPath: string, options?: Word.XmlSetMappingOptions)](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-setmapping-member(1))|Allows creating or changing the XML mapping on the content control.|
||[setMappingByNode(node: Word.CustomXmlNode)](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-setmappingbynode-member(1))|Allows creating or changing the XML data mapping on the content control.|
||[xpath](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-xpath-member)|Returns the XPath for the XML mapping, which evaluates to the currently mapped XML node.|
|[XmlSetMappingOptions](/javascript/api/word/word.xmlsetmappingoptions)|[prefixMapping](/javascript/api/word/word.xmlsetmappingoptions#word-word-xmlsetmappingoptions-prefixmapping-member)|If provided, specifies the prefix mappings to use when querying the expression provided in the `xPath` parameter of the `XmlMapping.setMapping` calling method.|
||[source](/javascript/api/word/word.xmlsetmappingoptions#word-word-xmlsetmappingoptions-source-member)|If provided, specifies the desired custom XML data to map the content control to.|
