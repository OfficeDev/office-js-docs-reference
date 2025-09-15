| Class | Fields | Description |
|:---|:---|:---|
|[Application](/.application)|[bibliography](/.application#word-javascript/api/word/-application-bibliography-member)|Returns a `Bibliography` object that represents the bibliography reference sources stored in Microsoft Word.|
||[checkLanguage](/.application#word-javascript/api/word/-application-checklanguage-member)|Specifies if Microsoft Word automatically detects the language you are using as you type.|
||[language](/.application#word-javascript/api/word/-application-language-member)|Gets a `LanguageId` value that represents the language selected for the Microsoft Word user interface.|
||[templates](/.application#word-javascript/api/word/-application-templates-member)|Returns a `TemplateCollection` object that represents all the available templates: global templates and those attached to open documents.|
|[Bibliography](/.bibliography)|[bibliographyStyle](/.bibliography#word-javascript/api/word/-bibliography-bibliographystyle-member)|Specifies the name of the active style to use for the bibliography.|
||[generateUniqueTag()](/.bibliography#word-javascript/api/word/-bibliography-generateuniquetag-member(1))|Generates a unique identification tag for a bibliography source and returns a string that represents the tag.|
||[sources](/.bibliography#word-javascript/api/word/-bibliography-sources-member)|Returns a `SourceCollection` object that represents all the sources contained in the bibliography.|
|[Body](/.body)|[onCommentAdded](/.body#word-javascript/api/word/-body-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/.body#word-javascript/api/word/-body-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/.body#word-javascript/api/word/-body-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/.body#word-javascript/api/word/-body-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/.body#word-javascript/api/word/-body-oncommentselected-member)|Occurs when a comment is selected.|
||[type](/.body#word-javascript/api/word/-body-type-member)|Gets the type of the body.|
|[Bookmark](/.bookmark)|[copyTo(name: string)](/.bookmark#word-javascript/api/word/-bookmark-copyto-member(1))|Copies this bookmark to the new bookmark specified in the `name` argument and returns a `Bookmark` object.|
||[delete()](/.bookmark#word-javascript/api/word/-bookmark-delete-member(1))|Deletes the bookmark.|
||[end](/.bookmark#word-javascript/api/word/-bookmark-end-member)|Specifies the ending character position of the bookmark.|
||[isColumn](/.bookmark#word-javascript/api/word/-bookmark-iscolumn-member)|Returns `true` if the bookmark is a table column.|
||[isEmpty](/.bookmark#word-javascript/api/word/-bookmark-isempty-member)|Returns `true` if the bookmark is empty.|
||[name](/.bookmark#word-javascript/api/word/-bookmark-name-member)|Returns the name of the `Bookmark` object.|
||[range](/.bookmark#word-javascript/api/word/-bookmark-range-member)|Returns a `Range` object that represents the portion of the document that's contained in the `Bookmark` object.|
||[select()](/.bookmark#word-javascript/api/word/-bookmark-select-member(1))|Selects the bookmark.|
||[start](/.bookmark#word-javascript/api/word/-bookmark-start-member)|Specifies the starting character position of the bookmark.|
||[storyType](/.bookmark#word-javascript/api/word/-bookmark-storytype-member)|Returns the story type for the bookmark.|
|[BookmarkCollection](/.bookmarkcollection)|[exists(name: string)](/.bookmarkcollection#word-javascript/api/word/-bookmarkcollection-exists-member(1))|Determines whether the specified bookmark exists.|
||[getItem(index: number)](/.bookmarkcollection#word-javascript/api/word/-bookmarkcollection-getitem-member(1))|Gets a `Bookmark` object by its index in the collection.|
||[items](/.bookmarkcollection#word-javascript/api/word/-bookmarkcollection-items-member)|Gets the loaded child items in this collection.|
|[BorderUniversal](/.borderuniversal)|[artStyle](/.borderuniversal#word-javascript/api/word/-borderuniversal-artstyle-member)|Specifies the graphical page-border design for the document.|
||[artWidth](/.borderuniversal#word-javascript/api/word/-borderuniversal-artwidth-member)|Specifies the width (in points) of the graphical page border specified in the `artStyle` property.|
||[color](/.borderuniversal#word-javascript/api/word/-borderuniversal-color-member)|Specifies the color for the `BorderUniversal` object.|
||[colorIndex](/.borderuniversal#word-javascript/api/word/-borderuniversal-colorindex-member)|Specifies the color for the `BorderUniversal` or Word.Font object.|
||[inside](/.borderuniversal#word-javascript/api/word/-borderuniversal-inside-member)|Returns `true` if an inside border can be applied to the specified object.|
||[isVisible](/.borderuniversal#word-javascript/api/word/-borderuniversal-isvisible-member)|Specifies whether the border is visible.|
||[lineStyle](/.borderuniversal#word-javascript/api/word/-borderuniversal-linestyle-member)|Specifies the line style of the border.|
||[lineWidth](/.borderuniversal#word-javascript/api/word/-borderuniversal-linewidth-member)|Specifies the line width of an object's border.|
|[BorderUniversalCollection](/.borderuniversalcollection)|[applyPageBordersToAllSections()](/.borderuniversalcollection#word-javascript/api/word/-borderuniversalcollection-applypageborderstoallsections-member(1))|Applies the specified page-border formatting to all sections in the document.|
||[getItem(index: number)](/.borderuniversalcollection#word-javascript/api/word/-borderuniversalcollection-getitem-member(1))|Gets a `Border` object by its index in the collection.|
||[items](/.borderuniversalcollection#word-javascript/api/word/-borderuniversalcollection-items-member)|Gets the loaded child items in this collection.|
|[Break](/.break)|[pageIndex](/.break#word-javascript/api/word/-break-pageindex-member)|Returns the page number on which the break occurs.|
||[range](/.break#word-javascript/api/word/-break-range-member)|Returns a `Range` object that represents the portion of the document that's contained in the break.|
|[BreakCollection](/.breakcollection)|[items](/.breakcollection#word-javascript/api/word/-breakcollection-items-member)|Gets the loaded child items in this collection.|
|[BuildingBlock](/.buildingblock)|[category](/.buildingblock#word-javascript/api/word/-buildingblock-category-member)|Returns a `BuildingBlockCategory` object that represents the category for the building block.|
||[delete()](/.buildingblock#word-javascript/api/word/-buildingblock-delete-member(1))|Deletes the building block.|
||[description](/.buildingblock#word-javascript/api/word/-buildingblock-description-member)|Specifies the description for the building block.|
||[id](/.buildingblock#word-javascript/api/word/-buildingblock-id-member)|Returns the internal identification number for the building block.|
||[index](/.buildingblock#word-javascript/api/word/-buildingblock-index-member)|Returns the position of this building block in a collection.|
||[insert(range: Word.Range, richText: boolean)](/.buildingblock#word-javascript/api/word/-buildingblock-insert-member(1))|Inserts the value of the building block into the document and returns a `Range` object that represents the contents of the building block within the document.|
||[insertType](/.buildingblock#word-javascript/api/word/-buildingblock-inserttype-member)|Specifies a `DocPartInsertType` value that represents how to insert the contents of the building block into the document.|
||[name](/.buildingblock#word-javascript/api/word/-buildingblock-name-member)|Specifies the name of the building block.|
||[type](/.buildingblock#word-javascript/api/word/-buildingblock-type-member)|Returns a `BuildingBlockTypeItem` object that represents the type for the building block.|
||[value](/.buildingblock#word-javascript/api/word/-buildingblock-value-member)|Specifies the contents of the building block.|
|[BuildingBlockCategory](/.buildingblockcategory)|[buildingBlocks](/.buildingblockcategory#word-javascript/api/word/-buildingblockcategory-buildingblocks-member)|Returns a `BuildingBlockCollection` object that represents the building blocks for the category.|
||[index](/.buildingblockcategory#word-javascript/api/word/-buildingblockcategory-index-member)|Returns the position of the `BuildingBlockCategory` object in a collection.|
||[name](/.buildingblockcategory#word-javascript/api/word/-buildingblockcategory-name-member)|Returns the name of the `BuildingBlockCategory` object.|
||[type](/.buildingblockcategory#word-javascript/api/word/-buildingblockcategory-type-member)|Returns a `BuildingBlockTypeItem` object that represents the type of building block for the building block category.|
|[BuildingBlockCategoryCollection](/.buildingblockcategorycollection)|[getCount()](/.buildingblockcategorycollection#word-javascript/api/word/-buildingblockcategorycollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/.buildingblockcategorycollection#word-javascript/api/word/-buildingblockcategorycollection-getitemat-member(1))|Returns a `BuildingBlockCategory` object that represents the specified item in the collection.|
|[BuildingBlockCollection](/.buildingblockcollection)|[add(name: string, range: Word.Range, description: string, insertType: Word.DocPartInsertType)](/.buildingblockcollection#word-javascript/api/word/-buildingblockcollection-add-member(1))|Creates a new building block and returns a `BuildingBlock` object.|
||[getCount()](/.buildingblockcollection#word-javascript/api/word/-buildingblockcollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/.buildingblockcollection#word-javascript/api/word/-buildingblockcollection-getitemat-member(1))|Returns a `BuildingBlock` object that represents the specified item in the collection.|
|[BuildingBlockEntryCollection](/.buildingblockentrycollection)|[add(name: string, type: Word.BuildingBlockType, category: string, range: Word.Range, description: string, insertType: Word.DocPartInsertType)](/.buildingblockentrycollection#word-javascript/api/word/-buildingblockentrycollection-add-member(1))|Creates a new building block entry in a template and returns a `BuildingBlock` object that represents the new building block entry.|
||[getCount()](/.buildingblockentrycollection#word-javascript/api/word/-buildingblockentrycollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/.buildingblockentrycollection#word-javascript/api/word/-buildingblockentrycollection-getitemat-member(1))|Returns a `BuildingBlock` object that represents the specified item in the collection.|
|[BuildingBlockGalleryContentControl](/.buildingblockgallerycontentcontrol)|[appearance](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[buildingBlockCategory](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-buildingblockcategory-member)|Specifies the category for the building block content control.|
||[buildingBlockType](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-buildingblocktype-member)|Specifies a `BuildingBlockType` value that represents the type of building block for the building block content control.|
||[color](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-color-member)|Specifies the red-green-blue (RGB) value of the color of the content control.|
||[copy()](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-copy-member(1))|Copies the content control from the active document to the Clipboard.|
||[cut()](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-cut-member(1))|Removes the content control from the active document and moves the content control to the Clipboard.|
||[delete(deleteContents?: boolean)](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-delete-member(1))|Deletes the content control and optionally its contents.|
||[id](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-id-member)|Gets the identification for the content control.|
||[isTemporary](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-istemporary-member)|Specifies whether to remove the content control from the active document when the user edits the contents of the control.|
||[level](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-level-member)|Gets the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.|
||[lockContentControl](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-lockcontentcontrol-member)|Specifies if the content control is locked (can't be deleted).|
||[lockContents](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-lockcontents-member)|Specifies if the contents of the content control are locked (not editable).|
||[placeholderText](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-placeholdertext-member)|Returns a `BuildingBlock` object that represents the placeholder text for the content control.|
||[range](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-range-member)|Returns a `Range` object that represents the contents of the content control in the active document.|
||[setPlaceholderText(options?: Word.ContentControlPlaceholderOptions)](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-setplaceholdertext-member(1))|Sets the placeholder text that displays in the content control until a user enters their own text.|
||[showingPlaceholderText](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-showingplaceholdertext-member)|Gets if the placeholder text for the content control is being displayed.|
||[tag](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-tag-member)|Specifies a tag to identify the content control.|
||[title](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-title-member)|Specifies the title for the content control.|
||[xmlMapping](/.buildingblockgallerycontentcontrol#word-javascript/api/word/-buildingblockgallerycontentcontrol-xmlmapping-member)|Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[BuildingBlockTypeItem](/.buildingblocktypeitem)|[categories](/.buildingblocktypeitem#word-javascript/api/word/-buildingblocktypeitem-categories-member)|Returns a `BuildingBlockCategoryCollection` object that represents the categories for a building block type.|
||[index](/.buildingblocktypeitem#word-javascript/api/word/-buildingblocktypeitem-index-member)|Returns the position of an item in a collection.|
||[name](/.buildingblocktypeitem#word-javascript/api/word/-buildingblocktypeitem-name-member)|Returns the localized name of a building block type.|
|[BuildingBlockTypeItemCollection](/.buildingblocktypeitemcollection)|[getByType(type: Word.BuildingBlockType)](/.buildingblocktypeitemcollection#word-javascript/api/word/-buildingblocktypeitemcollection-getbytype-member(1))|Gets a `BuildingBlockTypeItem` object by its type in the collection.|
||[getCount()](/.buildingblocktypeitemcollection#word-javascript/api/word/-buildingblocktypeitemcollection-getcount-member(1))|Returns the number of items in the collection.|
|[ColorFormat](/.colorformat)|[brightness](/.colorformat#word-javascript/api/word/-colorformat-brightness-member)|Specifies the brightness of a specified shape color.|
||[objectThemeColor](/.colorformat#word-javascript/api/word/-colorformat-objectthemecolor-member)|Specifies the theme color for a color format.|
||[rgb](/.colorformat#word-javascript/api/word/-colorformat-rgb-member)|Specifies the red-green-blue (RGB) value of the specified color.|
||[tintAndShade](/.colorformat#word-javascript/api/word/-colorformat-tintandshade-member)|Specifies the lightening or darkening of a specified shape's color.|
||[type](/.colorformat#word-javascript/api/word/-colorformat-type-member)|Returns the shape color type.|
|[CommentDetail](/.commentdetail)|[id](/.commentdetail#word-javascript/api/word/-commentdetail-id-member)|Represents the ID of this comment.|
||[replyIds](/.commentdetail#word-javascript/api/word/-commentdetail-replyids-member)|Represents the IDs of the replies to this comment.|
|[CommentEventArgs](/.commenteventargs)|[changeType](/.commenteventargs#word-javascript/api/word/-commenteventargs-changetype-member)|Represents how the comment changed event is triggered.|
||[commentDetails](/.commenteventargs#word-javascript/api/word/-commenteventargs-commentdetails-member)|Gets the CommentDetail array which contains the IDs and reply IDs of the involved comments.|
||[source](/.commenteventargs#word-javascript/api/word/-commenteventargs-source-member)|The source of the event.|
||[type](/.commenteventargs#word-javascript/api/word/-commenteventargs-type-member)|The event type.|
|[ContentControl](/.contentcontrol)|[buildingBlockGalleryContentControl](/.contentcontrol#word-javascript/api/word/-contentcontrol-buildingblockgallerycontentcontrol-member)|Gets the building block gallery-related data if the content control's Word.ContentControlType is `BuildingBlockGallery`.|
||[datePickerContentControl](/.contentcontrol#word-javascript/api/word/-contentcontrol-datepickercontentcontrol-member)|Gets the date picker-related data if the content control's Word.ContentControlType is `DatePicker`.|
||[groupContentControl](/.contentcontrol#word-javascript/api/word/-contentcontrol-groupcontentcontrol-member)|Gets the group-related data if the content control's Word.ContentControlType is `Group`.|
||[onCommentAdded](/.contentcontrol#word-javascript/api/word/-contentcontrol-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/.contentcontrol#word-javascript/api/word/-contentcontrol-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/.contentcontrol#word-javascript/api/word/-contentcontrol-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/.contentcontrol#word-javascript/api/word/-contentcontrol-oncommentselected-member)|Occurs when a comment is selected.|
||[pictureContentControl](/.contentcontrol#word-javascript/api/word/-contentcontrol-picturecontentcontrol-member)|Gets the picture-related data if the content control's Word.ContentControlType is `Picture`.|
||[repeatingSectionContentControl](/.contentcontrol#word-javascript/api/word/-contentcontrol-repeatingsectioncontentcontrol-member)|Gets the repeating section-related data if the content control's Word.ContentControlType is `RepeatingSection`.|
||[resetState()](/.contentcontrol#word-javascript/api/word/-contentcontrol-resetstate-member(1))|Resets the state of the content control.|
||[setState(contentControlState: Word.ContentControlState)](/.contentcontrol#word-javascript/api/word/-contentcontrol-setstate-member(1))|Sets the state of the content control.|
||[xmlMapping](/.contentcontrol#word-javascript/api/word/-contentcontrol-xmlmapping-member)|Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[ContentControlAddedEventArgs](/.contentcontroladdedeventargs)|[eventType](/.contentcontroladdedeventargs#word-javascript/api/word/-contentcontroladdedeventargs-eventtype-member)|The event type.|
|[ContentControlDataChangedEventArgs](/.contentcontroldatachangedeventargs)|[eventType](/.contentcontroldatachangedeventargs#word-javascript/api/word/-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
|[ContentControlDeletedEventArgs](/.contentcontroldeletedeventargs)|[eventType](/.contentcontroldeletedeventargs#word-javascript/api/word/-contentcontroldeletedeventargs-eventtype-member)|The event type.|
|[ContentControlEnteredEventArgs](/.contentcontrolenteredeventargs)|[eventType](/.contentcontrolenteredeventargs#word-javascript/api/word/-contentcontrolenteredeventargs-eventtype-member)|The event type.|
|[ContentControlExitedEventArgs](/.contentcontrolexitedeventargs)|[eventType](/.contentcontrolexitedeventargs#word-javascript/api/word/-contentcontrolexitedeventargs-eventtype-member)|The event type.|
|[ContentControlPlaceholderOptions](/.contentcontrolplaceholderoptions)|[buildingBlock](/.contentcontrolplaceholderoptions#word-javascript/api/word/-contentcontrolplaceholderoptions-buildingblock-member)|If provided, specifies the `BuildingBlock` object to use as placeholder.|
||[range](/.contentcontrolplaceholderoptions#word-javascript/api/word/-contentcontrolplaceholderoptions-range-member)|If provided, specifies the `Range` object to use as placeholder.|
||[text](/.contentcontrolplaceholderoptions#word-javascript/api/word/-contentcontrolplaceholderoptions-text-member)|If provided, specifies the text to use as placeholder.|
|[ContentControlSelectionChangedEventArgs](/.contentcontrolselectionchangedeventargs)|[eventType](/.contentcontrolselectionchangedeventargs#word-javascript/api/word/-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
|[CustomXmlAddNodeOptions](/.customxmladdnodeoptions)|[name](/.customxmladdnodeoptions#word-javascript/api/word/-customxmladdnodeoptions-name-member)|If provided, specifies the base name of the element to be added.|
||[namespaceUri](/.customxmladdnodeoptions#word-javascript/api/word/-customxmladdnodeoptions-namespaceuri-member)|If provided, specifies the namespace of the element to be appended.|
||[nextSibling](/.customxmladdnodeoptions#word-javascript/api/word/-customxmladdnodeoptions-nextsibling-member)|If provided, specifies the node which should become the next sibling of the new node.|
||[nodeType](/.customxmladdnodeoptions#word-javascript/api/word/-customxmladdnodeoptions-nodetype-member)|If provided, specifies the type of node to add.|
||[nodeValue](/.customxmladdnodeoptions#word-javascript/api/word/-customxmladdnodeoptions-nodevalue-member)|If provided, specifies the value of the added node for those nodes that allow text.|
|[CustomXmlAddSchemaOptions](/.customxmladdschemaoptions)|[alias](/.customxmladdschemaoptions#word-javascript/api/word/-customxmladdschemaoptions-alias-member)|If provided, specifies the alias of the schema to be added to the collection.|
||[fileName](/.customxmladdschemaoptions#word-javascript/api/word/-customxmladdschemaoptions-filename-member)|If provided, specifies the location of the schema on a disk.|
||[installForAllUsers](/.customxmladdschemaoptions#word-javascript/api/word/-customxmladdschemaoptions-installforallusers-member)|If provided, specifies whether, in the case where the schema is being added to the Schema Library, the Schema Library keys should be written to the registry (`HKEY_LOCAL_MACHINE` for all users or `HKEY_CURRENT_USER` for just the current user).|
||[namespaceUri](/.customxmladdschemaoptions#word-javascript/api/word/-customxmladdschemaoptions-namespaceuri-member)|If provided, specifies the namespace of the schema to be added to the collection.|
|[CustomXmlAddValidationErrorOptions](/.customxmladdvalidationerroroptions)|[clearedOnUpdate](/.customxmladdvalidationerroroptions#word-javascript/api/word/-customxmladdvalidationerroroptions-clearedonupdate-member)|If provided, specifies whether the error is to be cleared from the Word.CustomXmlValidationErrorCollection when the XML is corrected and updated.|
||[errorText](/.customxmladdvalidationerroroptions#word-javascript/api/word/-customxmladdvalidationerroroptions-errortext-member)|If provided, specifies the descriptive error text.|
|[CustomXmlAppendChildNodeOptions](/.customxmlappendchildnodeoptions)|[name](/.customxmlappendchildnodeoptions#word-javascript/api/word/-customxmlappendchildnodeoptions-name-member)|If provided, specifies the base name of the element to be appended.|
||[namespaceUri](/.customxmlappendchildnodeoptions#word-javascript/api/word/-customxmlappendchildnodeoptions-namespaceuri-member)|If provided, specifies the namespace of the element to be appended.|
||[nodeType](/.customxmlappendchildnodeoptions#word-javascript/api/word/-customxmlappendchildnodeoptions-nodetype-member)|If provided, specifies the type of node to append.|
||[nodeValue](/.customxmlappendchildnodeoptions#word-javascript/api/word/-customxmlappendchildnodeoptions-nodevalue-member)|If provided, specifies the value of the appended node for those nodes that allow text.|
|[CustomXmlInsertNodeBeforeOptions](/.customxmlinsertnodebeforeoptions)|[name](/.customxmlinsertnodebeforeoptions#word-javascript/api/word/-customxmlinsertnodebeforeoptions-name-member)|If provided, specifies the base name of the element to be inserted.|
||[namespaceUri](/.customxmlinsertnodebeforeoptions#word-javascript/api/word/-customxmlinsertnodebeforeoptions-namespaceuri-member)|If provided, specifies the namespace of the element to be inserted.|
||[nextSibling](/.customxmlinsertnodebeforeoptions#word-javascript/api/word/-customxmlinsertnodebeforeoptions-nextsibling-member)|If provided, specifies the context node.|
||[nodeType](/.customxmlinsertnodebeforeoptions#word-javascript/api/word/-customxmlinsertnodebeforeoptions-nodetype-member)|If provided, specifies the type of node to append.|
||[nodeValue](/.customxmlinsertnodebeforeoptions#word-javascript/api/word/-customxmlinsertnodebeforeoptions-nodevalue-member)|If provided, specifies the value of the inserted node for those nodes that allow text.|
|[CustomXmlInsertSubtreeBeforeOptions](/.customxmlinsertsubtreebeforeoptions)|[nextSibling](/.customxmlinsertsubtreebeforeoptions#word-javascript/api/word/-customxmlinsertsubtreebeforeoptions-nextsibling-member)|If provided, specifies the context node.|
|[CustomXmlNode](/.customxmlnode)|[appendChildNode(options?: Word.CustomXmlAppendChildNodeOptions)](/.customxmlnode#word-javascript/api/word/-customxmlnode-appendchildnode-member(1))|Appends a single node as the last child under the context element node in the tree.|
||[appendChildSubtree(xml: string)](/.customxmlnode#word-javascript/api/word/-customxmlnode-appendchildsubtree-member(1))|Adds a subtree as the last child under the context element node in the tree.|
||[attributes](/.customxmlnode#word-javascript/api/word/-customxmlnode-attributes-member)|Gets a `CustomXmlNodeCollection` object representing the attributes of the current element in the current node.|
||[baseName](/.customxmlnode#word-javascript/api/word/-customxmlnode-basename-member)|Gets the base name of the node without the namespace prefix, if one exists.|
||[childNodes](/.customxmlnode#word-javascript/api/word/-customxmlnode-childnodes-member)|Gets a `CustomXmlNodeCollection` object containing all of the child elements of the current node.|
||[delete()](/.customxmlnode#word-javascript/api/word/-customxmlnode-delete-member(1))|Deletes the current node from the tree (including all of its children, if any exist).|
||[firstChild](/.customxmlnode#word-javascript/api/word/-customxmlnode-firstchild-member)|Gets a `CustomXmlNode` object corresponding to the first child element of the current node.|
||[hasChildNodes()](/.customxmlnode#word-javascript/api/word/-customxmlnode-haschildnodes-member(1))|Specifies if the current element node has child element nodes.|
||[insertNodeBefore(options?: Word.CustomXmlInsertNodeBeforeOptions)](/.customxmlnode#word-javascript/api/word/-customxmlnode-insertnodebefore-member(1))|Inserts a new node just before the context node in the tree.|
||[insertSubtreeBefore(xml: string, options?: Word.CustomXmlInsertSubtreeBeforeOptions)](/.customxmlnode#word-javascript/api/word/-customxmlnode-insertsubtreebefore-member(1))|Inserts the specified subtree into the location just before the context node.|
||[lastChild](/.customxmlnode#word-javascript/api/word/-customxmlnode-lastchild-member)|Gets a `CustomXmlNode` object corresponding to the last child element of the current node.|
||[namespaceUri](/.customxmlnode#word-javascript/api/word/-customxmlnode-namespaceuri-member)|Gets the unique address identifier for the namespace of the node.|
||[nextSibling](/.customxmlnode#word-javascript/api/word/-customxmlnode-nextsibling-member)|Gets the next sibling node (element, comment, or processing instruction) of the current node.|
||[nodeType](/.customxmlnode#word-javascript/api/word/-customxmlnode-nodetype-member)|Gets the type of the current node.|
||[nodeValue](/.customxmlnode#word-javascript/api/word/-customxmlnode-nodevalue-member)|Specifies the value of the current node.|
||[ownerPart](/.customxmlnode#word-javascript/api/word/-customxmlnode-ownerpart-member)|Gets the object representing the part associated with this node.|
||[parentNode](/.customxmlnode#word-javascript/api/word/-customxmlnode-parentnode-member)|Gets the parent element node of the current node.|
||[previousSibling](/.customxmlnode#word-javascript/api/word/-customxmlnode-previoussibling-member)|Gets the previous sibling node (element, comment, or processing instruction) of the current node.|
||[removeChild(child: Word.CustomXmlNode)](/.customxmlnode#word-javascript/api/word/-customxmlnode-removechild-member(1))|Removes the specified child node from the tree.|
||[replaceChildNode(oldNode: Word.CustomXmlNode, options?: Word.CustomXmlReplaceChildNodeOptions)](/.customxmlnode#word-javascript/api/word/-customxmlnode-replacechildnode-member(1))|Removes the specified child node and replaces it with a different node in the same location.|
||[replaceChildSubtree(xml: string, oldNode: Word.CustomXmlNode)](/.customxmlnode#word-javascript/api/word/-customxmlnode-replacechildsubtree-member(1))|Removes the specified node and replaces it with a different subtree in the same location.|
||[selectNodes(xPath: string)](/.customxmlnode#word-javascript/api/word/-customxmlnode-selectnodes-member(1))|Selects a collection of nodes matching an XPath expression.|
||[selectSingleNode(xPath: string)](/.customxmlnode#word-javascript/api/word/-customxmlnode-selectsinglenode-member(1))|Selects a single node from a collection matching an XPath expression.|
||[text](/.customxmlnode#word-javascript/api/word/-customxmlnode-text-member)|Specifies the text for the current node.|
||[xml](/.customxmlnode#word-javascript/api/word/-customxmlnode-xml-member)|Gets the XML representation of the current node and its children.|
||[xpath](/.customxmlnode#word-javascript/api/word/-customxmlnode-xpath-member)|Gets a string with the canonicalized XPath for the current node.|
|[CustomXmlNodeCollection](/.customxmlnodecollection)|[getCount()](/.customxmlnodecollection#word-javascript/api/word/-customxmlnodecollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/.customxmlnodecollection#word-javascript/api/word/-customxmlnodecollection-getitem-member(1))|Returns a `CustomXmlNode` object that represents the specified item in the collection.|
||[items](/.customxmlnodecollection#word-javascript/api/word/-customxmlnodecollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPart](/.customxmlpart)|[addNode(parent: Word.CustomXmlNode, options?: Word.CustomXmlAddNodeOptions)](/.customxmlpart#word-javascript/api/word/-customxmlpart-addnode-member(1))|Adds a node to the XML tree.|
||[builtIn](/.customxmlpart#word-javascript/api/word/-customxmlpart-builtin-member)|Gets a value that indicates whether the `CustomXmlPart` is built-in.|
||[documentElement](/.customxmlpart#word-javascript/api/word/-customxmlpart-documentelement-member)|Gets the root element of a bound region of data in the document.|
||[errors](/.customxmlpart#word-javascript/api/word/-customxmlpart-errors-member)|Gets a `CustomXmlValidationErrorCollection` object that provides access to any XML validation errors.|
||[loadXml(xml: string)](/.customxmlpart#word-javascript/api/word/-customxmlpart-loadxml-member(1))|Populates the `CustomXmlPart` object from an XML string.|
||[namespaceManager](/.customxmlpart#word-javascript/api/word/-customxmlpart-namespacemanager-member)|Gets the set of namespace prefix mappings used against the current `CustomXmlPart` object.|
||[schemaCollection](/.customxmlpart#word-javascript/api/word/-customxmlpart-schemacollection-member)|Specifies a `CustomXmlSchemaCollection` object representing the set of schemas attached to a bound region of data in the document.|
||[selectNodes(xPath: string)](/.customxmlpart#word-javascript/api/word/-customxmlpart-selectnodes-member(1))|Selects a collection of nodes from a custom XML part.|
||[selectSingleNode(xPath: string)](/.customxmlpart#word-javascript/api/word/-customxmlpart-selectsinglenode-member(1))|Selects a single node within a custom XML part matching an XPath expression.|
||[xml](/.customxmlpart#word-javascript/api/word/-customxmlpart-xml-member)|Gets the XML representation of the current `CustomXmlPart` object.|
|[CustomXmlPrefixMapping](/.customxmlprefixmapping)|[namespaceUri](/.customxmlprefixmapping#word-javascript/api/word/-customxmlprefixmapping-namespaceuri-member)|Gets the unique address identifier for the namespace of the `CustomXmlPrefixMapping` object.|
||[prefix](/.customxmlprefixmapping#word-javascript/api/word/-customxmlprefixmapping-prefix-member)|Gets the prefix for the `CustomXmlPrefixMapping` object.|
|[CustomXmlPrefixMappingCollection](/.customxmlprefixmappingcollection)|[addNamespace(prefix: string, namespaceUri: string)](/.customxmlprefixmappingcollection#word-javascript/api/word/-customxmlprefixmappingcollection-addnamespace-member(1))|Adds a custom namespace/prefix mapping to use when querying an item.|
||[getCount()](/.customxmlprefixmappingcollection#word-javascript/api/word/-customxmlprefixmappingcollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/.customxmlprefixmappingcollection#word-javascript/api/word/-customxmlprefixmappingcollection-getitem-member(1))|Returns a `CustomXmlPrefixMapping` object that represents the specified item in the collection.|
||[items](/.customxmlprefixmappingcollection#word-javascript/api/word/-customxmlprefixmappingcollection-items-member)|Gets the loaded child items in this collection.|
||[lookupNamespace(prefix: string)](/.customxmlprefixmappingcollection#word-javascript/api/word/-customxmlprefixmappingcollection-lookupnamespace-member(1))|Gets the namespace corresponding to the specified prefix.|
||[lookupPrefix(namespaceUri: string)](/.customxmlprefixmappingcollection#word-javascript/api/word/-customxmlprefixmappingcollection-lookupprefix-member(1))|Gets the prefix corresponding to the specified namespace.|
|[CustomXmlReplaceChildNodeOptions](/.customxmlreplacechildnodeoptions)|[name](/.customxmlreplacechildnodeoptions#word-javascript/api/word/-customxmlreplacechildnodeoptions-name-member)|If provided, specifies the base name of the replacement element.|
||[namespaceUri](/.customxmlreplacechildnodeoptions#word-javascript/api/word/-customxmlreplacechildnodeoptions-namespaceuri-member)|If provided, specifies the namespace of the replacement element.|
||[nodeType](/.customxmlreplacechildnodeoptions#word-javascript/api/word/-customxmlreplacechildnodeoptions-nodetype-member)|If provided, specifies the type of the replacement node.|
||[nodeValue](/.customxmlreplacechildnodeoptions#word-javascript/api/word/-customxmlreplacechildnodeoptions-nodevalue-member)|If provided, specifies the value of the replacement node for those nodes that allow text.|
|[CustomXmlSchema](/.customxmlschema)|[delete()](/.customxmlschema#word-javascript/api/word/-customxmlschema-delete-member(1))|Deletes this schema from the Word.CustomXmlSchemaCollection object.|
||[location](/.customxmlschema#word-javascript/api/word/-customxmlschema-location-member)|Gets the location of the schema on a computer.|
||[namespaceUri](/.customxmlschema#word-javascript/api/word/-customxmlschema-namespaceuri-member)|Gets the unique address identifier for the namespace of the `CustomXmlSchema` object.|
||[reload()](/.customxmlschema#word-javascript/api/word/-customxmlschema-reload-member(1))|Reloads the schema from a file.|
|[CustomXmlSchemaCollection](/.customxmlschemacollection)|[add(options?: Word.CustomXmlAddSchemaOptions)](/.customxmlschemacollection#word-javascript/api/word/-customxmlschemacollection-add-member(1))|Adds one or more schemas to the schema collection that can then be added to a stream in the data store and to the schema library.|
||[addCollection(schemaCollection: Word.CustomXmlSchemaCollection)](/.customxmlschemacollection#word-javascript/api/word/-customxmlschemacollection-addcollection-member(1))|Adds an existing schema collection to the current schema collection.|
||[getCount()](/.customxmlschemacollection#word-javascript/api/word/-customxmlschemacollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/.customxmlschemacollection#word-javascript/api/word/-customxmlschemacollection-getitem-member(1))|Returns a `CustomXmlSchema` object that represents the specified item in the collection.|
||[getNamespaceUri()](/.customxmlschemacollection#word-javascript/api/word/-customxmlschemacollection-getnamespaceuri-member(1))|Returns the number of items in the collection.|
||[items](/.customxmlschemacollection#word-javascript/api/word/-customxmlschemacollection-items-member)|Gets the loaded child items in this collection.|
||[validate()](/.customxmlschemacollection#word-javascript/api/word/-customxmlschemacollection-validate-member(1))|Specifies whether the schemas in the schema collection are valid (conforms to the syntactic rules of XML and the rules for a specified vocabulary).|
|[CustomXmlValidationError](/.customxmlvalidationerror)|[delete()](/.customxmlvalidationerror#word-javascript/api/word/-customxmlvalidationerror-delete-member(1))|Deletes this `CustomXmlValidationError` object.|
||[errorCode](/.customxmlvalidationerror#word-javascript/api/word/-customxmlvalidationerror-errorcode-member)|Gets an integer representing the validation error in the `CustomXmlValidationError` object.|
||[name](/.customxmlvalidationerror#word-javascript/api/word/-customxmlvalidationerror-name-member)|Gets the name of the error in the `CustomXmlValidationError` object.If no errors exist, the property returns `Nothing`|
||[node](/.customxmlvalidationerror#word-javascript/api/word/-customxmlvalidationerror-node-member)|Gets the node associated with this `CustomXmlValidationError` object, if any exist.If no nodes exist, the property returns `Nothing`.|
||[text](/.customxmlvalidationerror#word-javascript/api/word/-customxmlvalidationerror-text-member)|Gets the text in the `CustomXmlValidationError` object.|
||[type](/.customxmlvalidationerror#word-javascript/api/word/-customxmlvalidationerror-type-member)|Gets the type of error generated from the `CustomXmlValidationError` object.|
|[CustomXmlValidationErrorCollection](/.customxmlvalidationerrorcollection)|[add(node: Word.CustomXmlNode, errorName: string, options?: Word.CustomXmlAddValidationErrorOptions)](/.customxmlvalidationerrorcollection#word-javascript/api/word/-customxmlvalidationerrorcollection-add-member(1))|Adds a `CustomXmlValidationError` object containing an XML validation error to the `CustomXmlValidationErrorCollection` object.|
||[getCount()](/.customxmlvalidationerrorcollection#word-javascript/api/word/-customxmlvalidationerrorcollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/.customxmlvalidationerrorcollection#word-javascript/api/word/-customxmlvalidationerrorcollection-getitem-member(1))|Returns a `CustomXmlValidationError` object that represents the specified item in the collection.|
||[items](/.customxmlvalidationerrorcollection#word-javascript/api/word/-customxmlvalidationerrorcollection-items-member)|Gets the loaded child items in this collection.|
|[DatePickerContentControl](/.datepickercontentcontrol)|[appearance](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[color](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-color-member)|Specifies the red-green-blue (RGB) value of the color of the content control.|
||[copy()](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-copy-member(1))|Copies the content control from the active document to the Clipboard.|
||[cut()](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-cut-member(1))|Removes the content control from the active document and moves the content control to the Clipboard.|
||[dateCalendarType](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-datecalendartype-member)|Specifies a `CalendarType` value that represents the calendar type for the date picker content control.|
||[dateDisplayFormat](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-datedisplayformat-member)|Specifies the format in which dates are displayed.|
||[dateDisplayLocale](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-datedisplaylocale-member)|Specifies a `LanguageId` that represents the language format for the date displayed in the date picker content control.|
||[dateStorageFormat](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-datestorageformat-member)|Specifies a `ContentControlDateStorageFormat` value that represents the format for storage and retrieval of dates when the date picker content control is bound to the XML data store of the active document.|
||[delete(deleteContents?: boolean)](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-delete-member(1))|Deletes this content control and the contents of the content control.|
||[id](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-id-member)|Gets the identification for the content control.|
||[isTemporary](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-istemporary-member)|Specifies whether to remove the content control from the active document when the user edits the contents of the control.|
||[level](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-level-member)|Specifies the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.|
||[lockContentControl](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-lockcontentcontrol-member)|Specifies if the content control is locked (can't be deleted).|
||[lockContents](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-lockcontents-member)|Specifies if the contents of the content control are locked (not editable).|
||[placeholderText](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-placeholdertext-member)|Returns a `BuildingBlock` object that represents the placeholder text for the content control.|
||[range](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-range-member)|Gets a `Range` object that represents the contents of the content control in the active document.|
||[setPlaceholderText(options?: Word.ContentControlPlaceholderOptions)](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-setplaceholdertext-member(1))|Sets the placeholder text that displays in the content control until a user enters their own text.|
||[showingPlaceholderText](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-showingplaceholdertext-member)|Gets whether the placeholder text for the content control is being displayed.|
||[tag](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-tag-member)|Specifies a tag to identify the content control.|
||[title](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-title-member)|Specifies the title for the content control.|
||[xmlMapping](/.datepickercontentcontrol#word-javascript/api/word/-datepickercontentcontrol-xmlmapping-member)|Gets an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[Document](/.document)|[attachedTemplate](/.document#word-javascript/api/word/-document-attachedtemplate-member)|Specifies a `Template` object that represents the template attached to the document.|
||[autoHyphenation](/.document#word-javascript/api/word/-document-autohyphenation-member)|Specifies if automatic hyphenation is turned on for the document.|
||[autoSaveOn](/.document#word-javascript/api/word/-document-autosaveon-member)|Specifies if the edits in the document are automatically saved.|
||[bibliography](/.document#word-javascript/api/word/-document-bibliography-member)|Returns a `Bibliography` object that represents the bibliography references contained within the document.|
||[bookmarks](/.document#word-javascript/api/word/-document-bookmarks-member)|Returns a `BookmarkCollection` object that represents all the bookmarks in the document.|
||[consecutiveHyphensLimit](/.document#word-javascript/api/word/-document-consecutivehyphenslimit-member)|Specifies the maximum number of consecutive lines that can end with hyphens.|
||[detectLanguage()](/.document#word-javascript/api/word/-document-detectlanguage-member(1))|Analyzes the document text to determine the language.|
||[documentLibraryVersions](/.document#word-javascript/api/word/-document-documentlibraryversions-member)|Returns a `DocumentLibraryVersionCollection` object that represents the collection of versions of a shared document that has versioning enabled and that's stored in a document library on a server.|
||[frames](/.document#word-javascript/api/word/-document-frames-member)|Returns a `FrameCollection` object that represents all the frames in the document.|
||[hyperlinks](/.document#word-javascript/api/word/-document-hyperlinks-member)|Returns a `HyperlinkCollection` object that represents all the hyperlinks in the document.|
||[hyphenateCaps](/.document#word-javascript/api/word/-document-hyphenatecaps-member)|Specifies whether words in all capital letters can be hyphenated.|
||[indexes](/.document#word-javascript/api/word/-document-indexes-member)|Returns an `IndexCollection` object that represents all the indexes in the document.|
||[languageDetected](/.document#word-javascript/api/word/-document-languagedetected-member)|Specifies whether Microsoft Word has detected the language of the document text.|
||[manualHyphenation()](/.document#word-javascript/api/word/-document-manualhyphenation-member(1))|Initiates manual hyphenation of a document, one line at a time.|
||[pageSetup](/.document#word-javascript/api/word/-document-pagesetup-member)|Returns a `PageSetup` object that's associated with the document.|
|[DocumentLibraryVersion](/.documentlibraryversion)|[comments](/.documentlibraryversion#word-javascript/api/word/-documentlibraryversion-comments-member)|Gets any optional comments associated with this version of the shared document.|
||[modified](/.documentlibraryversion#word-javascript/api/word/-documentlibraryversion-modified-member)|Gets the date and time at which this version of the shared document was last saved to the server.|
||[modifiedBy](/.documentlibraryversion#word-javascript/api/word/-documentlibraryversion-modifiedby-member)|Gets the name of the user who last saved this version of the shared document to the server.|
|[DocumentLibraryVersionCollection](/.documentlibraryversioncollection)|[getItem(index: number)](/.documentlibraryversioncollection#word-javascript/api/word/-documentlibraryversioncollection-getitem-member(1))|Gets a `DocumentLibraryVersion` object by its index in the collection.|
||[isVersioningEnabled()](/.documentlibraryversioncollection#word-javascript/api/word/-documentlibraryversioncollection-isversioningenabled-member(1))|Returns whether the document library in which the active document is saved on the server is configured to create a backup copy, or version, each time the file is edited on the website.|
||[items](/.documentlibraryversioncollection#word-javascript/api/word/-documentlibraryversioncollection-items-member)|Gets the loaded child items in this collection.|
|[DropCap](/.dropcap)|[clear()](/.dropcap#word-javascript/api/word/-dropcap-clear-member(1))|Removes the dropped capital letter formatting.|
||[distanceFromText](/.dropcap#word-javascript/api/word/-dropcap-distancefromtext-member)|Gets the distance (in points) between the dropped capital letter and the paragraph text.|
||[enable()](/.dropcap#word-javascript/api/word/-dropcap-enable-member(1))|Formats the first character in the specified paragraph as a dropped capital letter.|
||[fontName](/.dropcap#word-javascript/api/word/-dropcap-fontname-member)|Gets the name of the font for the dropped capital letter.|
||[linesToDrop](/.dropcap#word-javascript/api/word/-dropcap-linestodrop-member)|Gets the height (in lines) of the dropped capital letter.|
||[position](/.dropcap#word-javascript/api/word/-dropcap-position-member)|Gets the position of the dropped capital letter.|
|[Field](/.field)|[copyToClipboard()](/.field#word-javascript/api/word/-field-copytoclipboard-member(1))|Copies the field to the Clipboard.|
||[cut()](/.field#word-javascript/api/word/-field-cut-member(1))|Removes the field from the document and places it on the Clipboard.|
||[doClick()](/.field#word-javascript/api/word/-field-doclick-member(1))|Clicks the field.|
||[linkFormat](/.field#word-javascript/api/word/-field-linkformat-member)|Gets a `LinkFormat` object that represents the link options of the field.|
||[oleFormat](/.field#word-javascript/api/word/-field-oleformat-member)|Gets an `OleFormat` object that represents the OLE characteristics (other than linking) for the field.|
||[unlink()](/.field#word-javascript/api/word/-field-unlink-member(1))|Replaces the field with its most recent result.|
||[updateSource()](/.field#word-javascript/api/word/-field-updatesource-member(1))|Saves the changes made to the results of an {@link https://support.microsoft.com/office/1c34d6d6-0de3-4b5c-916a-2ff950fb629e | INCLUDETEXT field} back to the source document.|
|[FillFormat](/.fillformat)|[backgroundColor](/.fillformat#word-javascript/api/word/-fillformat-backgroundcolor-member)|Returns a `ColorFormat` object that represents the background color for the fill.|
||[foregroundColor](/.fillformat#word-javascript/api/word/-fillformat-foregroundcolor-member)|Returns a `ColorFormat` object that represents the foreground color for the fill.|
||[gradientAngle](/.fillformat#word-javascript/api/word/-fillformat-gradientangle-member)|Specifies the angle of the gradient fill.|
||[gradientColorType](/.fillformat#word-javascript/api/word/-fillformat-gradientcolortype-member)|Gets the gradient color type.|
||[gradientDegree](/.fillformat#word-javascript/api/word/-fillformat-gradientdegree-member)|Returns how dark or light a one-color gradient fill is.|
||[gradientStyle](/.fillformat#word-javascript/api/word/-fillformat-gradientstyle-member)|Returns the gradient style for the fill.|
||[gradientVariant](/.fillformat#word-javascript/api/word/-fillformat-gradientvariant-member)|Returns the gradient variant for the fill as an integer value from 1 to 4 for most gradient fills.|
||[isVisible](/.fillformat#word-javascript/api/word/-fillformat-isvisible-member)|Specifies if the object, or the formatting applied to it, is visible.|
||[pattern](/.fillformat#word-javascript/api/word/-fillformat-pattern-member)|Returns a `PatternType` value that represents the pattern applied to the fill or line.|
||[presetGradientType](/.fillformat#word-javascript/api/word/-fillformat-presetgradienttype-member)|Returns the preset gradient type for the fill.|
||[presetTexture](/.fillformat#word-javascript/api/word/-fillformat-presettexture-member)|Gets the preset texture.|
||[rotateWithObject](/.fillformat#word-javascript/api/word/-fillformat-rotatewithobject-member)|Specifies whether the fill rotates with the shape.|
||[setOneColorGradient(style: Word.GradientStyle, variant: number, degree: number)](/.fillformat#word-javascript/api/word/-fillformat-setonecolorgradient-member(1))|Sets the fill to a one-color gradient.|
||[setPatterned(pattern: Word.PatternType)](/.fillformat#word-javascript/api/word/-fillformat-setpatterned-member(1))|Sets the fill to a pattern.|
||[setPresetGradient(style: Word.GradientStyle, variant: number, presetGradientType: Word.PresetGradientType)](/.fillformat#word-javascript/api/word/-fillformat-setpresetgradient-member(1))|Sets the fill to a preset gradient.|
||[setPresetTextured(presetTexture: Word.PresetTexture)](/.fillformat#word-javascript/api/word/-fillformat-setpresettextured-member(1))|Sets the fill to a preset texture.|
||[setTwoColorGradient(style: Word.GradientStyle, variant: number)](/.fillformat#word-javascript/api/word/-fillformat-settwocolorgradient-member(1))|Sets the fill to a two-color gradient.|
||[solid()](/.fillformat#word-javascript/api/word/-fillformat-solid-member(1))|Sets the fill to a uniform color.|
||[textureAlignment](/.fillformat#word-javascript/api/word/-fillformat-texturealignment-member)|Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.|
||[textureHorizontalScale](/.fillformat#word-javascript/api/word/-fillformat-texturehorizontalscale-member)|Specifies the horizontal scaling factor for the texture fill.|
||[textureName](/.fillformat#word-javascript/api/word/-fillformat-texturename-member)|Returns the name of the custom texture file for the fill.|
||[textureOffsetX](/.fillformat#word-javascript/api/word/-fillformat-textureoffsetx-member)|Specifies the horizontal offset of the texture from the origin in points.|
||[textureOffsetY](/.fillformat#word-javascript/api/word/-fillformat-textureoffsety-member)|Specifies the vertical offset of the texture.|
||[textureTile](/.fillformat#word-javascript/api/word/-fillformat-texturetile-member)|Specifies whether the texture is tiled.|
||[textureType](/.fillformat#word-javascript/api/word/-fillformat-texturetype-member)|Returns the texture type for the fill.|
||[textureVerticalScale](/.fillformat#word-javascript/api/word/-fillformat-textureverticalscale-member)|Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.|
||[transparency](/.fillformat#word-javascript/api/word/-fillformat-transparency-member)|Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).|
||[type](/.fillformat#word-javascript/api/word/-fillformat-type-member)|Gets the fill format type.|
|[Font](/.font)|[allCaps](/.font#word-javascript/api/word/-font-allcaps-member)|Specifies whether the font is formatted as all capital letters, which makes lowercase letters appear as uppercase letters.|
||[boldBidirectional](/.font#word-javascript/api/word/-font-boldbidirectional-member)|Specifies whether the font is formatted as bold in a right-to-left language document.|
||[borders](/.font#word-javascript/api/word/-font-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the font.|
||[colorIndex](/.font#word-javascript/api/word/-font-colorindex-member)|Specifies a `ColorIndex` value that represents the color for the font.|
||[colorIndexBidirectional](/.font#word-javascript/api/word/-font-colorindexbidirectional-member)|Specifies the color for the `Font` object in a right-to-left language document.|
||[contextualAlternates](/.font#word-javascript/api/word/-font-contextualalternates-member)|Specifies whether contextual alternates are enabled for the font.|
||[decreaseFontSize()](/.font#word-javascript/api/word/-font-decreasefontsize-member(1))|Decreases the font size to the next available size.|
||[diacriticColor](/.font#word-javascript/api/word/-font-diacriticcolor-member)|Specifies the color to be used for diacritics for the `Font` object.|
||[disableCharacterSpaceGrid](/.font#word-javascript/api/word/-font-disablecharacterspacegrid-member)|Specifies whether Microsoft Word ignores the number of characters per line for the corresponding `Font` object.|
||[emboss](/.font#word-javascript/api/word/-font-emboss-member)|Specifies whether the font is formatted as embossed.|
||[emphasisMark](/.font#word-javascript/api/word/-font-emphasismark-member)|Specifies an `EmphasisMark` value that represents the emphasis mark for a character or designated character string.|
||[engrave](/.font#word-javascript/api/word/-font-engrave-member)|Specifies whether the font is formatted as engraved.|
||[fill](/.font#word-javascript/api/word/-font-fill-member)|Returns a `FillFormat` object that contains fill formatting properties for the font used by the range of text.|
||[glow](/.font#word-javascript/api/word/-font-glow-member)|Returns a `GlowFormat` object that represents the glow formatting for the font used by the range of text.|
||[increaseFontSize()](/.font#word-javascript/api/word/-font-increasefontsize-member(1))|Increases the font size to the next available size.|
||[italicBidirectional](/.font#word-javascript/api/word/-font-italicbidirectional-member)|Specifies whether the font is italicized in a right-to-left language document.|
||[kerning](/.font#word-javascript/api/word/-font-kerning-member)|Specifies the minimum font size for which Microsoft Word will adjust kerning automatically.|
||[ligature](/.font#word-javascript/api/word/-font-ligature-member)|Specifies the ligature setting for the `Font` object.|
||[line](/.font#word-javascript/api/word/-font-line-member)|Returns a `LineFormat` object that specifies the formatting for a line.|
||[nameAscii](/.font#word-javascript/api/word/-font-nameascii-member)|Specifies the font used for Latin text (characters with character codes from 0 (zero) through 127).|
||[nameBidirectional](/.font#word-javascript/api/word/-font-namebidirectional-member)|Specifies the font name in a right-to-left language document.|
||[nameFarEast](/.font#word-javascript/api/word/-font-namefareast-member)|Specifies the East Asian font name.|
||[nameOther](/.font#word-javascript/api/word/-font-nameother-member)|Specifies the font used for characters with codes from 128 through 255.|
||[numberForm](/.font#word-javascript/api/word/-font-numberform-member)|Specifies the number form setting for an OpenType font.|
||[numberSpacing](/.font#word-javascript/api/word/-font-numberspacing-member)|Specifies the number spacing setting for the font.|
||[outline](/.font#word-javascript/api/word/-font-outline-member)|Specifies if the font is formatted as outlined.|
||[position](/.font#word-javascript/api/word/-font-position-member)|Specifies the position of text (in points) relative to the base line.|
||[reflection](/.font#word-javascript/api/word/-font-reflection-member)|Returns a `ReflectionFormat` object that represents the reflection formatting for a shape.|
||[reset()](/.font#word-javascript/api/word/-font-reset-member(1))|Removes manual character formatting.|
||[scaling](/.font#word-javascript/api/word/-font-scaling-member)|Specifies the scaling percentage applied to the font.|
||[setAsTemplateDefault()](/.font#word-javascript/api/word/-font-setastemplatedefault-member(1))|Sets the specified font formatting as the default for the active document and all new documents based on the active template.|
||[shadow](/.font#word-javascript/api/word/-font-shadow-member)|Specifies if the font is formatted as shadowed.|
||[sizeBidirectional](/.font#word-javascript/api/word/-font-sizebidirectional-member)|Specifies the font size in points for right-to-left text.|
||[smallCaps](/.font#word-javascript/api/word/-font-smallcaps-member)|Specifies whether the font is formatted as small caps, which makes lowercase letters appear as small uppercase letters.|
||[spacing](/.font#word-javascript/api/word/-font-spacing-member)|Specifies the spacing between characters.|
||[stylisticSet](/.font#word-javascript/api/word/-font-stylisticset-member)|Specifies the stylistic set for the font.|
||[textColor](/.font#word-javascript/api/word/-font-textcolor-member)|Returns a `ColorFormat` object that represents the color for the font.|
||[textShadow](/.font#word-javascript/api/word/-font-textshadow-member)|Returns a `ShadowFormat` object that specifies the shadow formatting for the font.|
||[threeDimensionalFormat](/.font#word-javascript/api/word/-font-threedimensionalformat-member)|Returns a `ThreeDimensionalFormat` object that contains 3-dimensional (3D) effect formatting properties for the font.|
||[underlineColor](/.font#word-javascript/api/word/-font-underlinecolor-member)|Specifies the color of the underline for the `Font` object.|
|[Frame](/.frame)|[borders](/.frame#word-javascript/api/word/-frame-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the frame.|
||[copy()](/.frame#word-javascript/api/word/-frame-copy-member(1))|Copies the frame to the Clipboard.|
||[cut()](/.frame#word-javascript/api/word/-frame-cut-member(1))|Removes the frame from the document and places it on the Clipboard.|
||[delete()](/.frame#word-javascript/api/word/-frame-delete-member(1))|Deletes the frame.|
||[height](/.frame#word-javascript/api/word/-frame-height-member)|Specifies the height (in points) of the frame.|
||[heightRule](/.frame#word-javascript/api/word/-frame-heightrule-member)|Specifies a `FrameSizeRule` value that represents the rule for determining the height of the frame.|
||[horizontalDistanceFromText](/.frame#word-javascript/api/word/-frame-horizontaldistancefromtext-member)|Specifies the horizontal distance between the frame and the surrounding text, in points.|
||[horizontalPosition](/.frame#word-javascript/api/word/-frame-horizontalposition-member)|Specifies the horizontal distance between the edge of the frame and the item specified by the `relativeHorizontalPosition` property.|
||[lockAnchor](/.frame#word-javascript/api/word/-frame-lockanchor-member)|Specifies if the frame is locked.|
||[range](/.frame#word-javascript/api/word/-frame-range-member)|Returns a `Range` object that represents the portion of the document that's contained within the frame.|
||[relativeHorizontalPosition](/.frame#word-javascript/api/word/-frame-relativehorizontalposition-member)|Specifies the relative horizontal position of the frame.|
||[relativeVerticalPosition](/.frame#word-javascript/api/word/-frame-relativeverticalposition-member)|Specifies the relative vertical position of the frame.|
||[select()](/.frame#word-javascript/api/word/-frame-select-member(1))|Selects the frame.|
||[shading](/.frame#word-javascript/api/word/-frame-shading-member)|Returns a `ShadingUniversal` object that refers to the shading formatting for the frame.|
||[textWrap](/.frame#word-javascript/api/word/-frame-textwrap-member)|Specifies if document text wraps around the frame.|
||[verticalDistanceFromText](/.frame#word-javascript/api/word/-frame-verticaldistancefromtext-member)|Specifies the vertical distance (in points) between the frame and the surrounding text.|
||[verticalPosition](/.frame#word-javascript/api/word/-frame-verticalposition-member)|Specifies the vertical distance between the edge of the frame and the item specified by the `relativeVerticalPosition` property.|
||[width](/.frame#word-javascript/api/word/-frame-width-member)|Specifies the width (in points) of the frame.|
||[widthRule](/.frame#word-javascript/api/word/-frame-widthrule-member)|Specifies the rule used to determine the width of the frame.|
|[FrameCollection](/.framecollection)|[add(range: Word.Range)](/.framecollection#word-javascript/api/word/-framecollection-add-member(1))|Returns a `Frame` object that represents a new frame added to a range, selection, or document.|
||[delete()](/.framecollection#word-javascript/api/word/-framecollection-delete-member(1))|Deletes the `FrameCollection` object.|
||[getItem(index: number)](/.framecollection#word-javascript/api/word/-framecollection-getitem-member(1))|Gets a `Frame` object by its index in the collection.|
||[items](/.framecollection#word-javascript/api/word/-framecollection-items-member)|Gets the loaded child items in this collection.|
|[GlowFormat](/.glowformat)|[color](/.glowformat#word-javascript/api/word/-glowformat-color-member)|Returns a `ColorFormat` object that represents the color for a glow effect.|
||[radius](/.glowformat#word-javascript/api/word/-glowformat-radius-member)|Specifies the length of the radius for a glow effect.|
||[transparency](/.glowformat#word-javascript/api/word/-glowformat-transparency-member)|Specifies the degree of transparency for the glow effect as a value between 0.0 (opaque) and 1.0 (clear).|
|[GroupContentControl](/.groupcontentcontrol)|[appearance](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[color](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-color-member)|Specifies the red-green-blue (RGB) value of the color of the content control.|
||[copy()](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-copy-member(1))|Copies the content control from the active document to the Clipboard.|
||[cut()](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-cut-member(1))|Removes the content control from the active document and moves the content control to the Clipboard.|
||[delete(deleteContents: boolean)](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-delete-member(1))|Deletes the content control and optionally its contents.|
||[id](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-id-member)|Returns the identification for the content control.|
||[isTemporary](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-istemporary-member)|Specifies whether to remove the content control from the active document when the user edits the contents of the control.|
||[level](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-level-member)|Gets the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.|
||[lockContentControl](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-lockcontentcontrol-member)|Specifies if the content control is locked (can't be deleted).|
||[lockContents](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-lockcontents-member)|Specifies if the contents of the content control are locked (not editable).|
||[placeholderText](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-placeholdertext-member)|Returns a `BuildingBlock` object that represents the placeholder text for the content control.|
||[range](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-range-member)|Gets a `Range` object that represents the contents of the content control in the active document.|
||[setPlaceholderText(options?: Word.ContentControlPlaceholderOptions)](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-setplaceholdertext-member(1))|Sets the placeholder text that displays in the content control until a user enters their own text.|
||[showingPlaceholderText](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-showingplaceholdertext-member)|Returns whether the placeholder text for the content control is being displayed.|
||[tag](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-tag-member)|Specifies a tag to identify the content control.|
||[title](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-title-member)|Specifies the title for the content control.|
||[ungroup()](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-ungroup-member(1))|Removes the group content control from the document so that its child content controls are no longer nested and can be freely edited.|
||[xmlMapping](/.groupcontentcontrol#word-javascript/api/word/-groupcontentcontrol-xmlmapping-member)|Gets an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[Hyperlink](/.hyperlink)|[addToFavorites()](/.hyperlink#word-javascript/api/word/-hyperlink-addtofavorites-member(1))|Creates a shortcut to the document or hyperlink and adds it to the **Favorites** folder.|
||[address](/.hyperlink#word-javascript/api/word/-hyperlink-address-member)|Specifies the address (for example, a file name or URL) of the hyperlink.|
||[createNewDocument(fileName: string, editNow: boolean, overwrite: boolean)](/.hyperlink#word-javascript/api/word/-hyperlink-createnewdocument-member(1))|Creates a new document linked to the hyperlink.|
||[delete()](/.hyperlink#word-javascript/api/word/-hyperlink-delete-member(1))|Deletes the hyperlink.|
||[emailSubject](/.hyperlink#word-javascript/api/word/-hyperlink-emailsubject-member)|Specifies the text string for the hyperlink's subject line.|
||[isExtraInfoRequired](/.hyperlink#word-javascript/api/word/-hyperlink-isextrainforequired-member)|Returns `true` if extra information is required to resolve the hyperlink.|
||[name](/.hyperlink#word-javascript/api/word/-hyperlink-name-member)|Returns the name of the `Hyperlink` object.|
||[range](/.hyperlink#word-javascript/api/word/-hyperlink-range-member)|Returns a `Range` object that represents the portion of the document that's contained within the hyperlink.|
||[screenTip](/.hyperlink#word-javascript/api/word/-hyperlink-screentip-member)|Specifies the text that appears as a ScreenTip when the mouse pointer is positioned over the hyperlink.|
||[subAddress](/.hyperlink#word-javascript/api/word/-hyperlink-subaddress-member)|Specifies a named location in the destination of the hyperlink.|
||[target](/.hyperlink#word-javascript/api/word/-hyperlink-target-member)|Specifies the name of the frame or window in which to load the hyperlink.|
||[textToDisplay](/.hyperlink#word-javascript/api/word/-hyperlink-texttodisplay-member)|Specifies the hyperlink's visible text in the document.|
||[type](/.hyperlink#word-javascript/api/word/-hyperlink-type-member)|Returns the hyperlink type.|
|[HyperlinkAddOptions](/.hyperlinkaddoptions)|[address](/.hyperlinkaddoptions#word-javascript/api/word/-hyperlinkaddoptions-address-member)|If provided, specifies the address (e.g., URL or file path) of the hyperlink.|
||[screenTip](/.hyperlinkaddoptions#word-javascript/api/word/-hyperlinkaddoptions-screentip-member)|If provided, specifies the text that appears as a tooltip.|
||[subAddress](/.hyperlinkaddoptions#word-javascript/api/word/-hyperlinkaddoptions-subaddress-member)|If provided, specifies the location within the file or document.|
||[target](/.hyperlinkaddoptions#word-javascript/api/word/-hyperlinkaddoptions-target-member)|If provided, specifies the name of the frame or window in which to load the hyperlink.|
||[textToDisplay](/.hyperlinkaddoptions#word-javascript/api/word/-hyperlinkaddoptions-texttodisplay-member)|If provided, specifies the visible text of the hyperlink.|
|[HyperlinkCollection](/.hyperlinkcollection)|[add(anchor: Word.Range, options?: Word.HyperlinkAddOptions)](/.hyperlinkcollection#word-javascript/api/word/-hyperlinkcollection-add-member(1))|Returns a `Hyperlink` object that represents a new hyperlink added to a range, selection, or document.|
||[items](/.hyperlinkcollection#word-javascript/api/word/-hyperlinkcollection-items-member)|Gets the loaded child items in this collection.|
|[Index](/.index)|[delete()](/.index#word-javascript/api/word/-index-delete-member(1))|Deletes this index.|
||[filter](/.index#word-javascript/api/word/-index-filter-member)|Gets a value that represents how Microsoft Word classifies the first character of entries in the index.|
||[headingSeparator](/.index#word-javascript/api/word/-index-headingseparator-member)|Gets the text between alphabetical groups (entries that start with the same letter) in the index.|
||[indexLanguage](/.index#word-javascript/api/word/-index-indexlanguage-member)|Gets a `LanguageId` value that represents the sorting language to use for the index.|
||[numberOfColumns](/.index#word-javascript/api/word/-index-numberofcolumns-member)|Gets the number of columns for each page of the index.|
||[range](/.index#word-javascript/api/word/-index-range-member)|Returns a `Range` object that represents the portion of the document that is contained within the index.|
||[rightAlignPageNumbers](/.index#word-javascript/api/word/-index-rightalignpagenumbers-member)|Specifies if page numbers are aligned with the right margin in the index.|
||[separateAccentedLetterHeadings](/.index#word-javascript/api/word/-index-separateaccentedletterheadings-member)|Gets if the index contains separate headings for accented letters (for example, words that begin with "À" are under|
||[sortBy](/.index#word-javascript/api/word/-index-sortby-member)|Specifies the sorting criteria for the index.|
||[tabLeader](/.index#word-javascript/api/word/-index-tableader-member)|Specifies the leader character between entries in the index and their associated page numbers.|
||[type](/.index#word-javascript/api/word/-index-type-member)|Gets the index type.|
|[IndexAddOptions](/.indexaddoptions)|[headingSeparator](/.indexaddoptions#word-javascript/api/word/-indexaddoptions-headingseparator-member)|If provided, specifies the text between alphabetical groups (entries that start with the same letter) in the index.|
||[indexLanguage](/.indexaddoptions#word-javascript/api/word/-indexaddoptions-indexlanguage-member)|If provided, specifies the sorting language to be used for the index being added.|
||[numberOfColumns](/.indexaddoptions#word-javascript/api/word/-indexaddoptions-numberofcolumns-member)|If provided, specifies the number of columns for each page of the index.|
||[rightAlignPageNumbers](/.indexaddoptions#word-javascript/api/word/-indexaddoptions-rightalignpagenumbers-member)|If provided, specifies whether the page numbers in the generated index are aligned with the right margin.|
||[separateAccentedLetterHeadings](/.indexaddoptions#word-javascript/api/word/-indexaddoptions-separateaccentedletterheadings-member)|If provided, specifies whether to include separate headings for accented letters in the index.|
||[sortBy](/.indexaddoptions#word-javascript/api/word/-indexaddoptions-sortby-member)|If provided, specifies the sorting criteria to be used for the index being added.|
||[type](/.indexaddoptions#word-javascript/api/word/-indexaddoptions-type-member)|If provided, specifies whether subentries are on the same line (run-in) as the main entry or on a separate line (indented) from the main entry.|
|[IndexCollection](/.indexcollection)|[add(range: Word.Range, indexAddOptions?: Word.IndexAddOptions)](/.indexcollection#word-javascript/api/word/-indexcollection-add-member(1))|Returns an `Index` object that represents a new index added to the document.|
||[getFormat()](/.indexcollection#word-javascript/api/word/-indexcollection-getformat-member(1))|Gets the `IndexFormat` value that represents the formatting for the indexes in the document.|
||[getItem(index: number)](/.indexcollection#word-javascript/api/word/-indexcollection-getitem-member(1))|Gets an `Index` object by its index in the collection.|
||[items](/.indexcollection#word-javascript/api/word/-indexcollection-items-member)|Gets the loaded child items in this collection.|
||[markAllEntries(range: Word.Range, markAllEntriesOptions?: Word.IndexMarkAllEntriesOptions)](/.indexcollection#word-javascript/api/word/-indexcollection-markallentries-member(1))|Inserts an {@link https://support.microsoft.com/office/abaf7c78-6e21-418d-bf8b-f8186d2e4d08 | XE (Index Entry) field} after all instances of the text in the range.|
|[IndexMarkAllEntriesOptions](/.indexmarkallentriesoptions)|[bold](/.indexmarkallentriesoptions#word-javascript/api/word/-indexmarkallentriesoptions-bold-member)|If provided, specifies whether to add bold formatting to page numbers for index entries.|
||[bookmarkName](/.indexmarkallentriesoptions#word-javascript/api/word/-indexmarkallentriesoptions-bookmarkname-member)|If provided, specifies the bookmark name that marks the range of pages you want to appear in the index.|
||[crossReference](/.indexmarkallentriesoptions#word-javascript/api/word/-indexmarkallentriesoptions-crossreference-member)|If provided, specifies the cross-reference that will appear in the index.|
||[crossReferenceAutoText](/.indexmarkallentriesoptions#word-javascript/api/word/-indexmarkallentriesoptions-crossreferenceautotext-member)|If provided, specifies the name of the `AutoText` entry that contains the text for a cross-reference (if this property is specified, `crossReference` is ignored).|
||[entry](/.indexmarkallentriesoptions#word-javascript/api/word/-indexmarkallentriesoptions-entry-member)|If provided, specifies the text you want to appear in the index, in the form `MainEntry[:Subentry]`.|
||[entryAutoText](/.indexmarkallentriesoptions#word-javascript/api/word/-indexmarkallentriesoptions-entryautotext-member)|If provided, specifies the `AutoText` entry that contains the text you want to appear in the index (if this property is specified, `entry` is ignored).|
||[italic](/.indexmarkallentriesoptions#word-javascript/api/word/-indexmarkallentriesoptions-italic-member)|If provided, specifies whether to add italic formatting to page numbers for index entries.|
|[IndexMarkEntryOptions](/.indexmarkentryoptions)|[bold](/.indexmarkentryoptions#word-javascript/api/word/-indexmarkentryoptions-bold-member)|If provided, specifies whether to add bold formatting to page numbers for index entries.|
||[bookmarkName](/.indexmarkentryoptions#word-javascript/api/word/-indexmarkentryoptions-bookmarkname-member)|If provided, specifies the bookmark name that marks the range of pages you want to appear in the index.|
||[crossReference](/.indexmarkentryoptions#word-javascript/api/word/-indexmarkentryoptions-crossreference-member)|If provided, specifies the cross-reference that will appear in the index.|
||[crossReferenceAutoText](/.indexmarkentryoptions#word-javascript/api/word/-indexmarkentryoptions-crossreferenceautotext-member)|If provided, specifies the name of the `AutoText` entry that contains the text for a cross-reference (if this property is specified, `crossReference` is ignored).|
||[entry](/.indexmarkentryoptions#word-javascript/api/word/-indexmarkentryoptions-entry-member)|If provided, specifies the text you want to appear in the index, in the form `MainEntry[:Subentry]`.|
||[entryAutoText](/.indexmarkentryoptions#word-javascript/api/word/-indexmarkentryoptions-entryautotext-member)|If provided, specifies the `AutoText` entry that contains the text you want to appear in the index (if this property is specified, `entry` is ignored).|
||[italic](/.indexmarkentryoptions#word-javascript/api/word/-indexmarkentryoptions-italic-member)|If provided, specifies whether to add italic formatting to page numbers for index entries.|
||[reading](/.indexmarkentryoptions#word-javascript/api/word/-indexmarkentryoptions-reading-member)|If provided, specifies whether to show an index entry in the right location when indexes are sorted phonetically (East Asian languages only).|
|[LineFormat](/.lineformat)|[backgroundColor](/.lineformat#word-javascript/api/word/-lineformat-backgroundcolor-member)|Gets a `ColorFormat` object that represents the background color for a patterned line.|
||[beginArrowheadLength](/.lineformat#word-javascript/api/word/-lineformat-beginarrowheadlength-member)|Specifies the length of the arrowhead at the beginning of the line.|
||[beginArrowheadStyle](/.lineformat#word-javascript/api/word/-lineformat-beginarrowheadstyle-member)|Specifies the style of the arrowhead at the beginning of the line.|
||[beginArrowheadWidth](/.lineformat#word-javascript/api/word/-lineformat-beginarrowheadwidth-member)|Specifies the width of the arrowhead at the beginning of the line.|
||[dashStyle](/.lineformat#word-javascript/api/word/-lineformat-dashstyle-member)|Specifies the dash style for the line.|
||[endArrowheadLength](/.lineformat#word-javascript/api/word/-lineformat-endarrowheadlength-member)|Specifies the length of the arrowhead at the end of the line.|
||[endArrowheadStyle](/.lineformat#word-javascript/api/word/-lineformat-endarrowheadstyle-member)|Specifies the style of the arrowhead at the end of the line.|
||[endArrowheadWidth](/.lineformat#word-javascript/api/word/-lineformat-endarrowheadwidth-member)|Specifies the width of the arrowhead at the end of the line.|
||[foregroundColor](/.lineformat#word-javascript/api/word/-lineformat-foregroundcolor-member)|Gets a `ColorFormat` object that represents the foreground color for the line.|
||[insetPen](/.lineformat#word-javascript/api/word/-lineformat-insetpen-member)|Specifies if to draw lines inside a shape.|
||[isVisible](/.lineformat#word-javascript/api/word/-lineformat-isvisible-member)|Specifies if the object, or the formatting applied to it, is visible.|
||[pattern](/.lineformat#word-javascript/api/word/-lineformat-pattern-member)|Specifies the pattern applied to the line.|
||[style](/.lineformat#word-javascript/api/word/-lineformat-style-member)|Specifies the line format style.|
||[transparency](/.lineformat#word-javascript/api/word/-lineformat-transparency-member)|Specifies the degree of transparency of the line as a value between 0.0 (opaque) and 1.0 (clear).|
||[weight](/.lineformat#word-javascript/api/word/-lineformat-weight-member)|Specifies the thickness of the line in points.|
|[LineNumbering](/.linenumbering)|[countBy](/.linenumbering#word-javascript/api/word/-linenumbering-countby-member)|Specifies the numeric increment for line numbers.|
||[distanceFromText](/.linenumbering#word-javascript/api/word/-linenumbering-distancefromtext-member)|Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.|
||[isActive](/.linenumbering#word-javascript/api/word/-linenumbering-isactive-member)|Specifies if line numbering is active for the specified document, section, or sections.|
||[restartMode](/.linenumbering#word-javascript/api/word/-linenumbering-restartmode-member)|Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.|
||[startingNumber](/.linenumbering#word-javascript/api/word/-linenumbering-startingnumber-member)|Specifies the starting line number.|
|[LinkFormat](/.linkformat)|[breakLink()](/.linkformat#word-javascript/api/word/-linkformat-breaklink-member(1))|Breaks the link between the source file and the OLE object, picture, or linked field.|
||[isAutoUpdated](/.linkformat#word-javascript/api/word/-linkformat-isautoupdated-member)|Specifies if the link is updated automatically when the container file is opened or when the source file is changed.|
||[isLocked](/.linkformat#word-javascript/api/word/-linkformat-islocked-member)|Specifies if a `Field`, `InlineShape`, or `Shape` object is locked to prevent automatic updating.|
||[isPictureSavedWithDocument](/.linkformat#word-javascript/api/word/-linkformat-ispicturesavedwithdocument-member)|Specifies if the linked picture is saved with the document.|
||[sourceFullName](/.linkformat#word-javascript/api/word/-linkformat-sourcefullname-member)|Specifies the path and name of the source file for the linked OLE object, picture, or field.|
||[sourceName](/.linkformat#word-javascript/api/word/-linkformat-sourcename-member)|Gets the name of the source file for the linked OLE object, picture, or field.|
||[sourcePath](/.linkformat#word-javascript/api/word/-linkformat-sourcepath-member)|Gets the path of the source file for the linked OLE object, picture, or field.|
||[type](/.linkformat#word-javascript/api/word/-linkformat-type-member)|Gets the link type.|
|[ListFormat](/.listformat)|[applyBulletDefault(defaultListBehavior: Word.DefaultListBehavior)](/.listformat#word-javascript/api/word/-listformat-applybulletdefault-member(1))|Adds bullets and formatting to the paragraphs in the range.|
||[applyListTemplateWithLevel(listTemplate: Word.ListTemplate, options?: Word.ListTemplateApplyOptions)](/.listformat#word-javascript/api/word/-listformat-applylisttemplatewithlevel-member(1))|Applies a list template with a specific level to the paragraphs in the range.|
||[applyNumberDefault(defaultListBehavior: Word.DefaultListBehavior)](/.listformat#word-javascript/api/word/-listformat-applynumberdefault-member(1))|Adds numbering and formatting to the paragraphs in the range.|
||[applyOutlineNumberDefault(defaultListBehavior: Word.DefaultListBehavior)](/.listformat#word-javascript/api/word/-listformat-applyoutlinenumberdefault-member(1))|Adds outline numbering and formatting to the paragraphs in the range.|
||[canContinuePreviousList(listTemplate: Word.ListTemplate)](/.listformat#word-javascript/api/word/-listformat-cancontinuepreviouslist-member(1))|Determines whether the `ListFormat` object can continue a previous list.|
||[convertNumbersToText(numberType: Word.NumberType)](/.listformat#word-javascript/api/word/-listformat-convertnumberstotext-member(1))|Converts numbers in the list to plain text.|
||[countNumberedItems(options?: Word.ListFormatCountNumberedItemsOptions)](/.listformat#word-javascript/api/word/-listformat-countnumbereditems-member(1))|Counts the numbered items in the list.|
||[isSingleList](/.listformat#word-javascript/api/word/-listformat-issinglelist-member)|Indicates whether the `ListFormat` object contains a single list.|
||[isSingleListTemplate](/.listformat#word-javascript/api/word/-listformat-issinglelisttemplate-member)|Indicates whether the `ListFormat` object contains a single list template.|
||[list](/.listformat#word-javascript/api/word/-listformat-list-member)|Returns a `List` object that represents the first formatted list contained in the `ListFormat` object.|
||[listIndent()](/.listformat#word-javascript/api/word/-listformat-listindent-member(1))|Indents the list by one level.|
||[listLevelNumber](/.listformat#word-javascript/api/word/-listformat-listlevelnumber-member)|Specifies the list level number for the first paragraph for the `ListFormat` object.|
||[listOutdent()](/.listformat#word-javascript/api/word/-listformat-listoutdent-member(1))|Outdents the list by one level.|
||[listString](/.listformat#word-javascript/api/word/-listformat-liststring-member)|Gets the string representation of the list value of the first paragraph in the range for the `ListFormat` object.|
||[listTemplate](/.listformat#word-javascript/api/word/-listformat-listtemplate-member)|Gets the list template associated with the `ListFormat` object.|
||[listType](/.listformat#word-javascript/api/word/-listformat-listtype-member)|Gets the type of the list for the `ListFormat` object.|
||[listValue](/.listformat#word-javascript/api/word/-listformat-listvalue-member)|Gets the numeric value of the the first paragraph in the range for the `ListFormat` object.|
||[removeNumbers(numberType: Word.NumberType)](/.listformat#word-javascript/api/word/-listformat-removenumbers-member(1))|Removes numbering from the list.|
|[ListFormatCountNumberedItemsOptions](/.listformatcountnumbereditemsoptions)|[level](/.listformatcountnumbereditemsoptions#word-javascript/api/word/-listformatcountnumbereditemsoptions-level-member)|If provided, specifies the level to count.|
||[numberType](/.listformatcountnumbereditemsoptions#word-javascript/api/word/-listformatcountnumbereditemsoptions-numbertype-member)|If provided, specifies the type of number to count.|
|[ListTemplateApplyOptions](/.listtemplateapplyoptions)|[applyLevel](/.listtemplateapplyoptions#word-javascript/api/word/-listtemplateapplyoptions-applylevel-member)|If provided, specifies the level to apply in the list template.|
||[applyTo](/.listtemplateapplyoptions#word-javascript/api/word/-listtemplateapplyoptions-applyto-member)|If provided, specifies which part of the list to apply the template to.|
||[continuePreviousList](/.listtemplateapplyoptions#word-javascript/api/word/-listtemplateapplyoptions-continuepreviouslist-member)|If provided, specifies whether to continue the previous list.|
||[defaultListBehavior](/.listtemplateapplyoptions#word-javascript/api/word/-listtemplateapplyoptions-defaultlistbehavior-member)|If provided, specifies the default list behavior.|
|[OleFormat](/.oleformat)|[activate()](/.oleformat#word-javascript/api/word/-oleformat-activate-member(1))|Activates the `OleFormat` object.|
||[activateAs(classType: string)](/.oleformat#word-javascript/api/word/-oleformat-activateas-member(1))|Sets the Windows registry value that determines the default application used to activate the specified OLE object.|
||[classType](/.oleformat#word-javascript/api/word/-oleformat-classtype-member)|Specifies the class type for the specified OLE object, picture, or field.|
||[doVerb(verbIndex: Word.OleVerb)](/.oleformat#word-javascript/api/word/-oleformat-doverb-member(1))|Requests that the OLE object perform one of its available verbs.|
||[edit()](/.oleformat#word-javascript/api/word/-oleformat-edit-member(1))|Opens the OLE object for editing in the application it was created in.|
||[iconIndex](/.oleformat#word-javascript/api/word/-oleformat-iconindex-member)|Specifies the icon that is used when the `displayAsIcon` property is `true`.|
||[iconLabel](/.oleformat#word-javascript/api/word/-oleformat-iconlabel-member)|Specifies the text displayed below the icon for the OLE object.|
||[iconName](/.oleformat#word-javascript/api/word/-oleformat-iconname-member)|Specifies the program file in which the icon for the OLE object is stored.|
||[iconPath](/.oleformat#word-javascript/api/word/-oleformat-iconpath-member)|Gets the path of the file in which the icon for the OLE object is stored.|
||[isDisplayedAsIcon](/.oleformat#word-javascript/api/word/-oleformat-isdisplayedasicon-member)|Gets whether the specified object is displayed as an icon.|
||[isFormattingPreservedOnUpdate](/.oleformat#word-javascript/api/word/-oleformat-isformattingpreservedonupdate-member)|Specifies whether formatting done in Microsoft Word to the linked OLE object is preserved.|
||[label](/.oleformat#word-javascript/api/word/-oleformat-label-member)|Gets a string that's used to identify the portion of the source file that's being linked.|
||[open()](/.oleformat#word-javascript/api/word/-oleformat-open-member(1))|Opens the `OleFormat` object.|
||[progID](/.oleformat#word-javascript/api/word/-oleformat-progid-member)|Gets the programmatic identifier (`ProgId`) for the specified OLE object.|
|[Page](/.page)|[breaks](/.page#word-javascript/api/word/-page-breaks-member)|Gets a `BreakCollection` object that represents the breaks on the page.|
|[PageSetup](/.pagesetup)|[bookFoldPrinting](/.pagesetup#word-javascript/api/word/-pagesetup-bookfoldprinting-member)|Specifies whether Microsoft Word prints the document as a booklet.|
||[bookFoldPrintingSheets](/.pagesetup#word-javascript/api/word/-pagesetup-bookfoldprintingsheets-member)|Specifies the number of pages for each booklet.|
||[bookFoldReversePrinting](/.pagesetup#word-javascript/api/word/-pagesetup-bookfoldreverseprinting-member)|Specifies if Microsoft Word reverses the printing order for book fold printing of bidirectional or Asian language documents.|
||[bottomMargin](/.pagesetup#word-javascript/api/word/-pagesetup-bottommargin-member)|Specifies the distance (in points) between the bottom edge of the page and the bottom boundary of the body text.|
||[charsLine](/.pagesetup#word-javascript/api/word/-pagesetup-charsline-member)|Specifies the number of characters per line in the document grid.|
||[differentFirstPageHeaderFooter](/.pagesetup#word-javascript/api/word/-pagesetup-differentfirstpageheaderfooter-member)|Specifies whether the first page has a different header and footer.|
||[footerDistance](/.pagesetup#word-javascript/api/word/-pagesetup-footerdistance-member)|Specifies the distance between the footer and the bottom of the page in points.|
||[gutter](/.pagesetup#word-javascript/api/word/-pagesetup-gutter-member)|Specifies the amount (in points) of extra margin space added to each page in a document or section for binding.|
||[gutterPosition](/.pagesetup#word-javascript/api/word/-pagesetup-gutterposition-member)|Specifies on which side the gutter appears in a document.|
||[gutterStyle](/.pagesetup#word-javascript/api/word/-pagesetup-gutterstyle-member)|Specifies whether Microsoft Word uses gutters for the current document based on a right-to-left language or a left-to-right language.|
||[headerDistance](/.pagesetup#word-javascript/api/word/-pagesetup-headerdistance-member)|Specifies the distance between the header and the top of the page in points.|
||[layoutMode](/.pagesetup#word-javascript/api/word/-pagesetup-layoutmode-member)|Specifies the layout mode for the current document.|
||[leftMargin](/.pagesetup#word-javascript/api/word/-pagesetup-leftmargin-member)|Specifies the distance (in points) between the left edge of the page and the left boundary of the body text.|
||[lineNumbering](/.pagesetup#word-javascript/api/word/-pagesetup-linenumbering-member)|Specifies a `LineNumbering` object that represents the line numbers for the `PageSetup` object.|
||[linesPage](/.pagesetup#word-javascript/api/word/-pagesetup-linespage-member)|Specifies the number of lines per page in the document grid.|
||[mirrorMargins](/.pagesetup#word-javascript/api/word/-pagesetup-mirrormargins-member)|Specifies if the inside and outside margins of facing pages are the same width.|
||[oddAndEvenPagesHeaderFooter](/.pagesetup#word-javascript/api/word/-pagesetup-oddandevenpagesheaderfooter-member)|Specifies whether odd and even pages have different headers and footers.|
||[orientation](/.pagesetup#word-javascript/api/word/-pagesetup-orientation-member)|Specifies the orientation of the page.|
||[pageHeight](/.pagesetup#word-javascript/api/word/-pagesetup-pageheight-member)|Specifies the page height in points.|
||[pageWidth](/.pagesetup#word-javascript/api/word/-pagesetup-pagewidth-member)|Specifies the page width in points.|
||[paperSize](/.pagesetup#word-javascript/api/word/-pagesetup-papersize-member)|Specifies the paper size of the page.|
||[rightMargin](/.pagesetup#word-javascript/api/word/-pagesetup-rightmargin-member)|Specifies the distance (in points) between the right edge of the page and the right boundary of the body text.|
||[sectionDirection](/.pagesetup#word-javascript/api/word/-pagesetup-sectiondirection-member)|Specifies the reading order and alignment for the specified sections.|
||[sectionStart](/.pagesetup#word-javascript/api/word/-pagesetup-sectionstart-member)|Specifies the type of section break for the specified object.|
||[setAsTemplateDefault()](/.pagesetup#word-javascript/api/word/-pagesetup-setastemplatedefault-member(1))|Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.|
||[showGrid](/.pagesetup#word-javascript/api/word/-pagesetup-showgrid-member)|Specifies whether to show the grid.|
||[suppressEndnotes](/.pagesetup#word-javascript/api/word/-pagesetup-suppressendnotes-member)|Specifies if endnotes are printed at the end of the next section that doesn't suppress endnotes.|
||[textColumns](/.pagesetup#word-javascript/api/word/-pagesetup-textcolumns-member)|Gets a `TextColumnCollection` object that represents the set of text columns for the `PageSetup` object.|
||[togglePortrait()](/.pagesetup#word-javascript/api/word/-pagesetup-toggleportrait-member(1))|Switches between portrait and landscape page orientations for a document or section.|
||[topMargin](/.pagesetup#word-javascript/api/word/-pagesetup-topmargin-member)|Specifies the top margin of the page in points.|
||[twoPagesOnOne](/.pagesetup#word-javascript/api/word/-pagesetup-twopagesonone-member)|Specifies whether to print two pages per sheet.|
||[verticalAlignment](/.pagesetup#word-javascript/api/word/-pagesetup-verticalalignment-member)|Specifies the vertical alignment of text on each page in a document or section.|
|[Paragraph](/.paragraph)|[borders](/.paragraph#word-javascript/api/word/-paragraph-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the paragraph.|
||[closeUp()](/.paragraph#word-javascript/api/word/-paragraph-closeup-member(1))|Removes any spacing before the paragraph.|
||[indent()](/.paragraph#word-javascript/api/word/-paragraph-indent-member(1))|Indents the paragraph by one level.|
||[indentCharacterWidth(count: number)](/.paragraph#word-javascript/api/word/-paragraph-indentcharacterwidth-member(1))|Indents the paragraph by a specified number of characters.|
||[indentFirstLineCharacterWidth(count: number)](/.paragraph#word-javascript/api/word/-paragraph-indentfirstlinecharacterwidth-member(1))|Indents the first line of the paragraph by the specified number of characters.|
||[joinList()](/.paragraph#word-javascript/api/word/-paragraph-joinlist-member(1))|Joins a list paragraph with the closest list above or below this paragraph.|
||[next(count: number)](/.paragraph#word-javascript/api/word/-paragraph-next-member(1))|Returns a `Paragraph` object that represents the next paragraph.|
||[onCommentAdded](/.paragraph#word-javascript/api/word/-paragraph-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/.paragraph#word-javascript/api/word/-paragraph-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/.paragraph#word-javascript/api/word/-paragraph-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/.paragraph#word-javascript/api/word/-paragraph-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/.paragraph#word-javascript/api/word/-paragraph-oncommentselected-member)|Occurs when a comment is selected.|
||[openOrCloseUp()](/.paragraph#word-javascript/api/word/-paragraph-openorcloseup-member(1))|Toggles the spacing before the paragraph.|
||[openUp()](/.paragraph#word-javascript/api/word/-paragraph-openup-member(1))|Sets spacing before the paragraph to 12 points.|
||[outdent()](/.paragraph#word-javascript/api/word/-paragraph-outdent-member(1))|Removes one level of indent for the paragraph.|
||[outlineDemote()](/.paragraph#word-javascript/api/word/-paragraph-outlinedemote-member(1))|Applies the next heading level style (Heading 1 through Heading 8) to the paragraph.|
||[outlineDemoteToBody()](/.paragraph#word-javascript/api/word/-paragraph-outlinedemotetobody-member(1))|Demotes the paragraph to body text by applying the Normal style.|
||[outlinePromote()](/.paragraph#word-javascript/api/word/-paragraph-outlinepromote-member(1))|Applies the previous heading level style (Heading 1 through Heading 8) to the paragraph.|
||[previous(count: number)](/.paragraph#word-javascript/api/word/-paragraph-previous-member(1))|Returns the previous paragraph as a `Paragraph` object.|
||[range](/.paragraph#word-javascript/api/word/-paragraph-range-member)|Gets a `Range` object that represents the portion of the document that's contained within the paragraph.|
||[reset()](/.paragraph#word-javascript/api/word/-paragraph-reset-member(1))|Removes manual paragraph formatting (formatting not applied using a style).|
||[resetAdvanceTo()](/.paragraph#word-javascript/api/word/-paragraph-resetadvanceto-member(1))|Resets the paragraph that uses custom list levels to the original level settings.|
||[selectNumber()](/.paragraph#word-javascript/api/word/-paragraph-selectnumber-member(1))|Selects the number or bullet in a list.|
||[separateList()](/.paragraph#word-javascript/api/word/-paragraph-separatelist-member(1))|Separates a list into two separate lists.|
||[shading](/.paragraph#word-javascript/api/word/-paragraph-shading-member)|Returns a `ShadingUniversal` object that refers to the shading formatting for the paragraph.|
||[space1()](/.paragraph#word-javascript/api/word/-paragraph-space1-member(1))|Sets the paragraph to single spacing.|
||[space1Pt5()](/.paragraph#word-javascript/api/word/-paragraph-space1pt5-member(1))|Sets the paragraph to 1.5-line spacing.|
||[space2()](/.paragraph#word-javascript/api/word/-paragraph-space2-member(1))|Sets the paragraph to double spacing.|
||[tabHangingIndent(count: number)](/.paragraph#word-javascript/api/word/-paragraph-tabhangingindent-member(1))|Sets a hanging indent to a specified number of tab stops.|
||[tabIndent(count: number)](/.paragraph#word-javascript/api/word/-paragraph-tabindent-member(1))|Sets the left indent for the paragraph to a specified number of tab stops.|
|[ParagraphAddedEventArgs](/.paragraphaddedeventargs)|[type](/.paragraphaddedeventargs#word-javascript/api/word/-paragraphaddedeventargs-type-member)|The event type.|
|[ParagraphChangedEventArgs](/.paragraphchangedeventargs)|[type](/.paragraphchangedeventargs#word-javascript/api/word/-paragraphchangedeventargs-type-member)|The event type.|
|[ParagraphCollection](/.paragraphcollection)|[add(range: Word.Range)](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-add-member(1))|Returns a `Paragraph` object that represents a new, blank paragraph added to the document.|
||[closeUp()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-closeup-member(1))|Removes any spacing before the specified paragraphs.|
||[decreaseSpacing()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-decreasespacing-member(1))|Decreases the spacing before and after paragraphs in six-point increments.|
||[increaseSpacing()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-increasespacing-member(1))|Increases the spacing before and after paragraphs in six-point increments.|
||[indent()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-indent-member(1))|Indents the paragraphs by one level.|
||[indentCharacterWidth(count: number)](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-indentcharacterwidth-member(1))|Indents the paragraphs in the collection by the specified number of characters.|
||[indentFirstLineCharacterWidth(count: number)](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-indentfirstlinecharacterwidth-member(1))|Indents the first line of the paragraphs in the collection by the specified number of characters.|
||[openOrCloseUp()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-openorcloseup-member(1))|Toggles spacing before paragraphs.|
||[openUp()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-openup-member(1))|Sets spacing before the specified paragraphs to 12 points.|
||[outdent()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-outdent-member(1))|Removes one level of indent for the paragraphs.|
||[outlineDemote()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-outlinedemote-member(1))|Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraphs.|
||[outlineDemoteToBody()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-outlinedemotetobody-member(1))|Demotes the specified paragraphs to body text by applying the Normal style.|
||[outlinePromote()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-outlinepromote-member(1))|Applies the previous heading level style (Heading 1 through Heading 8) to the paragraphs in the collection.|
||[space1()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-space1-member(1))|Sets the specified paragraphs to single spacing.|
||[space1Pt5()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-space1pt5-member(1))|Sets the specified paragraphs to 1.5-line spacing.|
||[space2()](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-space2-member(1))|Sets the specified paragraphs to double spacing.|
||[tabHangingIndent(count: number)](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-tabhangingindent-member(1))|Sets a hanging indent to the specified number of tab stops.|
||[tabIndent(count: number)](/.paragraphcollection#word-javascript/api/word/-paragraphcollection-tabindent-member(1))|Sets the left indent for the specified paragraphs to the specified number of tab stops.|
|[ParagraphDeletedEventArgs](/.paragraphdeletedeventargs)|[type](/.paragraphdeletedeventargs#word-javascript/api/word/-paragraphdeletedeventargs-type-member)|The event type.|
|[PictureContentControl](/.picturecontentcontrol)|[appearance](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[color](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-color-member)|Specifies the red-green-blue (RGB) value of the color of the content control.|
||[copy()](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-copy-member(1))|Copies the content control from the active document to the Clipboard.|
||[cut()](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-cut-member(1))|Removes the content control from the active document and moves the content control to the Clipboard.|
||[delete(deleteContents?: boolean)](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-delete-member(1))|Deletes the content control and optionally its contents.|
||[id](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-id-member)|Returns the identification for the content control.|
||[isTemporary](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-istemporary-member)|Specifies whether to remove the content control from the active document when the user edits the contents of the control.|
||[level](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-level-member)|Returns the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.|
||[lockContentControl](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-lockcontentcontrol-member)|Specifies if the content control is locked (can't be deleted).|
||[lockContents](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-lockcontents-member)|Specifies if the contents of the content control are locked (not editable).|
||[placeholderText](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-placeholdertext-member)|Returns a `BuildingBlock` object that represents the placeholder text for the content control.|
||[range](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-range-member)|Returns a `Range` object that represents the contents of the content control in the active document.|
||[setPlaceholderText(options?: Word.ContentControlPlaceholderOptions)](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-setplaceholdertext-member(1))|Sets the placeholder text that displays in the content control until a user enters their own text.|
||[showingPlaceholderText](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-showingplaceholdertext-member)|Returns whether the placeholder text for the content control is being displayed.|
||[tag](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-tag-member)|Specifies a tag to identify the content control.|
||[title](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-title-member)|Specifies the title for the content control.|
||[xmlMapping](/.picturecontentcontrol#word-javascript/api/word/-picturecontentcontrol-xmlmapping-member)|Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[Range](/.range)|[bold](/.range#word-javascript/api/word/-range-bold-member)|Specifies whether the range is formatted as bold.|
||[boldBidirectional](/.range#word-javascript/api/word/-range-boldbidirectional-member)|Specifies whether the range is formatted as bold in a right-to-left language document.|
||[bookmarks](/.range#word-javascript/api/word/-range-bookmarks-member)|Returns a `BookmarkCollection` object that represents all the bookmarks in the range.|
||[borders](/.range#word-javascript/api/word/-range-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the range.|
||[case](/.range#word-javascript/api/word/-range-case-member)|Specifies a `CharacterCase` value that represents the case of the text in the range.|
||[characterWidth](/.range#word-javascript/api/word/-range-characterwidth-member)|Specifies the character width of the range.|
||[combineCharacters](/.range#word-javascript/api/word/-range-combinecharacters-member)|Specifies if the range contains combined characters.|
||[detectLanguage()](/.range#word-javascript/api/word/-range-detectlanguage-member(1))|Analyzes the range text to determine the language that it's written in.|
||[disableCharacterSpaceGrid](/.range#word-javascript/api/word/-range-disablecharacterspacegrid-member)|Specifies if Microsoft Word ignores the number of characters per line for the corresponding `Range` object.|
||[emphasisMark](/.range#word-javascript/api/word/-range-emphasismark-member)|Specifies the emphasis mark for a character or designated character string.|
||[end](/.range#word-javascript/api/word/-range-end-member)|Specifies the ending character position of the range.|
||[fitTextWidth](/.range#word-javascript/api/word/-range-fittextwidth-member)|Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range.|
||[frames](/.range#word-javascript/api/word/-range-frames-member)|Gets a `FrameCollection` object that represents all the frames in the range.|
||[grammarChecked](/.range#word-javascript/api/word/-range-grammarchecked-member)|Specifies if a grammar check has been run on the range or document.|
||[hasNoProofing](/.range#word-javascript/api/word/-range-hasnoproofing-member)|Specifies the proofing status (spelling and grammar checking) of the range.|
||[highlightColorIndex](/.range#word-javascript/api/word/-range-highlightcolorindex-member)|Specifies the highlight color for the range.|
||[horizontalInVertical](/.range#word-javascript/api/word/-range-horizontalinvertical-member)|Specifies the formatting for horizontal text set within vertical text.|
||[hyperlinks](/.range#word-javascript/api/word/-range-hyperlinks-member)|Returns a `HyperlinkCollection` object that represents all the hyperlinks in the range.|
||[id](/.range#word-javascript/api/word/-range-id-member)|Specifies the ID for the range.|
||[isEndOfRowMark](/.range#word-javascript/api/word/-range-isendofrowmark-member)|Gets if the range is collapsed and is located at the end-of-row mark in a table.|
||[isTextVisibleOnScreen](/.range#word-javascript/api/word/-range-istextvisibleonscreen-member)|Gets whether the text in the range is visible on the screen.|
||[italic](/.range#word-javascript/api/word/-range-italic-member)|Specifies if the font or range is formatted as italic.|
||[italicBidirectional](/.range#word-javascript/api/word/-range-italicbidirectional-member)|Specifies if the font or range is formatted as italic (right-to-left languages).|
||[kana](/.range#word-javascript/api/word/-range-kana-member)|Specifies whether the range of Japanese language text is hiragana or katakana.|
||[languageDetected](/.range#word-javascript/api/word/-range-languagedetected-member)|Specifies whether Microsoft Word has detected the language of the text in the range.|
||[languageId](/.range#word-javascript/api/word/-range-languageid-member)|Specifies a `LanguageId` value that represents the language for the range.|
||[languageIdFarEast](/.range#word-javascript/api/word/-range-languageidfareast-member)|Specifies an East Asian language for the range.|
||[languageIdOther](/.range#word-javascript/api/word/-range-languageidother-member)|Specifies a language for the range that isn't classified as an East Asian language.|
||[listFormat](/.range#word-javascript/api/word/-range-listformat-member)|Returns a `ListFormat` object that represents all the list formatting characteristics of the range.|
||[onCommentAdded](/.range#word-javascript/api/word/-range-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/.range#word-javascript/api/word/-range-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/.range#word-javascript/api/word/-range-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/.range#word-javascript/api/word/-range-oncommentselected-member)|Occurs when a comment is selected.|
||[sections](/.range#word-javascript/api/word/-range-sections-member)|Gets the collection of sections in the range.|
||[shading](/.range#word-javascript/api/word/-range-shading-member)|Returns a `ShadingUniversal` object that refers to the shading formatting for the range.|
||[showAll](/.range#word-javascript/api/word/-range-showall-member)|Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed.|
||[spellingChecked](/.range#word-javascript/api/word/-range-spellingchecked-member)|Specifies if spelling has been checked throughout the range or document.|
||[start](/.range#word-javascript/api/word/-range-start-member)|Specifies the starting character position of the range.|
||[storyLength](/.range#word-javascript/api/word/-range-storylength-member)|Gets the number of characters in the story that contains the range.|
||[storyType](/.range#word-javascript/api/word/-range-storytype-member)|Gets the story type for the range.|
||[tableColumns](/.range#word-javascript/api/word/-range-tablecolumns-member)|Gets a `TableColumnCollection` object that represents all the table columns in the range.|
||[twoLinesInOne](/.range#word-javascript/api/word/-range-twolinesinone-member)|Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.|
||[underline](/.range#word-javascript/api/word/-range-underline-member)|Specifies the type of underline applied to the range.|
|[ReflectionFormat](/.reflectionformat)|[blur](/.reflectionformat#word-javascript/api/word/-reflectionformat-blur-member)|Specifies the degree of blur effect applied to the `ReflectionFormat` object as a value between 0.0 and 100.0.|
||[offset](/.reflectionformat#word-javascript/api/word/-reflectionformat-offset-member)|Specifies the amount of separation, in points, of the reflected image from the shape.|
||[size](/.reflectionformat#word-javascript/api/word/-reflectionformat-size-member)|Specifies the size of the reflection as a percentage of the reflected shape from 0 to 100.|
||[transparency](/.reflectionformat#word-javascript/api/word/-reflectionformat-transparency-member)|Specifies the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear).|
||[type](/.reflectionformat#word-javascript/api/word/-reflectionformat-type-member)|Specifies a `ReflectionType` value that represents the type and direction of the lighting for a shape reflection.|
|[RepeatingSectionContentControl](/.repeatingsectioncontentcontrol)|[allowInsertDeleteSection](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-allowinsertdeletesection-member)|Specifies whether users can add or remove sections from this repeating section content control by using the user interface.|
||[appearance](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-appearance-member)|Specifies the appearance of the content control.|
||[color](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-color-member)|Specifies the red-green-blue (RGB) value of the color of the content control.|
||[copy()](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-copy-member(1))|Copies the content control from the active document to the Clipboard.|
||[cut()](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-cut-member(1))|Removes the content control from the active document and moves the content control to the Clipboard.|
||[delete(deleteContents?: boolean)](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-delete-member(1))|Deletes the content control and the contents of the content control.|
||[id](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-id-member)|Returns the identification for the content control.|
||[isTemporary](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-istemporary-member)|Specifies whether to remove the content control from the active document when the user edits the contents of the control.|
||[level](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-level-member)|Returns the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.|
||[lockContentControl](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-lockcontentcontrol-member)|Specifies if the content control is locked (can't be deleted).|
||[lockContents](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-lockcontents-member)|Specifies if the contents of the content control are locked (not editable).|
||[placeholderText](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-placeholdertext-member)|Returns a `BuildingBlock` object that represents the placeholder text for the content control.|
||[range](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-range-member)|Gets a `Range` object that represents the contents of the content control in the active document.|
||[repeatingSectionItemTitle](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-repeatingsectionitemtitle-member)|Specifies the name of the repeating section items used in the context menu associated with this repeating section content control.|
||[repeatingSectionItems](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-repeatingsectionitems-member)|Returns the collection of repeating section items in this repeating section content control.|
||[setPlaceholderText(options?: Word.ContentControlPlaceholderOptions)](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-setplaceholdertext-member(1))|Sets the placeholder text that displays in the content control until a user enters their own text.|
||[showingPlaceholderText](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-showingplaceholdertext-member)|Returns whether the placeholder text for the content control is being displayed.|
||[tag](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-tag-member)|Specifies a tag to identify the content control.|
||[title](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-title-member)|Specifies the title for the content control.|
||[xmlapping](/.repeatingsectioncontentcontrol#word-javascript/api/word/-repeatingsectioncontentcontrol-xmlapping-member)|Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[RepeatingSectionItem](/.repeatingsectionitem)|[delete()](/.repeatingsectionitem#word-javascript/api/word/-repeatingsectionitem-delete-member(1))|Deletes this `RepeatingSectionItem` object.|
||[insertItemAfter()](/.repeatingsectionitem#word-javascript/api/word/-repeatingsectionitem-insertitemafter-member(1))|Adds a repeating section item after this item and returns the new item.|
||[insertItemBefore()](/.repeatingsectionitem#word-javascript/api/word/-repeatingsectionitem-insertitembefore-member(1))|Adds a repeating section item before this item and returns the new item.|
||[range](/.repeatingsectionitem#word-javascript/api/word/-repeatingsectionitem-range-member)|Returns the range of this repeating section item, excluding the start and end tags.|
|[RepeatingSectionItemCollection](/.repeatingsectionitemcollection)|[getItemAt(index: number)](/.repeatingsectionitemcollection#word-javascript/api/word/-repeatingsectionitemcollection-getitemat-member(1))|Returns an individual repeating section item.|
|[Reviewer](/.reviewer)|[isVisible](/.reviewer#word-javascript/api/word/-reviewer-isvisible-member)|Specifies if the `Reviewer` object is visible.|
|[ReviewerCollection](/.reviewercollection)|[getItem(index: number)](/.reviewercollection#word-javascript/api/word/-reviewercollection-getitem-member(1))|Returns a `Reviewer` object that represents the specified item in the collection.|
||[items](/.reviewercollection#word-javascript/api/word/-reviewercollection-items-member)|Gets the loaded child items in this collection.|
|[RevisionsFilter](/.revisionsfilter)|[markup](/.revisionsfilter#word-javascript/api/word/-revisionsfilter-markup-member)|Specifies a `RevisionsMarkup` value that represents the extent of reviewer markup displayed in the document.|
||[reviewers](/.revisionsfilter#word-javascript/api/word/-revisionsfilter-reviewers-member)|Gets the `ReviewerCollection` object that represents the collection of reviewers of one or more documents.|
||[toggleShowAllReviewers()](/.revisionsfilter#word-javascript/api/word/-revisionsfilter-toggleshowallreviewers-member(1))|Shows or hides all revisions in the document that contain comments and tracked changes.|
||[view](/.revisionsfilter#word-javascript/api/word/-revisionsfilter-view-member)|Specifies a `RevisionsView` value that represents globally whether Word displays the original version of the document or the final version, which might have revisions and formatting changes applied.|
|[Section](/.section)|[borders](/.section#word-javascript/api/word/-section-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders in the section.|
||[pageSetup](/.section#word-javascript/api/word/-section-pagesetup-member)|Returns a `PageSetup` object that's associated with the section.|
||[protectedForForms](/.section#word-javascript/api/word/-section-protectedforforms-member)|Specifies if the section is protected for forms.|
|[ShadingUniversal](/.shadinguniversal)|[backgroundPatternColor](/.shadinguniversal#word-javascript/api/word/-shadinguniversal-backgroundpatterncolor-member)|Specifies the color that's applied to the background of the `ShadingUniversal` object.|
||[backgroundPatternColorIndex](/.shadinguniversal#word-javascript/api/word/-shadinguniversal-backgroundpatterncolorindex-member)|Specifies the color that's applied to the background of the `ShadingUniversal` object.|
||[foregroundPatternColor](/.shadinguniversal#word-javascript/api/word/-shadinguniversal-foregroundpatterncolor-member)|Specifies the color that's applied to the foreground of the `ShadingUniversal` object.|
||[foregroundPatternColorIndex](/.shadinguniversal#word-javascript/api/word/-shadinguniversal-foregroundpatterncolorindex-member)|Specifies the color that's applied to the foreground of the `ShadingUniversal` object.|
||[texture](/.shadinguniversal#word-javascript/api/word/-shadinguniversal-texture-member)|Specifies the shading texture of the object.|
|[ShadowFormat](/.shadowformat)|[blur](/.shadowformat#word-javascript/api/word/-shadowformat-blur-member)|Specifies the blur level for a shadow format as a value between 0.0 and 100.0.|
||[foregroundColor](/.shadowformat#word-javascript/api/word/-shadowformat-foregroundcolor-member)|Returns a `ColorFormat` object that represents the foreground color for the fill, line, or shadow.|
||[incrementOffsetX(increment: number)](/.shadowformat#word-javascript/api/word/-shadowformat-incrementoffsetx-member(1))|Changes the horizontal offset of the shadow by the number of points.|
||[incrementOffsetY(increment: number)](/.shadowformat#word-javascript/api/word/-shadowformat-incrementoffsety-member(1))|Changes the vertical offset of the shadow by the specified number of points.|
||[isVisible](/.shadowformat#word-javascript/api/word/-shadowformat-isvisible-member)|Specifies whether the object or the formatting applied to it is visible.|
||[obscured](/.shadowformat#word-javascript/api/word/-shadowformat-obscured-member)|Specifies `true` if the shadow of the shape appears filled in and is obscured by the shape, even if the shape has no fill,|
||[offsetX](/.shadowformat#word-javascript/api/word/-shadowformat-offsetx-member)|Specifies the horizontal offset (in points) of the shadow from the shape.|
||[offsetY](/.shadowformat#word-javascript/api/word/-shadowformat-offsety-member)|Specifies the vertical offset (in points) of the shadow from the shape.|
||[rotateWithShape](/.shadowformat#word-javascript/api/word/-shadowformat-rotatewithshape-member)|Specifies whether to rotate the shadow when rotating the shape.|
||[size](/.shadowformat#word-javascript/api/word/-shadowformat-size-member)|Specifies the width of the shadow.|
||[style](/.shadowformat#word-javascript/api/word/-shadowformat-style-member)|Specifies the type of shadow formatting to apply to a shape.|
||[transparency](/.shadowformat#word-javascript/api/word/-shadowformat-transparency-member)|Specifies the degree of transparency of the shadow as a value between 0.0 (opaque) and 1.0 (clear).|
||[type](/.shadowformat#word-javascript/api/word/-shadowformat-type-member)|Specifies the shape shadow type.|
|[Source](/.source)|[delete()](/.source#word-javascript/api/word/-source-delete-member(1))|Deletes the `Source` object.|
||[getFieldByName(name: string)](/.source#word-javascript/api/word/-source-getfieldbyname-member(1))|Returns the value of a field in the bibliography `Source` object.|
||[isCited](/.source#word-javascript/api/word/-source-iscited-member)|Gets if the `Source` object has been cited in the document.|
||[tag](/.source#word-javascript/api/word/-source-tag-member)|Gets the tag of the source.|
||[xml](/.source#word-javascript/api/word/-source-xml-member)|Gets the XML representation of the source.|
|[SourceCollection](/.sourcecollection)|[add(xml: string)](/.sourcecollection#word-javascript/api/word/-sourcecollection-add-member(1))|Adds a new `Source` object to the collection.|
||[getItem(index: number)](/.sourcecollection#word-javascript/api/word/-sourcecollection-getitem-member(1))|Gets a `Source` by its index in the collection.|
||[items](/.sourcecollection#word-javascript/api/word/-sourcecollection-items-member)|Gets the loaded child items in this collection.|
|[Style](/.style)|[automaticallyUpdate](/.style#word-javascript/api/word/-style-automaticallyupdate-member)|Specifies whether the style is automatically redefined based on the selection.|
||[description](/.style#word-javascript/api/word/-style-description-member)|Gets the description of the specified style.|
||[frame](/.style#word-javascript/api/word/-style-frame-member)|Returns a `Frame` object that represents the frame formatting for the style.|
||[hasProofing](/.style#word-javascript/api/word/-style-hasproofing-member)|Specifies whether the spelling and grammar checker ignores text formatted with this style.|
||[languageId](/.style#word-javascript/api/word/-style-languageid-member)|Specifies a `LanguageId` value that represents the language for the style.|
||[languageIdFarEast](/.style#word-javascript/api/word/-style-languageidfareast-member)|Specifies an East Asian language for the style.|
||[linkStyle](/.style#word-javascript/api/word/-style-linkstyle-member)|Specifies a link between a paragraph and a character style.|
||[linkToListTemplate(listTemplate: Word.ListTemplate)](/.style#word-javascript/api/word/-style-linktolisttemplate-member(1))|Links this style to a list template so that the style's formatting can be applied to lists.|
||[listLevelNumber](/.style#word-javascript/api/word/-style-listlevelnumber-member)|Returns the list level for the style.|
||[locked](/.style#word-javascript/api/word/-style-locked-member)|Specifies whether the style cannot be changed or edited.|
||[noSpaceBetweenParagraphsOfSameStyle](/.style#word-javascript/api/word/-style-nospacebetweenparagraphsofsamestyle-member)|Specifies whether to remove spacing between paragraphs that are formatted using the same style.|
|[TabStop](/.tabstop)|[alignment](/.tabstop#word-javascript/api/word/-tabstop-alignment-member)|Gets a `TabAlignment` value that represents the alignment for the tab stop.|
||[clear()](/.tabstop#word-javascript/api/word/-tabstop-clear-member(1))|Removes this custom tab stop.|
||[customTab](/.tabstop#word-javascript/api/word/-tabstop-customtab-member)|Gets whether this tab stop is a custom tab stop.|
||[leader](/.tabstop#word-javascript/api/word/-tabstop-leader-member)|Gets a `TabLeader` value that represents the leader for this `TabStop` object.|
||[next](/.tabstop#word-javascript/api/word/-tabstop-next-member)|Gets the next tab stop in the collection.|
||[position](/.tabstop#word-javascript/api/word/-tabstop-position-member)|Gets the position of the tab stop relative to the left margin.|
||[previous](/.tabstop#word-javascript/api/word/-tabstop-previous-member)|Gets the previous tab stop in the collection.|
|[TabStopAddOptions](/.tabstopaddoptions)|[alignment](/.tabstopaddoptions#word-javascript/api/word/-tabstopaddoptions-alignment-member)|If provided, specifies the alignment of the tab stop.|
||[leader](/.tabstopaddoptions#word-javascript/api/word/-tabstopaddoptions-leader-member)|If provided, specifies the leader character for the tab stop.|
|[TabStopCollection](/.tabstopcollection)|[add(position: number, options?: Word.TabStopAddOptions)](/.tabstopcollection#word-javascript/api/word/-tabstopcollection-add-member(1))|Returns a `TabStop` object that represents a custom tab stop added to the paragraph.|
||[after(Position: number)](/.tabstopcollection#word-javascript/api/word/-tabstopcollection-after-member(1))|Returns the next `TabStop` object to the right of the specified position.|
||[before(Position: number)](/.tabstopcollection#word-javascript/api/word/-tabstopcollection-before-member(1))|Returns the next `TabStop` object to the left of the specified position.|
||[clearAll()](/.tabstopcollection#word-javascript/api/word/-tabstopcollection-clearall-member(1))|Clears all the custom tab stops from the paragraph.|
||[getItem(index: number)](/.tabstopcollection#word-javascript/api/word/-tabstopcollection-getitem-member(1))|Gets a `TabStop` object by its index in the collection.|
||[items](/.tabstopcollection#word-javascript/api/word/-tabstopcollection-items-member)|Gets the loaded child items in this collection.|
|[TableColumn](/.tablecolumn)|[autoFit()](/.tablecolumn#word-javascript/api/word/-tablecolumn-autofit-member(1))|Changes the width of the table column to accommodate the width of the text without changing the way text wraps in the cells.|
||[borders](/.tablecolumn#word-javascript/api/word/-tablecolumn-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the table column.|
||[columnIndex](/.tablecolumn#word-javascript/api/word/-tablecolumn-columnindex-member)|Returns the position of this column in a collection.|
||[delete()](/.tablecolumn#word-javascript/api/word/-tablecolumn-delete-member(1))|Deletes the column.|
||[isFirst](/.tablecolumn#word-javascript/api/word/-tablecolumn-isfirst-member)|Returns `true` if the column or row is the first one in the table; `false` otherwise.|
||[isLast](/.tablecolumn#word-javascript/api/word/-tablecolumn-islast-member)|Returns `true` if the column or row is the last one in the table; `false` otherwise.|
||[nestingLevel](/.tablecolumn#word-javascript/api/word/-tablecolumn-nestinglevel-member)|Returns the nesting level of the column.|
||[preferredWidth](/.tablecolumn#word-javascript/api/word/-tablecolumn-preferredwidth-member)|Specifies the preferred width (in points or as a percentage of the window width) for the column.|
||[preferredWidthType](/.tablecolumn#word-javascript/api/word/-tablecolumn-preferredwidthtype-member)|Specifies the preferred unit of measurement to use for the width of the table column.|
||[select()](/.tablecolumn#word-javascript/api/word/-tablecolumn-select-member(1))|Selects the table column.|
||[setWidth(columnWidth: number, rulerStyle: Word.RulerStyle)](/.tablecolumn#word-javascript/api/word/-tablecolumn-setwidth-member(1))|Sets the width of the column in a table.|
||[shading](/.tablecolumn#word-javascript/api/word/-tablecolumn-shading-member)|Returns a `ShadingUniversal` object that refers to the shading formatting for the column.|
||[sort()](/.tablecolumn#word-javascript/api/word/-tablecolumn-sort-member(1))|Sorts the table column.|
||[width](/.tablecolumn#word-javascript/api/word/-tablecolumn-width-member)|Specifies the width of the column, in points.|
|[TableColumnCollection](/.tablecolumncollection)|[add(beforeColumn?: Word.TableColumn)](/.tablecolumncollection#word-javascript/api/word/-tablecolumncollection-add-member(1))|Returns a `TableColumn` object that represents a column added to a table.|
||[autoFit()](/.tablecolumncollection#word-javascript/api/word/-tablecolumncollection-autofit-member(1))|Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.|
||[delete()](/.tablecolumncollection#word-javascript/api/word/-tablecolumncollection-delete-member(1))|Deletes the specified columns.|
||[distributeWidth()](/.tablecolumncollection#word-javascript/api/word/-tablecolumncollection-distributewidth-member(1))|Adjusts the width of the specified columns so that they are equal.|
||[items](/.tablecolumncollection#word-javascript/api/word/-tablecolumncollection-items-member)|Gets the loaded child items in this collection.|
||[select()](/.tablecolumncollection#word-javascript/api/word/-tablecolumncollection-select-member(1))|Selects the specified table columns.|
||[setWidth(columnWidth: number, rulerStyle: Word.RulerStyle)](/.tablecolumncollection#word-javascript/api/word/-tablecolumncollection-setwidth-member(1))|Sets the width of columns in a table.|
|[Template](/.template)|[buildingBlockEntries](/.template#word-javascript/api/word/-template-buildingblockentries-member)|Returns a `BuildingBlockEntryCollection` object that represents the collection of building block entries in the template.|
||[buildingBlockTypes](/.template#word-javascript/api/word/-template-buildingblocktypes-member)|Returns a `BuildingBlockTypeItemCollection` object that represents the collection of building block types that are contained in the template.|
||[farEastLineBreakLanguage](/.template#word-javascript/api/word/-template-fareastlinebreaklanguage-member)|Specifies the East Asian language to use when breaking lines of text in the document or template.|
||[farEastLineBreakLevel](/.template#word-javascript/api/word/-template-fareastlinebreaklevel-member)|Specifies the line break control level for the document.|
||[fullName](/.template#word-javascript/api/word/-template-fullname-member)|Returns the name of the template, including the drive or Web path.|
||[hasNoProofing](/.template#word-javascript/api/word/-template-hasnoproofing-member)|Specifies whether the spelling and grammar checker ignores documents based on this template.|
||[justificationMode](/.template#word-javascript/api/word/-template-justificationmode-member)|Specifies the character spacing adjustment for the template.|
||[kerningByAlgorithm](/.template#word-javascript/api/word/-template-kerningbyalgorithm-member)|Specifies if Microsoft Word kerns half-width Latin characters and punctuation marks in the document.|
||[languageId](/.template#word-javascript/api/word/-template-languageid-member)|Specifies a `LanguageId` value that represents the language in the template.|
||[languageIdFarEast](/.template#word-javascript/api/word/-template-languageidfareast-member)|Specifies an East Asian language for the language in the template.|
||[name](/.template#word-javascript/api/word/-template-name-member)|Returns only the name of the document template (excluding any path or other location information).|
||[noLineBreakAfter](/.template#word-javascript/api/word/-template-nolinebreakafter-member)|Specifies the kinsoku characters after which Microsoft Word will not break a line.|
||[noLineBreakBefore](/.template#word-javascript/api/word/-template-nolinebreakbefore-member)|Specifies the kinsoku characters before which Microsoft Word will not break a line.|
||[path](/.template#word-javascript/api/word/-template-path-member)|Returns the path to the document template.|
||[save()](/.template#word-javascript/api/word/-template-save-member(1))|Saves the template.|
||[saved](/.template#word-javascript/api/word/-template-saved-member)|Specifies `true` if the template has not changed since it was last saved, `false` if Microsoft Word displays a prompt to save changes when the document is closed.|
||[type](/.template#word-javascript/api/word/-template-type-member)|Returns the template type.|
|[TemplateCollection](/.templatecollection)|[getCount()](/.templatecollection#word-javascript/api/word/-templatecollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/.templatecollection#word-javascript/api/word/-templatecollection-getitemat-member(1))|Gets a `Template` object by its index in the collection.|
||[importBuildingBlocks()](/.templatecollection#word-javascript/api/word/-templatecollection-importbuildingblocks-member(1))|Imports the building blocks for all templates into Microsoft Word.|
||[items](/.templatecollection#word-javascript/api/word/-templatecollection-items-member)|Gets the loaded child items in this collection.|
|[TextColumn](/.textcolumn)|[spaceAfter](/.textcolumn#word-javascript/api/word/-textcolumn-spaceafter-member)|Specifies the amount of spacing (in points) after the specified paragraph or text column.|
||[width](/.textcolumn#word-javascript/api/word/-textcolumn-width-member)|Specifies the width, in points, of the specified text columns.|
|[TextColumnAddOptions](/.textcolumnaddoptions)|[isEvenlySpaced](/.textcolumnaddoptions#word-javascript/api/word/-textcolumnaddoptions-isevenlyspaced-member)|If provided, specifies whether to evenly space all the text columns in the document.|
||[spacing](/.textcolumnaddoptions#word-javascript/api/word/-textcolumnaddoptions-spacing-member)|If provided, specifies the spacing between the text columns in the document, in points.|
||[width](/.textcolumnaddoptions#word-javascript/api/word/-textcolumnaddoptions-width-member)|If provided, specifies the width of the new text column in the document, in points.|
|[TextColumnCollection](/.textcolumncollection)|[add(options?: Word.TextColumnAddOptions)](/.textcolumncollection#word-javascript/api/word/-textcolumncollection-add-member(1))|Returns a `TextColumn` object that represents a new text column added to a section or document.|
||[getFlowDirection()](/.textcolumncollection#word-javascript/api/word/-textcolumncollection-getflowdirection-member(1))|Gets the direction in which text flows from one text column to the next.|
||[getHasLineBetween()](/.textcolumncollection#word-javascript/api/word/-textcolumncollection-gethaslinebetween-member(1))|Gets whether vertical lines appear between all the columns in the `TextColumnCollection` object.|
||[getIsEvenlySpaced()](/.textcolumncollection#word-javascript/api/word/-textcolumncollection-getisevenlyspaced-member(1))|Gets whether text columns are evenly spaced.|
||[getItem(index: number)](/.textcolumncollection#word-javascript/api/word/-textcolumncollection-getitem-member(1))|Gets a `TextColumn` by its index in the collection.|
||[items](/.textcolumncollection#word-javascript/api/word/-textcolumncollection-items-member)|Gets the loaded child items in this collection.|
||[setCount(numColumns: number)](/.textcolumncollection#word-javascript/api/word/-textcolumncollection-setcount-member(1))|Arranges text into the specified number of text columns.|
||[setFlowDirection(value: Word.FlowDirection)](/.textcolumncollection#word-javascript/api/word/-textcolumncollection-setflowdirection-member(1))|Sets the direction in which text flows from one text column to the next.|
||[setHasLineBetween(value: boolean)](/.textcolumncollection#word-javascript/api/word/-textcolumncollection-sethaslinebetween-member(1))|Sets whether vertical lines appear between all the columns in the `TextColumnCollection` object.|
||[setIsEvenlySpaced(value: boolean)](/.textcolumncollection#word-javascript/api/word/-textcolumncollection-setisevenlyspaced-member(1))|Sets whether text columns are evenly spaced.|
|[ThreeDimensionalFormat](/.threedimensionalformat)|[bevelBottomDepth](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-bevelbottomdepth-member)|Specifies the depth of the bottom bevel.|
||[bevelBottomInset](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-bevelbottominset-member)|Specifies the inset size for the bottom bevel.|
||[bevelBottomType](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-bevelbottomtype-member)|Specifies a `BevelType` value that represents the bevel type for the bottom bevel.|
||[bevelTopDepth](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-beveltopdepth-member)|Specifies the depth of the top bevel.|
||[bevelTopInset](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-beveltopinset-member)|Specifies the inset size for the top bevel.|
||[bevelTopType](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-beveltoptype-member)|Specifies a `BevelType` value that represents the bevel type for the top bevel.|
||[contourColor](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-contourcolor-member)|Returns a `ColorFormat` object that represents color of the contour of a shape.|
||[contourWidth](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-contourwidth-member)|Specifies the width of the contour of a shape.|
||[depth](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-depth-member)|Specifies the depth of the shape's extrusion.|
||[extrusionColor](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-extrusioncolor-member)|Returns a `ColorFormat` object that represents the color of the shape's extrusion.|
||[extrusionColorType](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-extrusioncolortype-member)|Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion)|
||[fieldOfView](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-fieldofview-member)|Specifies the amount of perspective for a shape.|
||[incrementRotationHorizontal(increment: number)](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-incrementrotationhorizontal-member(1))|Horizontally rotates a shape on the x-axis.|
||[incrementRotationVertical(increment: number)](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-incrementrotationvertical-member(1))|Vertically rotates a shape on the y-axis.|
||[incrementRotationX(increment: number)](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-incrementrotationx-member(1))|Changes the rotation around the x-axis.|
||[incrementRotationY(increment: number)](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-incrementrotationy-member(1))|Changes the rotation around the y-axis.|
||[incrementRotationZ(increment: number)](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-incrementrotationz-member(1))|Rotates a shape on the z-axis.|
||[isPerspective](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-isperspective-member)|Specifies `true` if the extrusion appears in perspective — that is, if the walls of the extrusion narrow toward a vanishing point,|
||[isVisible](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-isvisible-member)|Specifies if the specified object, or the formatting applied to it, is visible.|
||[lightAngle](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-lightangle-member)|Specifies the angle of the lighting.|
||[presetCamera](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-presetcamera-member)|Returns a `PresetCamera` value that represents the camera presets.|
||[presetExtrusionDirection](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-presetextrusiondirection-member)|Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion).|
||[presetLighting](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-presetlighting-member)|Specifies a `LightRigType` value that represents the lighting preset.|
||[presetLightingDirection](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-presetlightingdirection-member)|Specifies the position of the light source relative to the extrusion.|
||[presetLightingSoftness](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-presetlightingsoftness-member)|Specifies the intensity of the extrusion lighting.|
||[presetMaterial](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-presetmaterial-member)|Specifies the extrusion surface material.|
||[presetThreeDimensionalFormat](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-presetthreedimensionalformat-member)|Returns the preset extrusion format.|
||[projectText](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-projecttext-member)|Specifies whether text on a shape rotates with shape.|
||[resetRotation()](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-resetrotation-member(1))|Resets the extrusion rotation around the x-axis, y-axis, and z-axis to 0.|
||[rotationX](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-rotationx-member)|Specifies the rotation of the extruded shape around the x-axis in degrees.|
||[rotationY](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-rotationy-member)|Specifies the rotation of the extruded shape around the y-axis in degrees.|
||[rotationZ](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-rotationz-member)|Specifies the z-axis rotation of the camera.|
||[setExtrusionDirection(presetExtrusionDirection: Word.PresetExtrusionDirection)](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-setextrusiondirection-member(1))|Sets the direction of the extrusion's sweep path.|
||[setPresetCamera(presetCamera: Word.PresetCamera)](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-setpresetcamera-member(1))|Sets the camera preset for the shape.|
||[setThreeDimensionalFormat(presetThreeDimensionalFormat: Word.PresetThreeDimensionalFormat)](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-setthreedimensionalformat-member(1))|Sets the preset extrusion format.|
||[z](/.threedimensionalformat#word-javascript/api/word/-threedimensionalformat-z-member)|Specifies the position on the z-axis for the shape.|
|[View](/.view)|[areAllNonprintingCharactersDisplayed](/.view#word-javascript/api/word/-view-areallnonprintingcharactersdisplayed-member)|Specifies whether all nonprinting characters are displayed.|
||[areBackgroundsDisplayed](/.view#word-javascript/api/word/-view-arebackgroundsdisplayed-member)|Gets whether background colors and images are shown when the document is displayed in print layout view.|
||[areBookmarksIndicated](/.view#word-javascript/api/word/-view-arebookmarksindicated-member)|Gets whether square brackets are displayed at the beginning and end of each bookmark.|
||[areCommentsDisplayed](/.view#word-javascript/api/word/-view-arecommentsdisplayed-member)|Specifies whether Microsoft Word displays the comments in the document.|
||[areConnectingLinesToRevisionsBalloonDisplayed](/.view#word-javascript/api/word/-view-areconnectinglinestorevisionsballoondisplayed-member)|Specifies whether Microsoft Word displays connecting lines from the text to the revision and comment balloons.|
||[areCropMarksDisplayed](/.view#word-javascript/api/word/-view-arecropmarksdisplayed-member)|Gets whether crop marks are shown in the corners of pages to indicate where margins are located.|
||[areDrawingsDisplayed](/.view#word-javascript/api/word/-view-aredrawingsdisplayed-member)|Gets whether objects created with the drawing tools are displayed in print layout view.|
||[areEditableRangesShaded](/.view#word-javascript/api/word/-view-areeditablerangesshaded-member)|Specifies whether shading is applied to the ranges in the document that users have permission to modify.|
||[areFieldCodesDisplayed](/.view#word-javascript/api/word/-view-arefieldcodesdisplayed-member)|Specifies whether field codes are displayed.|
||[areFormatChangesDisplayed](/.view#word-javascript/api/word/-view-areformatchangesdisplayed-member)|Specifies whether Microsoft Word displays formatting changes made to the document with Track Changes enabled.|
||[areInkAnnotationsDisplayed](/.view#word-javascript/api/word/-view-areinkannotationsdisplayed-member)|Specifies whether handwritten ink annotations are shown or hidden.|
||[areInsertionsAndDeletionsDisplayed](/.view#word-javascript/api/word/-view-areinsertionsanddeletionsdisplayed-member)|Specifies whether Microsoft Word displays insertions and deletions made to the document with Track Changes enabled.|
||[areLinesWrappedToWindow](/.view#word-javascript/api/word/-view-arelineswrappedtowindow-member)|Gets whether lines wrap at the right edge of the document window rather than at the right margin or the right column boundary.|
||[areObjectAnchorsDisplayed](/.view#word-javascript/api/word/-view-areobjectanchorsdisplayed-member)|Gets whether object anchors are displayed next to items that can be positioned in print layout view.|
||[areOptionalBreaksDisplayed](/.view#word-javascript/api/word/-view-areoptionalbreaksdisplayed-member)|Gets whether Microsoft Word displays optional line breaks.|
||[areOptionalHyphensDisplayed](/.view#word-javascript/api/word/-view-areoptionalhyphensdisplayed-member)|Gets whether optional hyphens are displayed.|
||[areOtherAuthorsVisible](/.view#word-javascript/api/word/-view-areotherauthorsvisible-member)|Gets whether other authors' presence should be visible in the document.|
||[arePageBoundariesDisplayed](/.view#word-javascript/api/word/-view-arepageboundariesdisplayed-member)|Gets whether the top and bottom margins and the gray area between pages in the document are displayed.|
||[areParagraphsMarksDisplayed](/.view#word-javascript/api/word/-view-areparagraphsmarksdisplayed-member)|Gets whether paragraph marks are displayed.|
||[arePicturePlaceholdersDisplayed](/.view#word-javascript/api/word/-view-arepictureplaceholdersdisplayed-member)|Gets whether blank boxes are displayed as placeholders for pictures.|
||[areRevisionsAndCommentsDisplayed](/.view#word-javascript/api/word/-view-arerevisionsandcommentsdisplayed-member)|Specifies whether Microsoft Word displays revisions and comments made to the document with Track Changes enabled.|
||[areSpacesIndicated](/.view#word-javascript/api/word/-view-arespacesindicated-member)|Gets whether space characters are displayed.|
||[areTableGridlinesDisplayed](/.view#word-javascript/api/word/-view-aretablegridlinesdisplayed-member)|Specifies whether table gridlines are displayed.|
||[areTabsDisplayed](/.view#word-javascript/api/word/-view-aretabsdisplayed-member)|Gets whether tab characters are displayed.|
||[areTextBoundariesDisplayed](/.view#word-javascript/api/word/-view-aretextboundariesdisplayed-member)|Gets whether dotted lines are displayed around page margins, text columns, objects, and frames in print layout view.|
||[collapseAllHeadings()](/.view#word-javascript/api/word/-view-collapseallheadings-member(1))|Collapses all the headings in the document.|
||[collapseOutline(range: Word.Range)](/.view#word-javascript/api/word/-view-collapseoutline-member(1))|Collapses the text under the selection or the specified range by one heading level.|
||[columnWidth](/.view#word-javascript/api/word/-view-columnwidth-member)|Specifies the column width in Reading mode.|
||[expandAllHeadings()](/.view#word-javascript/api/word/-view-expandallheadings-member(1))|Expands all the headings in the document.|
||[expandOutline(range: Word.Range)](/.view#word-javascript/api/word/-view-expandoutline-member(1))|Expands the text under the selection by one heading level.|
||[fieldShading](/.view#word-javascript/api/word/-view-fieldshading-member)|Gets on-screen shading for fields.|
||[isDraft](/.view#word-javascript/api/word/-view-isdraft-member)|Specifies whether all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.|
||[isFirstLineOnlyDisplayed](/.view#word-javascript/api/word/-view-isfirstlineonlydisplayed-member)|Specifies whether only the first line of body text is shown in outline view.|
||[isFormatDisplayed](/.view#word-javascript/api/word/-view-isformatdisplayed-member)|Specifies whether character formatting is visible in outline view.|
||[isFullScreen](/.view#word-javascript/api/word/-view-isfullscreen-member)|Specifies whether the window is in full-screen view.|
||[isHiddenTextDisplayed](/.view#word-javascript/api/word/-view-ishiddentextdisplayed-member)|Gets whether text formatted as hidden text is displayed.|
||[isHighlightingDisplayed](/.view#word-javascript/api/word/-view-ishighlightingdisplayed-member)|Gets whether highlight formatting is displayed and printed with the document.|
||[isInConflictMode](/.view#word-javascript/api/word/-view-isinconflictmode-member)|Specifies whether the document is in conflict mode view.|
||[isInPanning](/.view#word-javascript/api/word/-view-isinpanning-member)|Specifies whether Microsoft Word is in Panning mode.|
||[isInReadingLayout](/.view#word-javascript/api/word/-view-isinreadinglayout-member)|Specifies whether the document is being viewed in reading layout view.|
||[isMailMergeDataView](/.view#word-javascript/api/word/-view-ismailmergedataview-member)|Specifies whether mail merge data is displayed instead of mail merge fields.|
||[isMainTextLayerVisible](/.view#word-javascript/api/word/-view-ismaintextlayervisible-member)|Specifies whether the text in the document is visible when the header and footer areas are displayed.|
||[isPointerShownAsMagnifier](/.view#word-javascript/api/word/-view-ispointershownasmagnifier-member)|Specifies whether the pointer is displayed as a magnifying glass in print preview.|
||[isReadingLayoutActualView](/.view#word-javascript/api/word/-view-isreadinglayoutactualview-member)|Specifies whether pages displayed in reading layout view are displayed using the same layout as printed pages.|
||[isXmlMarkupVisible](/.view#word-javascript/api/word/-view-isxmlmarkupvisible-member)|Specifies whether XML tags are visible in the document.|
||[markupMode](/.view#word-javascript/api/word/-view-markupmode-member)|Specifies the display mode for tracked changes.|
||[nextHeaderFooter()](/.view#word-javascript/api/word/-view-nextheaderfooter-member(1))|Moves to the next header or footer, depending on whether a header or footer is displayed in the view.|
||[pageColor](/.view#word-javascript/api/word/-view-pagecolor-member)|Specifies the page color in Reading mode.|
||[pageMovementType](/.view#word-javascript/api/word/-view-pagemovementtype-member)|Specifies the page movement type.|
||[previousHeaderFooter()](/.view#word-javascript/api/word/-view-previousheaderfooter-member(1))|Moves to the previous header or footer, depending on whether a header or footer is displayed in the view.|
||[readingLayoutTruncateMargins](/.view#word-javascript/api/word/-view-readinglayouttruncatemargins-member)|Specifies whether margins are visible or hidden when the document is viewed in Full Screen Reading view.|
||[revisionsBalloonSide](/.view#word-javascript/api/word/-view-revisionsballoonside-member)|Gets whether Word displays revision balloons in the left or right margin in the document.|
||[revisionsBalloonWidth](/.view#word-javascript/api/word/-view-revisionsballoonwidth-member)|Specifies the width of the revision balloons.|
||[revisionsBalloonWidthType](/.view#word-javascript/api/word/-view-revisionsballoonwidthtype-member)|Specifies how Microsoft Word measures the width of revision balloons.|
||[revisionsFilter](/.view#word-javascript/api/word/-view-revisionsfilter-member)|Gets the instance of a `RevisionsFilter` object.|
||[seekView](/.view#word-javascript/api/word/-view-seekview-member)|Specifies the document element displayed in print layout view.|
||[showAllHeadings()](/.view#word-javascript/api/word/-view-showallheadings-member(1))|Switches between showing all text (headings and body text) and showing only headings.|
||[showHeading(level: number)](/.view#word-javascript/api/word/-view-showheading-member(1))|Shows all headings up to the specified heading level and hides subordinate headings and body text.|
||[splitSpecial](/.view#word-javascript/api/word/-view-splitspecial-member)|Specifies the active window pane.|
||[type](/.view#word-javascript/api/word/-view-type-member)|Specifies the view type.|
|[Window](/.window)|[activate()](/.window#word-javascript/api/word/-window-activate-member(1))|Activates the window.|
||[areRulersDisplayed](/.window#word-javascript/api/word/-window-arerulersdisplayed-member)|Specifies whether rulers are displayed for the window or pane.|
||[areScreenTipsDisplayed](/.window#word-javascript/api/word/-window-arescreentipsdisplayed-member)|Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.|
||[areThumbnailsDisplayed](/.window#word-javascript/api/word/-window-arethumbnailsdisplayed-member)|Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.|
||[caption](/.window#word-javascript/api/word/-window-caption-member)|Specifies the caption text for the window that is displayed in the title bar of the document or application window.|
||[close(options?: Word.WindowCloseOptions)](/.window#word-javascript/api/word/-window-close-member(1))|Closes the window.|
||[height](/.window#word-javascript/api/word/-window-height-member)|Specifies the height of the window (in points).|
||[horizontalPercentScrolled](/.window#word-javascript/api/word/-window-horizontalpercentscrolled-member)|Specifies the horizontal scroll position as a percentage of the document width.|
||[imemode](/.window#word-javascript/api/word/-window-imemode-member)|Specifies the default start-up mode for the Japanese Input Method Editor (IME).|
||[index](/.window#word-javascript/api/word/-window-index-member)|Gets the position of an item in a collection.|
||[isActive](/.window#word-javascript/api/word/-window-isactive-member)|Specifies whether the window is active.|
||[isDocumentMapVisible](/.window#word-javascript/api/word/-window-isdocumentmapvisible-member)|Specifies whether the document map is visible.|
||[isEnvelopeVisible](/.window#word-javascript/api/word/-window-isenvelopevisible-member)|Specifies whether the email message header is visible in the document window.|
||[isHorizontalScrollBarDisplayed](/.window#word-javascript/api/word/-window-ishorizontalscrollbardisplayed-member)|Specifies whether a horizontal scroll bar is displayed for the window.|
||[isLeftScrollBarDisplayed](/.window#word-javascript/api/word/-window-isleftscrollbardisplayed-member)|Specifies whether the vertical scroll bar appears on the left side of the document window.|
||[isRightRulerDisplayed](/.window#word-javascript/api/word/-window-isrightrulerdisplayed-member)|Specifies whether the vertical ruler appears on the right side of the document window in print layout view.|
||[isSplit](/.window#word-javascript/api/word/-window-issplit-member)|Specifies whether the window is split into multiple panes.|
||[isVerticalRulerDisplayed](/.window#word-javascript/api/word/-window-isverticalrulerdisplayed-member)|Specifies whether a vertical ruler is displayed for the window or pane.|
||[isVerticalScrollBarDisplayed](/.window#word-javascript/api/word/-window-isverticalscrollbardisplayed-member)|Specifies whether a vertical scroll bar is displayed for the window.|
||[isVisible](/.window#word-javascript/api/word/-window-isvisible-member)|Specifies whether the window is visible.|
||[largeScroll(options?: Word.WindowScrollOptions)](/.window#word-javascript/api/word/-window-largescroll-member(1))|Scrolls the window by the specified number of screens.|
||[left](/.window#word-javascript/api/word/-window-left-member)|Specifies the horizontal position of the window, measured in points.|
||[next](/.window#word-javascript/api/word/-window-next-member)|Gets the next document window in the collection of open document windows.|
||[pageScroll(options?: Word.WindowPageScrollOptions)](/.window#word-javascript/api/word/-window-pagescroll-member(1))|Scrolls through the window page by page.|
||[previous](/.window#word-javascript/api/word/-window-previous-member)|Gets the previous document window in the collection open document windows.|
||[setFocus()](/.window#word-javascript/api/word/-window-setfocus-member(1))|Sets the focus of the document window to the body of an email message.|
||[showSourceDocuments](/.window#word-javascript/api/word/-window-showsourcedocuments-member)|Specifies how Microsoft Word displays source documents after a compare and merge process.|
||[smallScroll(options?: Word.WindowScrollOptions)](/.window#word-javascript/api/word/-window-smallscroll-member(1))|Scrolls the window by the specified number of lines.|
||[splitVertical](/.window#word-javascript/api/word/-window-splitvertical-member)|Specifies the vertical split percentage for the window.|
||[styleAreaWidth](/.window#word-javascript/api/word/-window-styleareawidth-member)|Specifies the width of the style area in points.|
||[toggleRibbon()](/.window#word-javascript/api/word/-window-toggleribbon-member(1))|Shows or hides the ribbon.|
||[top](/.window#word-javascript/api/word/-window-top-member)|Specifies the vertical position of the document window, in points.|
||[type](/.window#word-javascript/api/word/-window-type-member)|Gets the window type.|
||[usableHeight](/.window#word-javascript/api/word/-window-usableheight-member)|Gets the height (in points) of the active working area in the document window.|
||[usableWidth](/.window#word-javascript/api/word/-window-usablewidth-member)|Gets the width (in points) of the active working area in the document window.|
||[verticalPercentScrolled](/.window#word-javascript/api/word/-window-verticalpercentscrolled-member)|Specifies the vertical scroll position as a percentage of the document length.|
||[view](/.window#word-javascript/api/word/-window-view-member)|Gets the `View` object that represents the view for the window.|
||[width](/.window#word-javascript/api/word/-window-width-member)|Specifies the width of the document window, in points.|
||[windowNumber](/.window#word-javascript/api/word/-window-windownumber-member)|Gets an integer that represents the position of the window.|
||[windowState](/.window#word-javascript/api/word/-window-windowstate-member)|Specifies the state of the document window or task window.|
|[WindowCloseOptions](/.windowcloseoptions)|[routeDocument](/.windowcloseoptions#word-javascript/api/word/-windowcloseoptions-routedocument-member)|If provided, specifies whether to route the document to the next recipient.|
||[saveChanges](/.windowcloseoptions#word-javascript/api/word/-windowcloseoptions-savechanges-member)|If provided, specifies the save action for the document.|
|[WindowCollection](/.windowcollection)|||
|[WindowPageScrollOptions](/.windowpagescrolloptions)|[down](/.windowpagescrolloptions#word-javascript/api/word/-windowpagescrolloptions-down-member)|If provided, specifies the number of pages to scroll the window down.|
||[up](/.windowpagescrolloptions#word-javascript/api/word/-windowpagescrolloptions-up-member)|If provided, specifies the number of pages to scroll the window up.|
|[WindowScrollOptions](/.windowscrolloptions)|[down](/.windowscrolloptions#word-javascript/api/word/-windowscrolloptions-down-member)|If provided, specifies the number of units to scroll the window down.|
||[left](/.windowscrolloptions#word-javascript/api/word/-windowscrolloptions-left-member)|If provided, specifies the number of screens to scroll the window to the left.|
||[right](/.windowscrolloptions#word-javascript/api/word/-windowscrolloptions-right-member)|If provided, specifies the number of screens to scroll the window to the right.|
||[up](/.windowscrolloptions#word-javascript/api/word/-windowscrolloptions-up-member)|If provided, specifies the number of units to scroll the window up.|
|[XmlMapping](/.xmlmapping)|[customXmlNode](/.xmlmapping#word-javascript/api/word/-xmlmapping-customxmlnode-member)|Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.|
||[customXmlPart](/.xmlmapping#word-javascript/api/word/-xmlmapping-customxmlpart-member)|Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.|
||[delete()](/.xmlmapping#word-javascript/api/word/-xmlmapping-delete-member(1))|Deletes the XML mapping from the parent content control.|
||[isMapped](/.xmlmapping#word-javascript/api/word/-xmlmapping-ismapped-member)|Returns whether the content control in the document is mapped to an XML node in the document's XML data store.|
||[prefixMappings](/.xmlmapping#word-javascript/api/word/-xmlmapping-prefixmappings-member)|Returns the prefix mappings used to evaluate the XPath for the current XML mapping.|
||[setMapping(xPath: string, options?: Word.XmlSetMappingOptions)](/.xmlmapping#word-javascript/api/word/-xmlmapping-setmapping-member(1))|Allows creating or changing the XML mapping on the content control.|
||[setMappingByNode(node: Word.CustomXmlNode)](/.xmlmapping#word-javascript/api/word/-xmlmapping-setmappingbynode-member(1))|Allows creating or changing the XML data mapping on the content control.|
||[xpath](/.xmlmapping#word-javascript/api/word/-xmlmapping-xpath-member)|Returns the XPath for the XML mapping, which evaluates to the currently mapped XML node.|
|[XmlSetMappingOptions](/.xmlsetmappingoptions)|[prefixMapping](/.xmlsetmappingoptions#word-javascript/api/word/-xmlsetmappingoptions-prefixmapping-member)|If provided, specifies the prefix mappings to use when querying the expression provided in the `xPath` parameter of the `XmlMapping.setMapping` calling method.|
||[source](/.xmlsetmappingoptions#word-javascript/api/word/-xmlsetmappingoptions-source-member)|If provided, specifies the desired custom XML data to map the content control to.|
