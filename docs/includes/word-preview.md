| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[fontNames](/javascript/api/word/word.application#word-word-application-fontnames-member)|Returns a `FontNameCollection` object that represents all the available font names in Microsoft Word.|
||[listTemplateGalleries](/javascript/api/word/word.application#word-word-application-listtemplategalleries-member)|Returns a `ListTemplateGalleryCollection` object that represents all the list template galleries in Microsoft Word.|
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
|[BorderUniversalCollection](/javascript/api/word/word.borderuniversalcollection)|[getItem(index: number)](/javascript/api/word/word.borderuniversalcollection#word-word-borderuniversalcollection-getitem-member(1))|Gets a `Border` object by its index in the collection.|
|[Coauthor](/javascript/api/word/word.coauthor)|[emailAddress](/javascript/api/word/word.coauthor#word-word-coauthor-emailaddress-member)|Gets the email address of the coauthor.|
||[id](/javascript/api/word/word.coauthor#word-word-coauthor-id-member)|Gets the unique identifier for the `Coauthor` object.|
||[isMe](/javascript/api/word/word.coauthor#word-word-coauthor-isme-member)|Gets whether this author represents the current user.|
||[locks](/javascript/api/word/word.coauthor#word-word-coauthor-locks-member)|Gets a `CoauthoringLockCollection` object that represents the locks in the document that are associated with this coauthor.|
||[name](/javascript/api/word/word.coauthor#word-word-coauthor-name-member)|Gets the display name of the coauthor.|
|[CoauthorCollection](/javascript/api/word/word.coauthorcollection)|[items](/javascript/api/word/word.coauthorcollection#word-word-coauthorcollection-items-member)|Gets the loaded child items in this collection.|
|[Coauthoring](/javascript/api/word/word.coauthoring)|[authors](/javascript/api/word/word.coauthoring#word-word-coauthoring-authors-member)|Gets a `CoauthorCollection` object that represents all the coauthors currently editing the document.|
||[canCoauthor](/javascript/api/word/word.coauthoring#word-word-coauthoring-cancoauthor-member)|Gets whether this document can be coauthored.|
||[canMerge](/javascript/api/word/word.coauthoring#word-word-coauthoring-canmerge-member)|Gets whether the document can be automatically merged.|
||[conflicts](/javascript/api/word/word.coauthoring#word-word-coauthoring-conflicts-member)|Gets a `ConflictCollection` object that represents all the conflicts in the document.|
||[locks](/javascript/api/word/word.coauthoring#word-word-coauthoring-locks-member)|Gets a `CoauthoringLockCollection` object that represents the locks in the document.|
||[me](/javascript/api/word/word.coauthoring#word-word-coauthoring-me-member)|Gets a `Coauthor` object that represents the current user.|
||[pendingUpdates](/javascript/api/word/word.coauthoring#word-word-coauthoring-pendingupdates-member)|Gets whether the document has pending updates that have not been accepted.|
||[updates](/javascript/api/word/word.coauthoring#word-word-coauthoring-updates-member)|Gets a `CoauthoringUpdateCollection` object that represents the most recent updates that were merged into the document.|
|[CoauthoringLock](/javascript/api/word/word.coauthoringlock)|[owner](/javascript/api/word/word.coauthoringlock#word-word-coauthoringlock-owner-member)|Gets the owner of the lock.|
||[range](/javascript/api/word/word.coauthoringlock#word-word-coauthoringlock-range-member)|Gets a `Range` object that represents the portion of the document that's contained in the `CoauthoringLock` object.|
||[type](/javascript/api/word/word.coauthoringlock#word-word-coauthoringlock-type-member)|Gets a `CoauthoringLockType` value that represents the lock type.|
||[unlock()](/javascript/api/word/word.coauthoringlock#word-word-coauthoringlock-unlock-member(1))|Removes this lock, even if it belongs to a different user.|
|[CoauthoringLockAddOptions](/javascript/api/word/word.coauthoringlockaddoptions)|[range](/javascript/api/word/word.coauthoringlockaddoptions#word-word-coauthoringlockaddoptions-range-member)|If provided, specifies the range to which the lock is added.|
||[type](/javascript/api/word/word.coauthoringlockaddoptions#word-word-coauthoringlockaddoptions-type-member)|If provided, specifies the type of lock.|
|[CoauthoringLockCollection](/javascript/api/word/word.coauthoringlockcollection)|[add(options?: Word.CoauthoringLockAddOptions)](/javascript/api/word/word.coauthoringlockcollection#word-word-coauthoringlockcollection-add-member(1))|Returns a `CoauthoringLock` object that represents a lock added to a specified range.|
||[items](/javascript/api/word/word.coauthoringlockcollection#word-word-coauthoringlockcollection-items-member)|Gets the loaded child items in this collection.|
||[unlockEphemeralLocks()](/javascript/api/word/word.coauthoringlockcollection#word-word-coauthoringlockcollection-unlockephemerallocks-member(1))|Removes all ephemeral locks from the document.|
|[CoauthoringUpdate](/javascript/api/word/word.coauthoringupdate)|[range](/javascript/api/word/word.coauthoringupdate#word-word-coauthoringupdate-range-member)|Gets a `Range` object that represents the portion of the document that's contained in the `CoauthoringUpdate` object.|
|[CoauthoringUpdateCollection](/javascript/api/word/word.coauthoringupdatecollection)|[items](/javascript/api/word/word.coauthoringupdatecollection#word-word-coauthoringupdatecollection-items-member)|Gets the loaded child items in this collection.|
|[CommentDetail](/javascript/api/word/word.commentdetail)|[id](/javascript/api/word/word.commentdetail#word-word-commentdetail-id-member)|Represents the ID of this comment.|
||[replyIds](/javascript/api/word/word.commentdetail#word-word-commentdetail-replyids-member)|Represents the IDs of the replies to this comment.|
|[CommentEventArgs](/javascript/api/word/word.commenteventargs)|[changeType](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-changetype-member)|Represents how the `commentChanged` event is raised.|
||[commentDetails](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-commentdetails-member)|Gets the `CommentDetail` array which contains the IDs and reply IDs of the involved comments.|
||[source](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-type-member)|The event type.|
|[ConditionalStyle](/javascript/api/word/word.conditionalstyle)|[borders](/javascript/api/word/word.conditionalstyle#word-word-conditionalstyle-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the conditional style.|
||[bottomPadding](/javascript/api/word/word.conditionalstyle#word-word-conditionalstyle-bottompadding-member)|Specifies the amount of space (in points) to add below the contents of a single cell or all the cells in a table.|
||[font](/javascript/api/word/word.conditionalstyle#word-word-conditionalstyle-font-member)|Returns a `Font` object that represents the font formatting for the conditional style.|
||[leftPadding](/javascript/api/word/word.conditionalstyle#word-word-conditionalstyle-leftpadding-member)|Specifies the amount of space (in points) to add to the left of the contents of a single cell or all the cells in a table.|
||[paragraphFormat](/javascript/api/word/word.conditionalstyle#word-word-conditionalstyle-paragraphformat-member)|Returns a `ParagraphFormat` object that represents the paragraph formatting for the conditional style.|
||[rightPadding](/javascript/api/word/word.conditionalstyle#word-word-conditionalstyle-rightpadding-member)|Specifies the amount of space (in points) to add to the right of the contents of a single cell or all the cells in a table.|
||[shading](/javascript/api/word/word.conditionalstyle#word-word-conditionalstyle-shading-member)|Returns a `ShadingUniversal` object that represents the shading of the conditional style.|
||[topPadding](/javascript/api/word/word.conditionalstyle#word-word-conditionalstyle-toppadding-member)|Specifies the amount of space (in points) to add above the contents of a single cell or all the cells in a table.|
|[Conflict](/javascript/api/word/word.conflict)|[accept()](/javascript/api/word/word.conflict#word-word-conflict-accept-member(1))|Accepts the user's change and removes the conflict.|
||[range](/javascript/api/word/word.conflict#word-word-conflict-range-member)|Gets a `Range` object that represents the portion of the document that's contained in the `Conflict` object.|
||[reject()](/javascript/api/word/word.conflict#word-word-conflict-reject-member(1))|Rejects the user's change, removes the conflict, and accepts the server copy of the change for the conflict.|
||[type](/javascript/api/word/word.conflict#word-word-conflict-type-member)|Gets the `RevisionType` for the `Conflict` object.|
|[ConflictCollection](/javascript/api/word/word.conflictcollection)|[acceptAll()](/javascript/api/word/word.conflictcollection#word-word-conflictcollection-acceptall-member(1))|Accepts all of the user's changes, removes the conflicts, and merges the changes into the server copy of the document.|
||[getItem(index: number)](/javascript/api/word/word.conflictcollection#word-word-conflictcollection-getitem-member(1))|Gets a `Conflict` object by its index in the collection.|
||[items](/javascript/api/word/word.conflictcollection#word-word-conflictcollection-items-member)|Gets the loaded child items in this collection.|
||[rejectAll()](/javascript/api/word/word.conflictcollection#word-word-conflictcollection-rejectall-member(1))|Rejects all of the user's changes and retains the server copy of the document.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onCommentAdded](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentselected-member)|Occurs when a comment is selected.|
||[resetState()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-resetstate-member(1))|Resets the state of the content control.|
||[setState(contentControlState: Word.ContentControlState)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-setstate-member(1))|Sets the state of the content control.|
|[CustomXmlAddValidationErrorOptions](/javascript/api/word/word.customxmladdvalidationerroroptions)|[clearedOnUpdate](/javascript/api/word/word.customxmladdvalidationerroroptions#word-word-customxmladdvalidationerroroptions-clearedonupdate-member)|If provided, specifies whether the error is to be cleared from the Word.CustomXmlValidationErrorCollection when the XML is corrected and updated.|
||[errorText](/javascript/api/word/word.customxmladdvalidationerroroptions#word-word-customxmladdvalidationerroroptions-errortext-member)|If provided, specifies the descriptive error text.|
|[CustomXmlNodeCollection](/javascript/api/word/word.customxmlnodecollection)|[getItem(index: number)](/javascript/api/word/word.customxmlnodecollection#word-word-customxmlnodecollection-getitem-member(1))|Returns a `CustomXmlNode` object that represents the specified item in the collection.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[errors](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-errors-member)|Gets a `CustomXmlValidationErrorCollection` object that provides access to any XML validation errors.|
|[CustomXmlPrefixMappingCollection](/javascript/api/word/word.customxmlprefixmappingcollection)|[getItem(index: number)](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-getitem-member(1))|Returns a `CustomXmlPrefixMapping` object that represents the specified item in the collection.|
|[CustomXmlValidationError](/javascript/api/word/word.customxmlvalidationerror)|[delete()](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-delete-member(1))|Deletes this `CustomXmlValidationError` object.|
||[errorCode](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-errorcode-member)|Gets an integer representing the validation error in the `CustomXmlValidationError` object.|
||[name](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-name-member)|Gets the name of the error in the `CustomXmlValidationError` object.|
||[node](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-node-member)|Gets the node associated with this `CustomXmlValidationError` object, if any exist.|
||[text](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-text-member)|Gets the text in the `CustomXmlValidationError` object.|
||[type](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-type-member)|Gets the type of error generated from the `CustomXmlValidationError` object.|
|[CustomXmlValidationErrorCollection](/javascript/api/word/word.customxmlvalidationerrorcollection)|[add(node: Word.CustomXmlNode, errorName: string, options?: Word.CustomXmlAddValidationErrorOptions)](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-add-member(1))|Adds a `CustomXmlValidationError` object containing an XML validation error to the `CustomXmlValidationErrorCollection` object.|
||[getCount()](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-getitem-member(1))|Returns a `CustomXmlValidationError` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[acceptAllRevisions()](/javascript/api/word/word.document#word-word-document-acceptallrevisions-member(1))|Accepts all tracked changes in the document.|
||[acceptAllRevisionsShown()](/javascript/api/word/word.document#word-word-document-acceptallrevisionsshown-member(1))|Accepts all revisions in the document that are displayed on the screen.|
||[activate()](/javascript/api/word/word.document#word-word-document-activate-member(1))|Activates the document so that it becomes the active document.|
||[activeTheme](/javascript/api/word/word.document#word-word-document-activetheme-member)|Gets the name of the active theme and formatting options.|
||[activeThemeDisplayName](/javascript/api/word/word.document#word-word-document-activethemedisplayname-member)|Gets the display name of the active theme.|
||[addToFavorites()](/javascript/api/word/word.document#word-word-document-addtofavorites-member(1))|Creates a shortcut to the document or hyperlink and adds it to the **Favorites** folder.|
||[applyQuickStyleSet(style: Word.ApplyQuickStyleSet)](/javascript/api/word/word.document#word-word-document-applyquickstyleset-member(1))|Applies the specified Quick Style set to the document.|
||[applyTheme(name: string)](/javascript/api/word/word.document#word-word-document-applytheme-member(1))|Applies the specified theme to this document.|
||[areFilePropertiesPasswordEncrypted](/javascript/api/word/word.document#word-word-document-arefilepropertiespasswordencrypted-member)|Gets whether Word encrypts file properties for password-protected documents.|
||[areGrammaticalErrorsShown](/javascript/api/word/word.document#word-word-document-aregrammaticalerrorsshown-member)|Specifies whether grammatical errors are marked by a wavy green line in the document.|
||[areMathDefaultsUsed](/javascript/api/word/word.document#word-word-document-aremathdefaultsused-member)|Specifies whether to use the default math settings when creating new equations.|
||[areNewerFeaturesDisabled](/javascript/api/word/word.document#word-word-document-arenewerfeaturesdisabled-member)|Specifies whether to disable features introduced after a specified version.|
||[areSpellingErrorsShown](/javascript/api/word/word.document#word-word-document-arespellingerrorsshown-member)|Specifies whether Microsoft Word underlines spelling errors in the document.|
||[areStylesUpdatedOnOpen](/javascript/api/word/word.document#word-word-document-arestylesupdatedonopen-member)|Specifies whether the styles in this document are updated to match the styles in the attached template each time the document is opened.|
||[areTrueTypeFontsEmbedded](/javascript/api/word/word.document#word-word-document-aretruetypefontsembedded-member)|Specifies whether Microsoft Word embeds TrueType fonts in the document when it's saved.|
||[autoFormat()](/javascript/api/word/word.document#word-word-document-autoformat-member(1))|Automatically formats the document.|
||[background](/javascript/api/word/word.document#word-word-document-background-member)|Gets a `Shape` object that represents the background image for the document.|
||[bookmarks](/javascript/api/word/word.document#word-word-document-bookmarks-member)|Returns a `BookmarkCollection` object that represents all the bookmarks in the document.|
||[builtInDocumentProperties](/javascript/api/word/word.document#word-word-document-builtindocumentproperties-member)|Gets a `DocumentProperties` object that represents all the built-in document properties for the document.|
||[canCheckin()](/javascript/api/word/word.document#word-word-document-cancheckin-member(1))|Returns `true` if Microsoft Word can check in the document to a server.|
||[characters](/javascript/api/word/word.document#word-word-document-characters-member)|Gets the `RangeScopedCollection` object that represents all the characters in the document.|
||[checkConsistencyJapanese()](/javascript/api/word/word.document#word-word-document-checkconsistencyjapanese-member(1))|Searches all text in a Japanese language document and displays instances where character usage is inconsistent for the same words.|
||[checkGrammar()](/javascript/api/word/word.document#word-word-document-checkgrammar-member(1))|Begins a spelling and grammar check for the document.|
||[checkIn(options?: Word.DocumentCheckInOptions)](/javascript/api/word/word.document#word-word-document-checkin-member(1))|Checks in the document from the local computer to a server and sets the local document to read-only so that it cannot be edited locally.|
||[checkInWithVersion(options?: Word.DocumentCheckInWithVersionOptions)](/javascript/api/word/word.document#word-word-document-checkinwithversion-member(1))|Saves the document to a server from a local computer, and sets the local document to read-only so that it cannot be edited locally.|
||[checkSpelling(options?: Word.DocumentCheckSpellingOptions)](/javascript/api/word/word.document#word-word-document-checkspelling-member(1))|Begins a spelling check for the document.|
||[closePrintPreview()](/javascript/api/word/word.document#word-word-document-closeprintpreview-member(1))|Switches the document from print preview to the previous view.|
||[coauthoring](/javascript/api/word/word.document#word-word-document-coauthoring-member)|Gets a `Coauthoring` object for managing coauthoring in the document.|
||[codeName](/javascript/api/word/word.document#word-word-document-codename-member)|Gets the code name for the document.|
||[comments](/javascript/api/word/word.document#word-word-document-comments-member)|Gets a `CommentCollection` object that represents all the comments in the document.|
||[compatibilityMode](/javascript/api/word/word.document#word-word-document-compatibilitymode-member)|Gets the compatibility mode that Word uses when opening the document.|
||[computeStatistics(statistic: Word.Statistic, includeFootnotesAndEndnotes?: boolean)](/javascript/api/word/word.document#word-word-document-computestatistics-member(1))|Returns a statistic based on the contents of the document.|
||[content](/javascript/api/word/word.document#word-word-document-content-member)|Gets a `Range` object that represents the main document story.|
||[convert()](/javascript/api/word/word.document#word-word-document-convert-member(1))|Converts the file to the newest format and enables all features.|
||[convertAutoHyphens()](/javascript/api/word/word.document#word-word-document-convertautohyphens-member(1))|Converts automatic hyphens to manual hyphens.|
||[convertNumbersToText(numberType?: Word.NumberType)](/javascript/api/word/word.document#word-word-document-convertnumberstotext-member(1))|Changes the list numbers and LISTNUM fields in the document to text.|
||[convertVietnameseDocument(codePageOrigin: number)](/javascript/api/word/word.document#word-word-document-convertvietnamesedocument-member(1))|Reconverts a Vietnamese document to Unicode using a code page other than the default.|
||[copyStylesFromTemplate(StyleTemplate: string)](/javascript/api/word/word.document#word-word-document-copystylesfromtemplate-member(1))|Copies styles from the specified template to the document.|
||[countNumberedItems(options?: Word.DocumentCountNumberedItemsOptions)](/javascript/api/word/word.document#word-word-document-countnumbereditems-member(1))|Returns the number of bulleted or numbered items and LISTNUM fields in the document.|
||[currentRsid](/javascript/api/word/word.document#word-word-document-currentrsid-member)|Gets a random number that Word assigns to changes in the document.|
||[customDocumentProperties](/javascript/api/word/word.document#word-word-document-customdocumentproperties-member)|Gets a `DocumentProperties` collection that represents all the custom document properties for the document.|
||[defaultTabStop](/javascript/api/word/word.document#word-word-document-defaulttabstop-member)|Specifies the interval (in points) between the default tab stops in the document.|
||[defaultTargetFrame](/javascript/api/word/word.document#word-word-document-defaulttargetframe-member)|Specifies the browser frame for displaying a webpage via hyperlink.|
||[deleteAllComments()](/javascript/api/word/word.document#word-word-document-deleteallcomments-member(1))|Deletes all comments from the document.|
||[deleteAllCommentsShown()](/javascript/api/word/word.document#word-word-document-deleteallcommentsshown-member(1))|Deletes all revisions in the document that are displayed on the screen.|
||[deleteAllInkAnnotations()](/javascript/api/word/word.document#word-word-document-deleteallinkannotations-member(1))|Deletes all handwritten ink annotations in the document.|
||[doNotEmbedSystemFonts](/javascript/api/word/word.document#word-word-document-donotembedsystemfonts-member)|Specifies whether Word should not embed common system fonts.|
||[encryptionProvider](/javascript/api/word/word.document#word-word-document-encryptionprovider-member)|Specifies the name of the algorithm encryption provider that Microsoft Word uses when encrypting documents.|
||[endReview(options?: Word.DocumentEndReviewOptions)](/javascript/api/word/word.document#word-word-document-endreview-member(1))|Terminates a review of the file that has been sent for review.|
||[exportAsFixedFormat(outputFileName: string, exportFormat: Word.ExportFormat, options?: Word.DocumentExportAsFixedFormatOptions)](/javascript/api/word/word.document#word-word-document-exportasfixedformat-member(1))|Saves the document in PDF or XPS format.|
||[exportAsFixedFormat2(outputFileName: string, exportFormat: Word.ExportFormat, options?: Word.DocumentExportAsFixedFormat2Options)](/javascript/api/word/word.document#word-word-document-exportasfixedformat2-member(1))|Saves the document in PDF or XPS format.|
||[exportAsFixedFormat3(outputFileName: string, exportFormat: Word.ExportFormat, options?: Word.DocumentExportAsFixedFormat3Options)](/javascript/api/word/word.document#word-word-document-exportasfixedformat3-member(1))|Saves the document in PDF or XPS format with improved tagging.|
||[farEastLineBreakLanguage](/javascript/api/word/word.document#word-word-document-fareastlinebreaklanguage-member)|Specifies the East Asian language used for line breaking.|
||[farEastLineBreakLevel](/javascript/api/word/word.document#word-word-document-fareastlinebreaklevel-member)|Specifies the line break control level.|
||[fields](/javascript/api/word/word.document#word-word-document-fields-member)|Gets a `FieldCollection` object that represents all the fields in the document.|
||[fitToPages()](/javascript/api/word/word.document#word-word-document-fittopages-member(1))|Decreases the font size of text just enough so that the document page count drops by one.|
||[followHyperlink(options?: Word.DocumentFollowHyperlinkOptions)](/javascript/api/word/word.document#word-word-document-followhyperlink-member(1))|Displays a cached document, if it has already been downloaded.|
||[formattingIsNextLevelShown](/javascript/api/word/word.document#word-word-document-formattingisnextlevelshown-member)|Specifies whether Word shows the next heading level when the previous is used.|
||[formattingIsUserStyleNameShown](/javascript/api/word/word.document#word-word-document-formattingisuserstylenameshown-member)|Specifies whether to show user-defined styles.|
||[freezeLayout()](/javascript/api/word/word.document#word-word-document-freezelayout-member(1))|Fixes the layout of the document in Web view.|
||[fullName](/javascript/api/word/word.document#word-word-document-fullname-member)|Gets the name of a document, including the path.|
||[getCrossReferenceItems(referenceType: Word.ReferenceType)](/javascript/api/word/word.document#word-word-document-getcrossreferenceitems-member(1))|Returns an array of items that can be cross-referenced based on the specified cross-reference type.|
||[getRange(options?: Word.DocumentRangeOptions)](/javascript/api/word/word.document#word-word-document-getrange-member(1))|Returns a `Range` object by using the specified starting and ending character positions.|
||[goTo(options?: Word.GoToOptions)](/javascript/api/word/word.document#word-word-document-goto-member(1))|Returns a `Range` object that represents the start position of the specified item, such as a page, bookmark, or field.|
||[grammaticalErrors](/javascript/api/word/word.document#word-word-document-grammaticalerrors-member)|Gets a `RangeCollection` object that represents the sentences that failed the grammar check in the document.|
||[gridDistanceHorizontal](/javascript/api/word/word.document#word-word-document-griddistancehorizontal-member)|Specifies the horizontal space between invisible gridlines that Microsoft Word uses when you draw, move, and resize AutoShapes or East Asian characters in the document.|
||[gridDistanceVertical](/javascript/api/word/word.document#word-word-document-griddistancevertical-member)|Specifies the vertical space between invisible gridlines that Microsoft Word uses when you draw, move, and resize AutoShapes or East Asian characters in the document.|
||[gridIsOriginFromMargin](/javascript/api/word/word.document#word-word-document-gridisoriginfrommargin-member)|Specifies whether the character grid starts from the upper-left corner of the page.|
||[gridOriginHorizontal](/javascript/api/word/word.document#word-word-document-gridoriginhorizontal-member)|Specifies the horizontal origin point for the invisible grid.|
||[gridOriginVertical](/javascript/api/word/word.document#word-word-document-gridoriginvertical-member)|Specifies the vertical origin point for the invisible grid.|
||[gridSpaceBetweenHorizontalLines](/javascript/api/word/word.document#word-word-document-gridspacebetweenhorizontallines-member)|Specifies the interval for horizontal character gridlines in print layout view.|
||[gridSpaceBetweenVerticalLines](/javascript/api/word/word.document#word-word-document-gridspacebetweenverticallines-member)|Specifies the interval for vertical character gridlines in print layout view.|
||[hasPassword](/javascript/api/word/word.document#word-word-document-haspassword-member)|Gets whether a password is required to open the document.|
||[hasVbProject](/javascript/api/word/word.document#word-word-document-hasvbproject-member)|Gets whether the document has an attached Microsoft Visual Basic for Applications project.|
||[hyphenationZone](/javascript/api/word/word.document#word-word-document-hyphenationzone-member)|Specifies the width of the hyphenation zone, in points.|
||[indexes](/javascript/api/word/word.document#word-word-document-indexes-member)|Returns an `IndexCollection` object that represents all the indexes in the document.|
||[isAutoFormatOverrideOn](/javascript/api/word/word.document#word-word-document-isautoformatoverrideon-member)|Specifies whether automatic formatting options override formatting restrictions.|
||[isChartDataPointTracked](/javascript/api/word/word.document#word-word-document-ischartdatapointtracked-member)|Specifies whether charts in the active document use cell-reference data-point tracking.|
||[isCompatible](/javascript/api/word/word.document#word-word-document-iscompatible-member)|Specifies whether the compatibility option specified by the `type` property is enabled.|
||[isFinal](/javascript/api/word/word.document#word-word-document-isfinal-member)|Specifies whether the document is final.|
||[isFontsSubsetSaved](/javascript/api/word/word.document#word-word-document-isfontssubsetsaved-member)|Specifies whether Microsoft Word saves a subset of the embedded TrueType fonts with the document.|
||[isFormsDataPrinted](/javascript/api/word/word.document#word-word-document-isformsdataprinted-member)|Specifies whether Microsoft Word prints onto a preprinted form only the data entered in the corresponding online form.|
||[isFormsDataSaved](/javascript/api/word/word.document#word-word-document-isformsdatasaved-member)|Specifies whether Microsoft Word saves the data entered in a form as a tab-delimited record for use in a database.|
||[isGrammarChecked](/javascript/api/word/word.document#word-word-document-isgrammarchecked-member)|Specifies whether a grammar check has been run on the document.|
||[isInAutoSave](/javascript/api/word/word.document#word-word-document-isinautosave-member)|Gets whether the most recent firing of the `Application.DocumentBeforeSave` event was the result of an automatic save by the document or a manual save by the user.|
||[isInFormsDesign](/javascript/api/word/word.document#word-word-document-isinformsdesign-member)|Gets whether the document is in form design mode.|
||[isKerningByAlgorithm](/javascript/api/word/word.document#word-word-document-iskerningbyalgorithm-member)|Specifies whether Word kerns half-width Latin characters and punctuation marks.|
||[isLinguisticDataEmbedded](/javascript/api/word/word.document#word-word-document-islinguisticdataembedded-member)|Specifies whether to embed speech and handwriting data.|
||[isMasterDocument](/javascript/api/word/word.document#word-word-document-ismasterdocument-member)|Gets whether this document is a master document.|
||[isOptimizedForWord97](/javascript/api/word/word.document#word-word-document-isoptimizedforword97-member)|Specifies whether Word optimizes the document for Word 97.|
||[isPostScriptPrintedOverText](/javascript/api/word/word.document#word-word-document-ispostscriptprintedovertext-member)|Specifies whether PRINT field instructions (such as PostScript commands) in the document are to be printed on top of text and graphics when a PostScript printer is used.|
||[isQuickStyleSetLocked](/javascript/api/word/word.document#word-word-document-isquickstylesetlocked-member)|Specifies whether users can change the Quick Style set.|
||[isReadOnly](/javascript/api/word/word.document#word-word-document-isreadonly-member)|Gets whether changes to the document cannot be saved to the original document.|
||[isReadOnlyRecommended](/javascript/api/word/word.document#word-word-document-isreadonlyrecommended-member)|Specifies whether Microsoft Word displays a message box whenever a user opens the document, suggesting that it be opened as read-only.|
||[isSnappedToGrid](/javascript/api/word/word.document#word-word-document-issnappedtogrid-member)|Specifies whether AutoShapes or East Asian characters are automatically aligned with an invisible grid.|
||[isSnappedToShapes](/javascript/api/word/word.document#word-word-document-issnappedtoshapes-member)|Specifies whether AutoShapes or East Asian characters align with invisible gridlines through other shapes.|
||[isSpellingChecked](/javascript/api/word/word.document#word-word-document-isspellingchecked-member)|Specifies whether spelling has been checked throughout the document.|
||[isStyleEnforced](/javascript/api/word/word.document#word-word-document-isstyleenforced-member)|Specifies whether formatting restrictions are enforced in a protected document.|
||[isSubdocument](/javascript/api/word/word.document#word-word-document-issubdocument-member)|Gets whether this document is a subdocument of a master document.|
||[isThemeLocked](/javascript/api/word/word.document#word-word-document-isthemelocked-member)|Specifies whether users can change the document theme.|
||[isUserControl](/javascript/api/word/word.document#word-word-document-isusercontrol-member)|Specifies whether the document was created or opened by the user.|
||[isVbaSigned](/javascript/api/word/word.document#word-word-document-isvbasigned-member)|Gets whether the VBA project is digitally signed.|
||[isWriteReserved](/javascript/api/word/word.document#word-word-document-iswritereserved-member)|Gets whether the document is protected with a write password.|
||[justificationMode](/javascript/api/word/word.document#word-word-document-justificationmode-member)|Specifies the character spacing adjustment.|
||[kind](/javascript/api/word/word.document#word-word-document-kind-member)|Specifies the format type that Microsoft Word uses when automatically formatting the document.|
||[listParagraphs](/javascript/api/word/word.document#word-word-document-listparagraphs-member)|Gets a `ParagraphCollection` object that represents all the numbered paragraphs in the document.|
||[listTemplates](/javascript/api/word/word.document#word-word-document-listtemplates-member)|Returns a `ListTemplateCollection` object that represents all the list templates in the document.|
||[lists](/javascript/api/word/word.document#word-word-document-lists-member)|Gets a `ListCollection` object that contains all the formatted lists in the document.|
||[lockServerFile()](/javascript/api/word/word.document#word-word-document-lockserverfile-member(1))|Locks the file on the server preventing anyone else from editing it.|
||[makeCompatibilityDefault()](/javascript/api/word/word.document#word-word-document-makecompatibilitydefault-member(1))|Sets the compatibility options.|
||[merge(fileName: string, options?: Word.DocumentMergeOptions)](/javascript/api/word/word.document#word-word-document-merge-member(1))|Merges the changes marked with revision marks from one document to another.|
||[noLineBreakAfter](/javascript/api/word/word.document#word-word-document-nolinebreakafter-member)|Specifies the kinsoku characters after which Word will not break a line.|
||[noLineBreakBefore](/javascript/api/word/word.document#word-word-document-nolinebreakbefore-member)|Specifies the kinsoku characters before which Word will not break a line.|
||[openEncoding](/javascript/api/word/word.document#word-word-document-openencoding-member)|Gets the encoding used to open the document.|
||[originalDocumentTitle](/javascript/api/word/word.document#word-word-document-originaldocumenttitle-member)|Gets the title of the original document after legal-blackline comparison.|
||[paragraphs](/javascript/api/word/word.document#word-word-document-paragraphs-member)|Gets a `ParagraphCollection` object that represents all the paragraphs in the document.|
||[password](/javascript/api/word/word.document#word-word-document-password-member)|Sets a password that must be supplied to open the document.|
||[passwordEncryptionAlgorithm](/javascript/api/word/word.document#word-word-document-passwordencryptionalgorithm-member)|Gets the algorithm used for password encryption.|
||[passwordEncryptionKeyLength](/javascript/api/word/word.document#word-word-document-passwordencryptionkeylength-member)|Gets the key length used for password encryption.|
||[passwordEncryptionProvider](/javascript/api/word/word.document#word-word-document-passwordencryptionprovider-member)|Gets the name of the password encryption provider.|
||[path](/javascript/api/word/word.document#word-word-document-path-member)|Gets the disk or the web path to the document (excludes the document name).|
||[post()](/javascript/api/word/word.document#word-word-document-post-member(1))|Posts the document to a public folder in Microsoft Exchange.|
||[presentIt()](/javascript/api/word/word.document#word-word-document-presentit-member(1))|Opens PowerPoint with the Word document loaded.|
||[printOut(options?: Word.DocumentPrintOutOptions)](/javascript/api/word/word.document#word-word-document-printout-member(1))|Prints all or part of the document.|
||[printPreview()](/javascript/api/word/word.document#word-word-document-printpreview-member(1))|Switches the view to print preview.|
||[printRevisions](/javascript/api/word/word.document#word-word-document-printrevisions-member)|Specifies whether revision marks are printed with the document.|
||[protect(type: Word.ProtectionType, options?: Word.DocumentProtectOptions)](/javascript/api/word/word.document#word-word-document-protect-member(1))|Protects the document from unauthorized changes.|
||[protectionType](/javascript/api/word/word.document#word-word-document-protectiontype-member)|Gets the protection type for the document.|
||[readabilityStatistics](/javascript/api/word/word.document#word-word-document-readabilitystatistics-member)|Gets a `ReadabilityStatisticCollection` object that represents the readability statistics for the document.|
||[readingLayoutSizeX](/javascript/api/word/word.document#word-word-document-readinglayoutsizex-member)|Specifies the width of pages in reading layout view when frozen.|
||[readingLayoutSizeY](/javascript/api/word/word.document#word-word-document-readinglayoutsizey-member)|Specifies the height of pages in reading layout view when frozen.|
||[readingModeIsLayoutFrozen](/javascript/api/word/word.document#word-word-document-readingmodeislayoutfrozen-member)|Specifies whether pages in reading layout view are frozen for handwritten markup.|
||[redo(times?: number)](/javascript/api/word/word.document#word-word-document-redo-member(1))|Redoes the last action that was undone (reverses the `undo` method).|
||[rejectAllRevisions()](/javascript/api/word/word.document#word-word-document-rejectallrevisions-member(1))|Rejects all tracked changes in the document.|
||[rejectAllRevisionsShown()](/javascript/api/word/word.document#word-word-document-rejectallrevisionsshown-member(1))|Rejects all revisions in the document that are displayed on the screen.|
||[reload()](/javascript/api/word/word.document#word-word-document-reload-member(1))|Reloads a cached document by resolving the hyperlink to the document and downloading it.|
||[reloadAs(encoding: Word.DocumentEncoding)](/javascript/api/word/word.document#word-word-document-reloadas-member(1))|Reloads the document based on an HTML document, using the document encoding.|
||[removeDocumentInformation(removeDocInfoType: Word.RemoveDocInfoType)](/javascript/api/word/word.document#word-word-document-removedocumentinformation-member(1))|Removes sensitive information, properties, comments, and other metadata from the document.|
||[removeLockedStyles()](/javascript/api/word/word.document#word-word-document-removelockedstyles-member(1))|Purges the document of locked styles when formatting restrictions have been applied in the document.|
||[removeNumbers(numberType?: Word.NumberType)](/javascript/api/word/word.document#word-word-document-removenumbers-member(1))|Removes numbers or bullets from the document.|
||[removePersonalInformationOnSave](/javascript/api/word/word.document#word-word-document-removepersonalinformationonsave-member)|Specifies whether Word removes user information upon saving.|
||[removeTheme()](/javascript/api/word/word.document#word-word-document-removetheme-member(1))|Removes the active theme from the current document.|
||[repaginate()](/javascript/api/word/word.document#word-word-document-repaginate-member(1))|Repaginates the entire document.|
||[replyWithChanges(options?: Word.DocumentReplyWithChangesOptions)](/javascript/api/word/word.document#word-word-document-replywithchanges-member(1))|Sends an email message to the author of the document that has been sent out for review, notifying them that a reviewer has completed review of the document.|
||[resetFormFields()](/javascript/api/word/word.document#word-word-document-resetformfields-member(1))|Clears all form fields in the document, preparing the form to be filled in again.|
||[returnToLastReadPosition()](/javascript/api/word/word.document#word-word-document-returntolastreadposition-member(1))|Returns the document to the last saved reading position.|
||[revisedDocumentTitle](/javascript/api/word/word.document#word-word-document-reviseddocumenttitle-member)|Gets the title of the revised document after legal-blackline comparison.|
||[revisions](/javascript/api/word/word.document#word-word-document-revisions-member)|Gets the collection of revisions that represents the tracked changes in the document.|
||[runAutoMacro(autoMacro: Word.AutoMacro)](/javascript/api/word/word.document#word-word-document-runautomacro-member(1))|Runs an auto macro that's stored in the document.|
||[saveAsQuickStyleSet(fileName: string)](/javascript/api/word/word.document#word-word-document-saveasquickstyleset-member(1))|Saves the group of quick styles currently in use.|
||[saveEncoding](/javascript/api/word/word.document#word-word-document-saveencoding-member)|Specifies the encoding used when saving the document.|
||[saveFormat](/javascript/api/word/word.document#word-word-document-saveformat-member)|Gets the file format of the document.|
||[select()](/javascript/api/word/word.document#word-word-document-select-member(1))|Selects the contents of the document.|
||[selectContentControlsByTag(tag: string)](/javascript/api/word/word.document#word-word-document-selectcontentcontrolsbytag-member(1))|Returns all content controls with the specified tag.|
||[selectContentControlsByTitle(title: string)](/javascript/api/word/word.document#word-word-document-selectcontentcontrolsbytitle-member(1))|Returns a `ContentControlCollection` object that represents all the content controls in the document with the specified title.|
||[selectLinkedControls(node: Word.CustomXmlNode)](/javascript/api/word/word.document#word-word-document-selectlinkedcontrols-member(1))|Returns a `ContentControlCollection` object that represents all content controls in the document that are linked to the specific custom XML node.|
||[selectNodes(xPath: string, options?: Word.SelectNodesOptions)](/javascript/api/word/word.document#word-word-document-selectnodes-member(1))|Returns an `XmlNodeCollection` object that represents all the nodes that match the XPath parameter in the order in which they appear in the document.|
||[selectSingleNode(xPath: string, options?: Word.SelectSingleNodeOptions)](/javascript/api/word/word.document#word-word-document-selectsinglenode-member(1))|Returns an `XmlNode` object that represents the first node that matches the XPath parameter in the document.|
||[selectUnlinkedControls(stream?: Word.CustomXmlPart)](/javascript/api/word/word.document#word-word-document-selectunlinkedcontrols-member(1))|Returns a `ContentControlCollection` object that represents all content controls in the document that are not linked to an XML node.|
||[selection](/javascript/api/word/word.document#word-word-document-selection-member)|Returns a `Selection` object that represents the current selection in the document.|
||[sendFax(address: string, subject?: string)](/javascript/api/word/word.document#word-word-document-sendfax-member(1))|Sends the document as a fax, without any user interaction.|
||[sendFaxOverInternet(options?: Word.DocumentSendFaxOverInternetOptions)](/javascript/api/word/word.document#word-word-document-sendfaxoverinternet-member(1))|Sends the document to a fax service provider, who faxes the document to one or more specified recipients.|
||[sendForReview(options?: Word.DocumentSendForReviewOptions)](/javascript/api/word/word.document#word-word-document-sendforreview-member(1))|Sends the document in an email message for review by the specified recipients.|
||[sendMail()](/javascript/api/word/word.document#word-word-document-sendmail-member(1))|Opens a message window for sending the document through Microsoft Exchange.|
||[sentences](/javascript/api/word/word.document#word-word-document-sentences-member)|Gets the `RangeScopedCollection` object that represents all the sentences in the document.|
||[setDefaultTableStyle(style: string, setInTemplate: boolean)](/javascript/api/word/word.document#word-word-document-setdefaulttablestyle-member(1))|Specifies the table style to use for newly created tables in the document.|
||[setPasswordEncryptionOptions(passwordEncryptionProvider: string, passwordEncryptionAlgorithm: string, passwordEncryptionKeyLength: number, passwordEncryptFileProperties?: boolean)](/javascript/api/word/word.document#word-word-document-setpasswordencryptionoptions-member(1))|Sets the options Microsoft Word uses for encrypting documents with passwords.|
||[spellingErrors](/javascript/api/word/word.document#word-word-document-spellingerrors-member)|Gets a `RangeCollection` object that represents the words identified as spelling errors in the document.|
||[storyRanges](/javascript/api/word/word.document#word-word-document-storyranges-member)|Gets a `RangeCollection` object that represents all the stories in the document.|
||[styles](/javascript/api/word/word.document#word-word-document-styles-member)|Gets a `StyleCollection` for the document.|
||[tableOfAuthoritiesCategories](/javascript/api/word/word.document#word-word-document-tableofauthoritiescategories-member)|Returns a `TableOfAuthoritiesCategoryCollection` object that represents the available table of authorities categories in the document.|
||[tables](/javascript/api/word/word.document#word-word-document-tables-member)|Gets a `TableCollection` object that represents all the tables in the document.|
||[tablesOfAuthorities](/javascript/api/word/word.document#word-word-document-tablesofauthorities-member)|Returns a `TableOfAuthoritiesCollection` object that represents all the tables of authorities in the document.|
||[tablesOfContents](/javascript/api/word/word.document#word-word-document-tablesofcontents-member)|Returns a `TableOfContentsCollection` object that represents all the tables of contents in the document.|
||[tablesOfFigures](/javascript/api/word/word.document#word-word-document-tablesoffigures-member)|Returns a `TableOfFiguresCollection` object that represents all the tables of figures in the document.|
||[textEncoding](/javascript/api/word/word.document#word-word-document-textencoding-member)|Specifies the encoding for saving as encoded text.|
||[textLineEnding](/javascript/api/word/word.document#word-word-document-textlineending-member)|Specifies how Word marks line and paragraph breaks in text files.|
||[toggleFormsDesign()](/javascript/api/word/word.document#word-word-document-toggleformsdesign-member(1))|Switches form design mode on or off.|
||[trackFormatting](/javascript/api/word/word.document#word-word-document-trackformatting-member)|Specifies whether to track formatting changes when change tracking is on.|
||[trackMoves](/javascript/api/word/word.document#word-word-document-trackmoves-member)|Specifies whether to mark moved text when Track Changes is on.|
||[trackRevisions](/javascript/api/word/word.document#word-word-document-trackrevisions-member)|Specifies whether changes are tracked in the document.|
||[trackedChangesAreDateAndTimeRemoved](/javascript/api/word/word.document#word-word-document-trackedchangesaredateandtimeremoved-member)|Specifies whether to remove or store date and time metadata for tracked changes.|
||[transformDocument(path: string, dataOnly?: boolean)](/javascript/api/word/word.document#word-word-document-transformdocument-member(1))|Applies the specified Extensible Stylesheet Language Transformation (XSLT) file to this document and replaces the document with the results.|
||[type](/javascript/api/word/word.document#word-word-document-type-member)|Gets the document type (template or document).|
||[undo(times?: number)](/javascript/api/word/word.document#word-word-document-undo-member(1))|Undoes the last action or a sequence of actions, which are displayed in the Undo list.|
||[undoClear()](/javascript/api/word/word.document#word-word-document-undoclear-member(1))|Clears the list of actions that can be undone in the document.|
||[unprotect(password?: string)](/javascript/api/word/word.document#word-word-document-unprotect-member(1))|Removes protection from the document.|
||[updateStyles()](/javascript/api/word/word.document#word-word-document-updatestyles-member(1))|Copies all styles from the attached template into the document, overwriting any existing styles in the document that have the same name.|
||[viewCode()](/javascript/api/word/word.document#word-word-document-viewcode-member(1))|Displays the code window for the selected Microsoft ActiveX control in the document.|
||[viewPropertyBrowser()](/javascript/api/word/word.document#word-word-document-viewpropertybrowser-member(1))|Displays the property window for the selected Microsoft ActiveX control in the document.|
||[webPagePreview()](/javascript/api/word/word.document#word-word-document-webpagepreview-member(1))|Displays a preview of the current document as it would look if saved as a webpage.|
||[webSettings](/javascript/api/word/word.document#word-word-document-websettings-member)|Gets the `WebSettings` object for webpage-related attributes.|
||[words](/javascript/api/word/word.document#word-word-document-words-member)|Gets the `RangeScopedCollection` object that represents each word in the document.|
||[writePassword](/javascript/api/word/word.document#word-word-document-writepassword-member)|Sets a password for saving changes to the document.|
||[xmlAreAdvancedErrorsShown](/javascript/api/word/word.document#word-word-document-xmlareadvancederrorsshown-member)|Specifies whether error messages are generated from built-in Word messages or MSXML (Microsoft XML).|
||[xmlIsXsltUsedWhenSaving](/javascript/api/word/word.document#word-word-document-xmlisxsltusedwhensaving-member)|Specifies whether to save a document through an Extensible Stylesheet Language Transformation (XSLT).|
||[xmlSaveThroughXslt](/javascript/api/word/word.document#word-word-document-xmlsavethroughxslt-member)|Specifies the path and file name for the XSLT to apply when saving a document.|
|[DocumentCheckInOptions](/javascript/api/word/word.documentcheckinoptions)|[comment](/javascript/api/word/word.documentcheckinoptions#word-word-documentcheckinoptions-comment-member)|If provided, specifies a comment for the check-in operation.|
||[makePublic](/javascript/api/word/word.documentcheckinoptions#word-word-documentcheckinoptions-makepublic-member)|If provided, specifies whether to make the document public after check-in.|
||[saveChanges](/javascript/api/word/word.documentcheckinoptions#word-word-documentcheckinoptions-savechanges-member)|If provided, specifies whether to save changes before checking in.|
|[DocumentCheckInWithVersionOptions](/javascript/api/word/word.documentcheckinwithversionoptions)|[comment](/javascript/api/word/word.documentcheckinwithversionoptions#word-word-documentcheckinwithversionoptions-comment-member)|If provided, specifies a comment for the check-in operation.|
||[makePublic](/javascript/api/word/word.documentcheckinwithversionoptions#word-word-documentcheckinwithversionoptions-makepublic-member)|If provided, specifies whether to make the document public after check-in.|
||[saveChanges](/javascript/api/word/word.documentcheckinwithversionoptions#word-word-documentcheckinwithversionoptions-savechanges-member)|If provided, specifies whether to save changes before checking in.|
||[versionType](/javascript/api/word/word.documentcheckinwithversionoptions#word-word-documentcheckinwithversionoptions-versiontype-member)|If provided, specifies the version type for the check-in.|
|[DocumentCheckSpellingOptions](/javascript/api/word/word.documentcheckspellingoptions)|[alwaysSuggest](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-alwayssuggest-member)|If provided, specifies whether to always suggest spelling corrections.|
||[customDictionary10](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-customdictionary10-member)|If provided, specifies an additional custom dictionary to use for spell checking.|
||[customDictionary2](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-customdictionary2-member)|If provided, specifies an additional custom dictionary to use for spell checking.|
||[customDictionary3](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-customdictionary3-member)|If provided, specifies an additional custom dictionary to use for spell checking.|
||[customDictionary4](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-customdictionary4-member)|If provided, specifies an additional custom dictionary to use for spell checking.|
||[customDictionary5](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-customdictionary5-member)|If provided, specifies an additional custom dictionary to use for spell checking.|
||[customDictionary6](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-customdictionary6-member)|If provided, specifies an additional custom dictionary to use for spell checking.|
||[customDictionary7](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-customdictionary7-member)|If provided, specifies an additional custom dictionary to use for spell checking.|
||[customDictionary8](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-customdictionary8-member)|If provided, specifies an additional custom dictionary to use for spell checking.|
||[customDictionary9](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-customdictionary9-member)|If provided, specifies an additional custom dictionary to use for spell checking.|
||[customDictionary](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-customdictionary-member)|If provided, specifies the custom dictionary to use for spell checking.|
||[ignoreUppercase](/javascript/api/word/word.documentcheckspellingoptions#word-word-documentcheckspellingoptions-ignoreuppercase-member)|If provided, specifies whether to ignore uppercase words during spell checking.|
|[DocumentCountNumberedItemsOptions](/javascript/api/word/word.documentcountnumbereditemsoptions)|[level](/javascript/api/word/word.documentcountnumbereditemsoptions#word-word-documentcountnumbereditemsoptions-level-member)|If provided, specifies the level of numbering to count.|
||[numberType](/javascript/api/word/word.documentcountnumbereditemsoptions#word-word-documentcountnumbereditemsoptions-numbertype-member)|If provided, specifies the type of numbered items to count.|
|[DocumentEndReviewOptions](/javascript/api/word/word.documentendreviewoptions)|[includeAttachment](/javascript/api/word/word.documentendreviewoptions#word-word-documentendreviewoptions-includeattachment-member)|If provided, specifies whether to include the document as an attachment.|
||[recipients](/javascript/api/word/word.documentendreviewoptions#word-word-documentendreviewoptions-recipients-member)|If provided, specifies the recipients to notify when ending the review.|
||[showMessage](/javascript/api/word/word.documentendreviewoptions#word-word-documentendreviewoptions-showmessage-member)|If provided, specifies whether to show the message before sending.|
||[subject](/javascript/api/word/word.documentendreviewoptions#word-word-documentendreviewoptions-subject-member)|If provided, specifies the subject of the notification email.|
|[DocumentExportAsFixedFormat2Options](/javascript/api/word/word.documentexportasfixedformat2options)|[bitmapMissingFonts](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-bitmapmissingfonts-member)|If provided, specifies whether to bitmap missing fonts.|
||[createBookmarks](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-createbookmarks-member)|If provided, specifies the bookmark creation mode.|
||[documentStructureTags](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-documentstructuretags-member)|If provided, specifies whether to include document structure tags.|
||[fixedFormatExtClassPtr](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-fixedformatextclassptr-member)|If provided, specifies the extension class pointer.|
||[from](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-from-member)|If provided, specifies the starting page number.|
||[includeDocProps](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-includedocprops-member)|If provided, specifies whether to include document properties.|
||[item](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-item-member)|If provided, specifies the item to export.|
||[keepInformationRightsManagement](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-keepinformationrightsmanagement-member)|If provided, specifies whether to keep Information Rights Management (IRM) settings.|
||[openAfterExport](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-openafterexport-member)|If provided, specifies whether to open the file after export.|
||[optimizeFor](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-optimizefor-member)|If provided, specifies the optimization target for the export.|
||[optimizeForImageQuality](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-optimizeforimagequality-member)|If provided, specifies whether to optimize for image quality in the exported file.|
||[range](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-range-member)|If provided, specifies the range to export.|
||[to](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-to-member)|If provided, specifies the ending page number.|
||[useIso19005_1](/javascript/api/word/word.documentexportasfixedformat2options#word-word-documentexportasfixedformat2options-useiso19005_1-member)|If provided, specifies whether to use ISO 19005-1 compliance.|
|[DocumentExportAsFixedFormat3Options](/javascript/api/word/word.documentexportasfixedformat3options)|[bitmapMissingFonts](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-bitmapmissingfonts-member)|If provided, specifies whether to bitmap missing fonts.|
||[createBookmarks](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-createbookmarks-member)|If provided, specifies the bookmark creation mode.|
||[documentStructureTags](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-documentstructuretags-member)|If provided, specifies whether to include document structure tags.|
||[fixedFormatExtClassPtr](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-fixedformatextclassptr-member)|If provided, specifies the extension class pointer.|
||[from](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-from-member)|If provided, specifies the starting page number.|
||[improveExportTagging](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-improveexporttagging-member)|If provided, specifies to improve export tagging for better accessibility.|
||[includeDocProps](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-includedocprops-member)|If provided, specifies whether to include document properties.|
||[item](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-item-member)|If provided, specifies the item to export.|
||[keepInformationRightsManagement](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-keepinformationrightsmanagement-member)|If provided, specifies whether to keep Information Rights Management (IRM) settings.|
||[openAfterExport](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-openafterexport-member)|If provided, specifies whether to open the file after export.|
||[optimizeFor](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-optimizefor-member)|If provided, specifies the optimization target for the export.|
||[optimizeForImageQuality](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-optimizeforimagequality-member)|If provided, specifies whether to optimize for image quality in the exported file.|
||[range](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-range-member)|If provided, specifies the range to export.|
||[to](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-to-member)|If provided, specifies the ending page number.|
||[useIso19005_1](/javascript/api/word/word.documentexportasfixedformat3options#word-word-documentexportasfixedformat3options-useiso19005_1-member)|If provided, specifies whether to use ISO 19005-1 compliance.|
|[DocumentExportAsFixedFormatOptions](/javascript/api/word/word.documentexportasfixedformatoptions)|[bitmapMissingFonts](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-bitmapmissingfonts-member)|If provided, specifies whether to bitmap missing fonts.|
||[createBookmarks](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-createbookmarks-member)|If provided, specifies the bookmark creation mode.|
||[documentStructureTags](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-documentstructuretags-member)|If provided, specifies whether to include document structure tags.|
||[fixedFormatExtensionClassPointer](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-fixedformatextensionclasspointer-member)|If provided, specifies the extension class pointer.|
||[from](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-from-member)|If provided, specifies the starting page number.|
||[includeDocProps](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-includedocprops-member)|If provided, specifies whether to include document properties.|
||[item](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-item-member)|If provided, specifies the item to export.|
||[keepInformationRightsManagement](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-keepinformationrightsmanagement-member)|If provided, specifies whether to keep Information Rights Management (IRM) settings.|
||[openAfterExport](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-openafterexport-member)|If provided, specifies whether to open the file after export.|
||[optimizeFor](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-optimizefor-member)|If provided, specifies the optimization target for the export.|
||[range](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-range-member)|If provided, specifies the range to export.|
||[to](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-to-member)|If provided, specifies the ending page number.|
||[useIso19005_1](/javascript/api/word/word.documentexportasfixedformatoptions#word-word-documentexportasfixedformatoptions-useiso19005_1-member)|If provided, specifies whether to use ISO 19005-1 compliance.|
|[DocumentFollowHyperlinkOptions](/javascript/api/word/word.documentfollowhyperlinkoptions)|[addHistory](/javascript/api/word/word.documentfollowhyperlinkoptions#word-word-documentfollowhyperlinkoptions-addhistory-member)|If provided, specifies whether to add the link to the browsing history.|
||[address](/javascript/api/word/word.documentfollowhyperlinkoptions#word-word-documentfollowhyperlinkoptions-address-member)|If provided, specifies the hyperlink address to follow.|
||[extraInfo](/javascript/api/word/word.documentfollowhyperlinkoptions#word-word-documentfollowhyperlinkoptions-extrainfo-member)|If provided, specifies additional information to pass with the request.|
||[headerInfo](/javascript/api/word/word.documentfollowhyperlinkoptions#word-word-documentfollowhyperlinkoptions-headerinfo-member)|If provided, specifies header information for the HTTP request.|
||[httpMethod](/javascript/api/word/word.documentfollowhyperlinkoptions#word-word-documentfollowhyperlinkoptions-httpmethod-member)|If provided, specifies the HTTP method to use for the request.|
||[newWindow](/javascript/api/word/word.documentfollowhyperlinkoptions#word-word-documentfollowhyperlinkoptions-newwindow-member)|If provided, specifies whether to open the link in a new window.|
||[subAddress](/javascript/api/word/word.documentfollowhyperlinkoptions#word-word-documentfollowhyperlinkoptions-subaddress-member)|If provided, specifies the sub-address within the document.|
|[DocumentMergeOptions](/javascript/api/word/word.documentmergeoptions)|[addToRecentFiles](/javascript/api/word/word.documentmergeoptions#word-word-documentmergeoptions-addtorecentfiles-member)|If provided, specifies whether to add the merged document to recent files.|
||[detectFormatChanges](/javascript/api/word/word.documentmergeoptions#word-word-documentmergeoptions-detectformatchanges-member)|If provided, specifies whether to detect format changes during the merge.|
||[mergeTarget](/javascript/api/word/word.documentmergeoptions#word-word-documentmergeoptions-mergetarget-member)|If provided, specifies the target of the merge operation.|
||[useFormattingFrom](/javascript/api/word/word.documentmergeoptions#word-word-documentmergeoptions-useformattingfrom-member)|If provided, specifies the source of formatting to use in the merge.|
|[DocumentPrintOutOptions](/javascript/api/word/word.documentprintoutoptions)|[activePrinterMacGX](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-activeprintermacgx-member)|If provided, specifies the printer name.|
||[append](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-append-member)|If provided, specifies whether to append to an existing file.|
||[background](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-background-member)|If provided, specifies whether to print in the background.|
||[collate](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-collate-member)|If provided, specifies whether to collate pages.|
||[copies](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-copies-member)|If provided, specifies the number of copies to print.|
||[from](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-from-member)|If provided, specifies the starting page number.|
||[item](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-item-member)|If provided, specifies the item to print.|
||[manualDuplexPrint](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-manualduplexprint-member)|If provided, specifies whether to manually duplex print.|
||[outputFileName](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-outputfilename-member)|If provided, specifies the name of the output file.|
||[pageType](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-pagetype-member)|If provided, specifies the page order.|
||[pages](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-pages-member)|If provided, specifies specific pages to print.|
||[printToFile](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-printtofile-member)|If provided, specifies whether to print to file.|
||[printZoomColumn](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-printzoomcolumn-member)|If provided, specifies the zoom column setting.|
||[printZoomPaperHeight](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-printzoompaperheight-member)|If provided, specifies the paper height for printing in twips (20 twips = 1 point; 72 points = 1 inch).|
||[printZoomPaperWidth](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-printzoompaperwidth-member)|If provided, specifies the paper width for printing in twips (20 twips = 1 point; 72 points = 1 inch).|
||[printZoomRow](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-printzoomrow-member)|If provided, specifies the zoom row setting.|
||[range](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-range-member)|If provided, specifies the range to print.|
||[to](/javascript/api/word/word.documentprintoutoptions#word-word-documentprintoutoptions-to-member)|If provided, specifies the ending page number.|
|[DocumentProtectOptions](/javascript/api/word/word.documentprotectoptions)|[enforceStyleLock](/javascript/api/word/word.documentprotectoptions#word-word-documentprotectoptions-enforcestylelock-member)|If provided, specifies whether to enforce style lock restrictions.|
||[noReset](/javascript/api/word/word.documentprotectoptions#word-word-documentprotectoptions-noreset-member)|If provided, specifies whether to reset form fields when protecting the document.|
||[password](/javascript/api/word/word.documentprotectoptions#word-word-documentprotectoptions-password-member)|If provided, specifies the password to apply for document protection.|
||[useInformationRightsManagement](/javascript/api/word/word.documentprotectoptions#word-word-documentprotectoptions-useinformationrightsmanagement-member)|If provided, specifies whether to use Information Rights Management (IRM).|
|[DocumentRangeOptions](/javascript/api/word/word.documentrangeoptions)|[end](/javascript/api/word/word.documentrangeoptions#word-word-documentrangeoptions-end-member)|If provided, specifies the ending character position.|
||[start](/javascript/api/word/word.documentrangeoptions#word-word-documentrangeoptions-start-member)|If provided, specifies the starting character position.|
|[DocumentReplyWithChangesOptions](/javascript/api/word/word.documentreplywithchangesoptions)|[includeAttachment](/javascript/api/word/word.documentreplywithchangesoptions#word-word-documentreplywithchangesoptions-includeattachment-member)|If provided, specifies whether to include the document as an attachment.|
||[recipients](/javascript/api/word/word.documentreplywithchangesoptions#word-word-documentreplywithchangesoptions-recipients-member)|If provided, specifies the recipients of the reply.|
||[showMessage](/javascript/api/word/word.documentreplywithchangesoptions#word-word-documentreplywithchangesoptions-showmessage-member)|If provided, specifies whether to show the message before sending.|
||[subject](/javascript/api/word/word.documentreplywithchangesoptions#word-word-documentreplywithchangesoptions-subject-member)|If provided, specifies the subject of the reply email.|
|[DocumentSendFaxOverInternetOptions](/javascript/api/word/word.documentsendfaxoverinternetoptions)|[recipients](/javascript/api/word/word.documentsendfaxoverinternetoptions#word-word-documentsendfaxoverinternetoptions-recipients-member)|If provided, specifies the recipients of the fax.|
||[showMessage](/javascript/api/word/word.documentsendfaxoverinternetoptions#word-word-documentsendfaxoverinternetoptions-showmessage-member)|If provided, specifies whether to show the message before sending.|
||[subject](/javascript/api/word/word.documentsendfaxoverinternetoptions#word-word-documentsendfaxoverinternetoptions-subject-member)|If provided, specifies the subject of the fax.|
|[DocumentSendForReviewOptions](/javascript/api/word/word.documentsendforreviewoptions)|[includeAttachment](/javascript/api/word/word.documentsendforreviewoptions#word-word-documentsendforreviewoptions-includeattachment-member)|If provided, specifies whether to include the document as an attachment.|
||[recipients](/javascript/api/word/word.documentsendforreviewoptions#word-word-documentsendforreviewoptions-recipients-member)|If provided, specifies the recipients of the review request.|
||[showMessage](/javascript/api/word/word.documentsendforreviewoptions#word-word-documentsendforreviewoptions-showmessage-member)|If provided, specifies whether to show the message before sending.|
||[subject](/javascript/api/word/word.documentsendforreviewoptions#word-word-documentsendforreviewoptions-subject-member)|If provided, specifies the subject of the review email.|
|[DropCap](/javascript/api/word/word.dropcap)|[clear()](/javascript/api/word/word.dropcap#word-word-dropcap-clear-member(1))|Removes the dropped capital letter formatting.|
||[distanceFromText](/javascript/api/word/word.dropcap#word-word-dropcap-distancefromtext-member)|Gets the distance (in points) between the dropped capital letter and the paragraph text.|
||[enable()](/javascript/api/word/word.dropcap#word-word-dropcap-enable-member(1))|Formats the first character in the specified paragraph as a dropped capital letter.|
||[fontName](/javascript/api/word/word.dropcap#word-word-dropcap-fontname-member)|Gets the name of the font for the dropped capital letter.|
||[linesToDrop](/javascript/api/word/word.dropcap#word-word-dropcap-linestodrop-member)|Gets the height (in lines) of the dropped capital letter.|
||[position](/javascript/api/word/word.dropcap#word-word-dropcap-position-member)|Gets the position of the dropped capital letter.|
|[Editor](/javascript/api/word/word.editor)|[delete()](/javascript/api/word/word.editor#word-word-editor-delete-member(1))|Deletes the `Editor` object.|
||[id](/javascript/api/word/word.editor#word-word-editor-id-member)|Gets the identifier for the `Editor` object when the parent document is saved as a webpage.|
||[name](/javascript/api/word/word.editor#word-word-editor-name-member)|Gets the name of the editor.|
||[nextRange](/javascript/api/word/word.editor#word-word-editor-nextrange-member)|Gets a `Range` object that represents the next range that the editor has permissions to modify.|
||[range](/javascript/api/word/word.editor#word-word-editor-range-member)|Gets a `Range` object that represents the portion of the document that's contained in the `Editor` object.|
||[removeAllPermissions()](/javascript/api/word/word.editor#word-word-editor-removeallpermissions-member(1))|Removes all editing permissions in the document for the editor.|
||[selectAllShapes()](/javascript/api/word/word.editor#word-word-editor-selectallshapes-member(1))|Selects all the shapes in the document that were inserted or edited by the editor.|
|[EditorCollection](/javascript/api/word/word.editorcollection)|[addById(editorId: string)](/javascript/api/word/word.editorcollection#word-word-editorcollection-addbyid-member(1))|Returns an `Editor` object that represents a new permission for the specified user to modify a range within the document.|
||[addByType(editorType: Word.EditorType)](/javascript/api/word/word.editorcollection#word-word-editorcollection-addbytype-member(1))|Returns an `Editor` object that represents a new permission for the specified group of users to modify a range within the document.|
||[getCount()](/javascript/api/word/word.editorcollection#word-word-editorcollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/javascript/api/word/word.editorcollection#word-word-editorcollection-getitemat-member(1))|Gets an `Editor` object by its index in the collection.|
|[Field](/javascript/api/word/word.field)|[copyToClipboard()](/javascript/api/word/word.field#word-word-field-copytoclipboard-member(1))|Copies the field to the Clipboard.|
||[cut()](/javascript/api/word/word.field#word-word-field-cut-member(1))|Removes the field from the document and places it on the Clipboard.|
||[doClick()](/javascript/api/word/word.field#word-word-field-doclick-member(1))|Clicks the field.|
||[linkFormat](/javascript/api/word/word.field#word-word-field-linkformat-member)|Gets a `LinkFormat` object that represents the link options of the field.|
||[oleFormat](/javascript/api/word/word.field#word-word-field-oleformat-member)|Gets an `OleFormat` object that represents the OLE characteristics (other than linking) for the field.|
||[unlink()](/javascript/api/word/word.field#word-word-field-unlink-member(1))|Replaces the field with its most recent result.|
||[updateSource()](/javascript/api/word/word.field#word-word-field-updatesource-member(1))|Saves the changes made to the results of an INCLUDETEXT field back to the source document.|
|[FontNameCollection](/javascript/api/word/word.fontnamecollection)|[getCount()](/javascript/api/word/word.fontnamecollection#word-word-fontnamecollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/javascript/api/word/word.fontnamecollection#word-word-fontnamecollection-getitemat-member(1))|Gets the font name at the specified index.|
|[GoToOptions](/javascript/api/word/word.gotooptions)|[count](/javascript/api/word/word.gotooptions#word-word-gotooptions-count-member)|If provided, specifies the number of the item in the document.|
||[direction](/javascript/api/word/word.gotooptions#word-word-gotooptions-direction-member)|If provided, specifies the direction the range or selection is moved to.|
||[item](/javascript/api/word/word.gotooptions#word-word-gotooptions-item-member)|If provided, specifies the kind of item the range or selection is moved to.|
||[name](/javascript/api/word/word.gotooptions#word-word-gotooptions-name-member)|If provided, specifies the name if the `item` property is set to Word.GoToItem type `bookmark`, `comment`, `field`, or `object`.|
|[HeadingStyle](/javascript/api/word/word.headingstyle)|[delete()](/javascript/api/word/word.headingstyle#word-word-headingstyle-delete-member(1))|Deletes the heading style.|
||[level](/javascript/api/word/word.headingstyle#word-word-headingstyle-level-member)|Specifies the level for the heading style in a table of contents or table of figures.|
||[name](/javascript/api/word/word.headingstyle#word-word-headingstyle-name-member)|Specifies the name of style for a heading.|
|[HeadingStyleCollection](/javascript/api/word/word.headingstylecollection)|[add(name: string, level: number)](/javascript/api/word/word.headingstylecollection#word-word-headingstylecollection-add-member(1))|Adds a new heading style to a document.|
||[items](/javascript/api/word/word.headingstylecollection#word-word-headingstylecollection-items-member)|Gets the loaded child items in this collection.|
|[HtmlDivision](/javascript/api/word/word.htmldivision)|[delete()](/javascript/api/word/word.htmldivision#word-word-htmldivision-delete-member(1))|Deletes this HTML division.|
||[htmlDivisionParent(levelsUp?: number)](/javascript/api/word/word.htmldivision#word-word-htmldivision-htmldivisionparent-member(1))|Returns an `HtmlDivision` object that represents a parent division of the current HTML division.|
||[htmlDivisions](/javascript/api/word/word.htmldivision#word-word-htmldivision-htmldivisions-member)||
||[leftIndent](/javascript/api/word/word.htmldivision#word-word-htmldivision-leftindent-member)|Specifies the left indent value (in points) for this HTML division.|
||[range](/javascript/api/word/word.htmldivision#word-word-htmldivision-range-member)|Gets a `Range` object that represents the portion of a document that's contained in this HTML division.|
||[rightIndent](/javascript/api/word/word.htmldivision#word-word-htmldivision-rightindent-member)|Specifies the right indent (in points) for this HTML division.|
||[spaceAfter](/javascript/api/word/word.htmldivision#word-word-htmldivision-spaceafter-member)|Specifies the amount of spacing (in points) after this HTML division.|
||[spaceBefore](/javascript/api/word/word.htmldivision#word-word-htmldivision-spacebefore-member)|Specifies the spacing (in points) before this HTML division.|
|[HtmlDivisionCollection](/javascript/api/word/word.htmldivisioncollection)|[getItemAt(index: number)](/javascript/api/word/word.htmldivisioncollection#word-word-htmldivisioncollection-getitemat-member(1))|Returns an `HtmlDivision` object from the collection based on the specified index.|
||[items](/javascript/api/word/word.htmldivisioncollection#word-word-htmldivisioncollection-items-member)|Gets the loaded child items in this collection.|
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
||[markAllEntries(range: Word.Range, markAllEntriesOptions?: Word.IndexMarkAllEntriesOptions)](/javascript/api/word/word.indexcollection#word-word-indexcollection-markallentries-member(1))|Inserts an XE (Index Entry) field after all instances of the text in the range.|
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
|[LinkFormat](/javascript/api/word/word.linkformat)|[breakLink()](/javascript/api/word/word.linkformat#word-word-linkformat-breaklink-member(1))|Breaks the link between the source file and the OLE object, picture, or linked field.|
||[isAutoUpdated](/javascript/api/word/word.linkformat#word-word-linkformat-isautoupdated-member)|Specifies if the link is updated automatically when the container file is opened or when the source file is changed.|
||[isLocked](/javascript/api/word/word.linkformat#word-word-linkformat-islocked-member)|Specifies if a `Field`, `InlineShape`, or `Shape` object is locked to prevent automatic updating.|
||[isPictureSavedWithDocument](/javascript/api/word/word.linkformat#word-word-linkformat-ispicturesavedwithdocument-member)|Specifies if the linked picture is saved with the document.|
||[sourceFullName](/javascript/api/word/word.linkformat#word-word-linkformat-sourcefullname-member)|Specifies the path and name of the source file for the linked OLE object, picture, or field.|
||[sourceName](/javascript/api/word/word.linkformat#word-word-linkformat-sourcename-member)|Gets the name of the source file for the linked OLE object, picture, or field.|
||[sourcePath](/javascript/api/word/word.linkformat#word-word-linkformat-sourcepath-member)|Gets the path of the source file for the linked OLE object, picture, or field.|
||[type](/javascript/api/word/word.linkformat#word-word-linkformat-type-member)|Gets the link type.|
|[ListTemplate](/javascript/api/word/word.listtemplate)|[name](/javascript/api/word/word.listtemplate#word-word-listtemplate-name-member)|Specifies the name of the list template.|
|[ListTemplateCollection](/javascript/api/word/word.listtemplatecollection)|[add(options?: Word.ListTemplateCollectionAddOptions)](/javascript/api/word/word.listtemplatecollection#word-word-listtemplatecollection-add-member(1))|Adds a new `ListTemplate` object.|
||[getItem(index: number)](/javascript/api/word/word.listtemplatecollection#word-word-listtemplatecollection-getitem-member(1))|Gets a `ListTemplate` object by its index in the collection.|
||[items](/javascript/api/word/word.listtemplatecollection#word-word-listtemplatecollection-items-member)|Gets the loaded child items in this collection.|
|[ListTemplateCollectionAddOptions](/javascript/api/word/word.listtemplatecollectionaddoptions)|[name](/javascript/api/word/word.listtemplatecollectionaddoptions#word-word-listtemplatecollectionaddoptions-name-member)|If provided, specifies the name of the list template to be added.|
||[outlineNumbered](/javascript/api/word/word.listtemplatecollectionaddoptions#word-word-listtemplatecollectionaddoptions-outlinenumbered-member)|If provided, specifies whether to apply outline numbering to the new list template.|
|[ListTemplateGallery](/javascript/api/word/word.listtemplategallery)|[listTemplates](/javascript/api/word/word.listtemplategallery#word-word-listtemplategallery-listtemplates-member)|Returns a `ListTemplateCollection` object that represents all the list templates for the specified list gallery.|
|[ListTemplateGalleryCollection](/javascript/api/word/word.listtemplategallerycollection)|[getByType(type: Word.ListTemplateGalleryType)](/javascript/api/word/word.listtemplategallerycollection#word-word-listtemplategallerycollection-getbytype-member(1))|Gets a `ListTemplateGallery` object by its type in the collection.|
||[getItem(index: number)](/javascript/api/word/word.listtemplategallerycollection#word-word-listtemplategallerycollection-getitem-member(1))|Gets a `ListTemplateGallery` object by its index in the collection.|
||[items](/javascript/api/word/word.listtemplategallerycollection#word-word-listtemplategallerycollection-items-member)|Gets the loaded child items in this collection.|
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
||[progId](/javascript/api/word/word.oleformat#word-word-oleformat-progid-member)|Gets the programmatic identifier (`ProgId`) for the specified OLE object.|
|[Pane](/javascript/api/word/word.pane)|[selection](/javascript/api/word/word.pane#word-word-pane-selection-member)|Returns a `Selection` object that represents the current selection in the pane.|
|[PaneCollection](/javascript/api/word/word.panecollection)||Represents the collection of Word.Pane objects.|
|[Paragraph](/javascript/api/word/word.paragraph)|[closeUp()](/javascript/api/word/word.paragraph#word-word-paragraph-closeup-member(1))|Removes any spacing before the paragraph.|
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
||[reset()](/javascript/api/word/word.paragraph#word-word-paragraph-reset-member(1))|Removes manual paragraph formatting (formatting not applied using a style).|
||[resetAdvanceTo()](/javascript/api/word/word.paragraph#word-word-paragraph-resetadvanceto-member(1))|Resets the paragraph that uses custom list levels to the original level settings.|
||[selectNumber()](/javascript/api/word/word.paragraph#word-word-paragraph-selectnumber-member(1))|Selects the number or bullet in a list.|
||[separateList()](/javascript/api/word/word.paragraph#word-word-paragraph-separatelist-member(1))|Separates a list into two separate lists.|
||[space1()](/javascript/api/word/word.paragraph#word-word-paragraph-space1-member(1))|Sets the paragraph to single spacing.|
||[space1Pt5()](/javascript/api/word/word.paragraph#word-word-paragraph-space1pt5-member(1))|Sets the paragraph to 1.5-line spacing.|
||[space2()](/javascript/api/word/word.paragraph#word-word-paragraph-space2-member(1))|Sets the paragraph to double spacing.|
||[tabHangingIndent(count: number)](/javascript/api/word/word.paragraph#word-word-paragraph-tabhangingindent-member(1))|Sets a hanging indent to a specified number of tab stops.|
||[tabIndent(count: number)](/javascript/api/word/word.paragraph#word-word-paragraph-tabindent-member(1))|Sets the left indent for the paragraph to a specified number of tab stops.|
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
|[Range](/javascript/api/word/word.range)|[bold](/javascript/api/word/word.range#word-word-range-bold-member)|Specifies whether the range is formatted as bold.|
||[boldBidirectional](/javascript/api/word/word.range#word-word-range-boldbidirectional-member)|Specifies whether the range is formatted as bold in a right-to-left language document.|
||[bookmarks](/javascript/api/word/word.range#word-word-range-bookmarks-member)|Returns a `BookmarkCollection` object that represents all the bookmarks in the range.|
||[borders](/javascript/api/word/word.range#word-word-range-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the range.|
||[case](/javascript/api/word/word.range#word-word-range-case-member)|Specifies a `CharacterCase` value that represents the case of the text in the range.|
||[characterWidth](/javascript/api/word/word.range#word-word-range-characterwidth-member)|Specifies the character width of the range.|
||[combineCharacters](/javascript/api/word/word.range#word-word-range-combinecharacters-member)|Specifies if the range contains combined characters.|
||[conflicts](/javascript/api/word/word.range#word-word-range-conflicts-member)|Returns a `ConflictCollection` object that contains all the Word.Conflict objects in the range.|
||[disableCharacterSpaceGrid](/javascript/api/word/word.range#word-word-range-disablecharacterspacegrid-member)|Specifies if Microsoft Word ignores the number of characters per line for the corresponding `Range` object.|
||[editors](/javascript/api/word/word.range#word-word-range-editors-member)|Returns an `EditorCollection` object that represents all the users authorized to modify the range when the document is in protected (read-only) mode.|
||[emphasisMark](/javascript/api/word/word.range#word-word-range-emphasismark-member)|Specifies the emphasis mark for a character or designated character string.|
||[end](/javascript/api/word/word.range#word-word-range-end-member)|Specifies the ending character position of the range.|
||[fitTextWidth](/javascript/api/word/word.range#word-word-range-fittextwidth-member)|Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range.|
||[grammarChecked](/javascript/api/word/word.range#word-word-range-grammarchecked-member)|Specifies if a grammar check has been run on the range or document.|
||[highlightColorIndex](/javascript/api/word/word.range#word-word-range-highlightcolorindex-member)|Specifies the highlight color for the range.|
||[horizontalInVertical](/javascript/api/word/word.range#word-word-range-horizontalinvertical-member)|Specifies the formatting for horizontal text set within vertical text.|
||[id](/javascript/api/word/word.range#word-word-range-id-member)|Specifies the ID for the range.|
||[isEndOfRowMark](/javascript/api/word/word.range#word-word-range-isendofrowmark-member)|Gets if the range is collapsed and is located at the end-of-row mark in a table.|
||[isTextVisibleOnScreen](/javascript/api/word/word.range#word-word-range-istextvisibleonscreen-member)|Gets whether the text in the range is visible on the screen.|
||[italic](/javascript/api/word/word.range#word-word-range-italic-member)|Specifies if the font or range is formatted as italic.|
||[italicBidirectional](/javascript/api/word/word.range#word-word-range-italicbidirectional-member)|Specifies if the font or range is formatted as italic (right-to-left languages).|
||[kana](/javascript/api/word/word.range#word-word-range-kana-member)|Specifies whether the range of Japanese language text is hiragana or katakana.|
||[onCommentAdded](/javascript/api/word/word.range#word-word-range-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.range#word-word-range-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.range#word-word-range-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.range#word-word-range-oncommentselected-member)|Occurs when a comment is selected.|
||[revisions](/javascript/api/word/word.range#word-word-range-revisions-member)|Gets the collection of revisions that represents the tracked changes in the range.|
||[showAll](/javascript/api/word/word.range#word-word-range-showall-member)|Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed.|
||[spellingChecked](/javascript/api/word/word.range#word-word-range-spellingchecked-member)|Specifies if spelling has been checked throughout the range or document.|
||[start](/javascript/api/word/word.range#word-word-range-start-member)|Specifies the starting character position of the range.|
||[storyLength](/javascript/api/word/word.range#word-word-range-storylength-member)|Gets the number of characters in the story that contains the range.|
||[storyType](/javascript/api/word/word.range#word-word-range-storytype-member)|Gets the story type for the range.|
||[twoLinesInOne](/javascript/api/word/word.range#word-word-range-twolinesinone-member)|Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.|
||[underline](/javascript/api/word/word.range#word-word-range-underline-member)|Specifies the type of underline applied to the range.|
|[RangeScopedCollection](/javascript/api/word/word.rangescopedcollection)|[getItem(index: number)](/javascript/api/word/word.rangescopedcollection#word-word-rangescopedcollection-getitem-member(1))|Gets a `Range` object by its index in the collection.|
||[items](/javascript/api/word/word.rangescopedcollection#word-word-rangescopedcollection-items-member)|Gets the loaded child items in this collection.|
|[ReadabilityStatistic](/javascript/api/word/word.readabilitystatistic)|[name](/javascript/api/word/word.readabilitystatistic#word-word-readabilitystatistic-name-member)|Returns the name of the readability statistic.|
||[value](/javascript/api/word/word.readabilitystatistic#word-word-readabilitystatistic-value-member)|Returns the value of the grammar statistic.|
|[ReadabilityStatisticCollection](/javascript/api/word/word.readabilitystatisticcollection)|[getItemAt(index: number)](/javascript/api/word/word.readabilitystatisticcollection#word-word-readabilitystatisticcollection-getitemat-member(1))|Gets the readability statistic at the specified index.|
||[items](/javascript/api/word/word.readabilitystatisticcollection#word-word-readabilitystatisticcollection-items-member)|Gets the loaded child items in this collection.|
|[RepeatingSectionContentControl](/javascript/api/word/word.repeatingsectioncontentcontrol)|[xmlapping](/javascript/api/word/word.repeatingsectioncontentcontrol#word-word-repeatingsectioncontentcontrol-xmlapping-member)|Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[Reviewer](/javascript/api/word/word.reviewer)|[isVisible](/javascript/api/word/word.reviewer#word-word-reviewer-isvisible-member)|Specifies if the `Reviewer` object is visible.|
|[ReviewerCollection](/javascript/api/word/word.reviewercollection)|[getItem(index: number)](/javascript/api/word/word.reviewercollection#word-word-reviewercollection-getitem-member(1))|Returns a `Reviewer` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.reviewercollection#word-word-reviewercollection-items-member)|Gets the loaded child items in this collection.|
|[Revision](/javascript/api/word/word.revision)|[accept()](/javascript/api/word/word.revision#word-word-revision-accept-member(1))|Accepts the tracked change, removes the revision mark, and incorporates the change into the document.|
||[author](/javascript/api/word/word.revision#word-word-revision-author-member)|Gets the name of the user who made the tracked change.|
||[date](/javascript/api/word/word.revision#word-word-revision-date-member)|Gets the date and time when the tracked change was made.|
||[formatDescription](/javascript/api/word/word.revision#word-word-revision-formatdescription-member)|Gets the description of tracked formatting changes in the revision.|
||[index](/javascript/api/word/word.revision#word-word-revision-index-member)|Gets a number that represents the position of this item in a collection.|
||[movedRange](/javascript/api/word/word.revision#word-word-revision-movedrange-member)|Gets a `Range` object that represents the range of text that was moved from one place to another in the document with tracked changes.|
||[range](/javascript/api/word/word.revision#word-word-revision-range-member)|Gets a `Range` object that represents the portion of the document that's contained within a revision mark.|
||[reject()](/javascript/api/word/word.revision#word-word-revision-reject-member(1))|Rejects the tracked change.|
||[type](/javascript/api/word/word.revision#word-word-revision-type-member)|Gets the revision type.|
|[RevisionCollection](/javascript/api/word/word.revisioncollection)|[acceptAll()](/javascript/api/word/word.revisioncollection#word-word-revisioncollection-acceptall-member(1))|Accepts all the tracked changes in the document or range, removes all revision marks, and incorporates the changes into the document.|
||[getItem(index: number)](/javascript/api/word/word.revisioncollection#word-word-revisioncollection-getitem-member(1))|Returns a `Revision` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.revisioncollection#word-word-revisioncollection-items-member)|Gets the loaded child items in this collection.|
||[rejectAll()](/javascript/api/word/word.revisioncollection#word-word-revisioncollection-rejectall-member(1))|Rejects all the tracked changes in the document or range.|
|[RevisionsFilter](/javascript/api/word/word.revisionsfilter)|[markup](/javascript/api/word/word.revisionsfilter#word-word-revisionsfilter-markup-member)|Specifies a `RevisionsMarkup` value that represents the extent of reviewer markup displayed in the document.|
||[reviewers](/javascript/api/word/word.revisionsfilter#word-word-revisionsfilter-reviewers-member)|Gets the `ReviewerCollection` object that represents the collection of reviewers of one or more documents.|
||[toggleShowAllReviewers()](/javascript/api/word/word.revisionsfilter#word-word-revisionsfilter-toggleshowallreviewers-member(1))|Shows or hides all revisions in the document that contain comments and tracked changes.|
||[view](/javascript/api/word/word.revisionsfilter#word-word-revisionsfilter-view-member)|Specifies a `RevisionsView` value that represents globally whether Word displays the original version of the document or the final version, which might have revisions and formatting changes applied.|
|[SelectNodesOptions](/javascript/api/word/word.selectnodesoptions)|[fastSearchSkippingTextNodes](/javascript/api/word/word.selectnodesoptions#word-word-selectnodesoptions-fastsearchskippingtextnodes-member)|If provided, specifies whether to skip text nodes in the search.|
||[prefixMapping](/javascript/api/word/word.selectnodesoptions#word-word-selectnodesoptions-prefixmapping-member)|If provided, specifies the prefix mapping for the XPath expression.|
|[SelectSingleNodeOptions](/javascript/api/word/word.selectsinglenodeoptions)|[fastSearchSkippingTextNodes](/javascript/api/word/word.selectsinglenodeoptions#word-word-selectsinglenodeoptions-fastsearchskippingtextnodes-member)|If provided, specifies whether to skip text nodes in the search.|
||[prefixMapping](/javascript/api/word/word.selectsinglenodeoptions#word-word-selectsinglenodeoptions-prefixmapping-member)|If provided, specifies the prefix mapping for the XPath expression.|
|[Selection](/javascript/api/word/word.selection)|[borders](/javascript/api/word/word.selection#word-word-selection-borders-member)|Returns a `BorderUniversalCollection` object that represents all the borders for the objects in the selection.|
||[calculate()](/javascript/api/word/word.selection#word-word-selection-calculate-member(1))|Calculates the first mathematical expression within the selection.|
||[cancelMode()](/javascript/api/word/word.selection#word-word-selection-cancelmode-member(1))|Cancels a mode such as extend or column select.|
||[characters](/javascript/api/word/word.selection#word-word-selection-characters-member)|Returns a `RangeScopedCollection` object that represents each character in the selection.|
||[clearCharacterStyleFormatting()](/javascript/api/word/word.selection#word-word-selection-clearcharacterstyleformatting-member(1))|Removes character formatting applied through character styles.|
||[clearFormatting()](/javascript/api/word/word.selection#word-word-selection-clearformatting-member(1))|Removes character and paragraph formatting from the selection.|
||[clearManualCharacterFormatting()](/javascript/api/word/word.selection#word-word-selection-clearmanualcharacterformatting-member(1))|Removes manually applied character formatting from the selected text.|
||[clearManualParagraphFormatting()](/javascript/api/word/word.selection#word-word-selection-clearmanualparagraphformatting-member(1))|Removes manually applied paragraph formatting from the selected text.|
||[clearParagraphFormatting()](/javascript/api/word/word.selection#word-word-selection-clearparagraphformatting-member(1))|Removes all paragraph formatting from the selected text.|
||[clearParagraphStyle()](/javascript/api/word/word.selection#word-word-selection-clearparagraphstyle-member(1))|Removes paragraph formatting applied through paragraph styles.|
||[collapse(direction?: Word.CollapseDirection)](/javascript/api/word/word.selection#word-word-selection-collapse-member(1))|Collapses the selection to the starting or ending position.|
||[comments](/javascript/api/word/word.selection#word-word-selection-comments-member)|Returns a `CommentCollection` object that represents all the comments in the selection.|
||[convertToTable(options?: Word.SelectionConvertToTableOptions)](/javascript/api/word/word.selection#word-word-selection-converttotable-member(1))|Converts text within a range to a table.|
||[copyAsPictureToClipboard()](/javascript/api/word/word.selection#word-word-selection-copyaspicturetoclipboard-member(1))|Copies the selection to the Clipboard as a picture.|
||[copyFormat()](/javascript/api/word/word.selection#word-word-selection-copyformat-member(1))|Copies the character formatting of the first character in the selected text.|
||[copyToClipboard()](/javascript/api/word/word.selection#word-word-selection-copytoclipboard-member(1))|Copies the selection to the Clipboard.|
||[createTextBox()](/javascript/api/word/word.selection#word-word-selection-createtextbox-member(1))|Adds a default-sized text box around the selection.|
||[cut()](/javascript/api/word/word.selection#word-word-selection-cut-member(1))|Removes the selected content from the document and moves it to the Clipboard.|
||[delete(options?: Word.SelectionDeleteOptions)](/javascript/api/word/word.selection#word-word-selection-delete-member(1))|Deletes the specified number of characters or words.|
||[detectLanguage()](/javascript/api/word/word.selection#word-word-selection-detectlanguage-member(1))|Analyzes the selected text to determine the language that it's written in.|
||[end](/javascript/api/word/word.selection#word-word-selection-end-member)|Specifies the ending character position of the selection.|
||[expand(unit?: Word.OperationUnit)](/javascript/api/word/word.selection#word-word-selection-expand-member(1))|Expands the selection.|
||[expandToWholeStory()](/javascript/api/word/word.selection#word-word-selection-expandtowholestory-member(1))|Expands the selection to include the entire story.|
||[extend(character?: Word.OperationUnit)](/javascript/api/word/word.selection#word-word-selection-extend-member(1))|Turns on extend mode, or if extend mode is already on, extends the selection to the next larger unit of text.|
||[fields](/javascript/api/word/word.selection#word-word-selection-fields-member)|Returns a `FieldCollection` object that represents all the fields in the selection.|
||[fitTextWidth](/javascript/api/word/word.selection#word-word-selection-fittextwidth-member)|Specifies the width in which Word fits the text in the current selection.|
||[font](/javascript/api/word/word.selection#word-word-selection-font-member)|Returns the `Font` object that represents the character formatting of the selection.|
||[formattedText](/javascript/api/word/word.selection#word-word-selection-formattedtext-member)|Specifies a `Range` object that includes the formatted text in the range or selection.|
||[getNextRange(options?: Word.SelectionNextOptions)](/javascript/api/word/word.selection#word-word-selection-getnextrange-member(1))|Returns a `Range` object that represents the next unit relative to the selection.|
||[getPreviousRange(options?: Word.SelectionPreviousOptions)](/javascript/api/word/word.selection#word-word-selection-getpreviousrange-member(1))|Returns a `Range` object that represents the previous unit relative to the selection.|
||[goTo(options?: Word.SelectionGoToOptions)](/javascript/api/word/word.selection#word-word-selection-goto-member(1))|Returns a `Range` object that represents the area specified by the `options` and moves the insertion point to the character position immediately preceding the specified item.|
||[goToNext(what: Word.GoToItem)](/javascript/api/word/word.selection#word-word-selection-gotonext-member(1))|Returns a `Range` object that refers to the start position of the next item or location specified by the `what` argument and moves the selection to the specified item.|
||[goToPrevious(what: Word.GoToItem)](/javascript/api/word/word.selection#word-word-selection-gotoprevious-member(1))|Returns a `Range` object that refers to the start position of the previous item or location specified by the `what` argument and moves the selection to the specified item.|
||[hasNoProofing](/javascript/api/word/word.selection#word-word-selection-hasnoproofing-member)|Returns whether the spelling and grammar checker ignores the selected text.|
||[insertAfter(text: string)](/javascript/api/word/word.selection#word-word-selection-insertafter-member(1))|Inserts the specified text at the end of the range or selection.|
||[insertBefore(text: string)](/javascript/api/word/word.selection#word-word-selection-insertbefore-member(1))|Inserts the specified text before the selection.|
||[insertCells(shiftCells?: Word.TableCellInsertionLocation)](/javascript/api/word/word.selection#word-word-selection-insertcells-member(1))|Adds cells to an existing table.|
||[insertColumnsLeft()](/javascript/api/word/word.selection#word-word-selection-insertcolumnsleft-member(1))|Inserts columns to the left of the column that contains the selection.|
||[insertColumnsRight()](/javascript/api/word/word.selection#word-word-selection-insertcolumnsright-member(1))|Inserts columns to the right of the current selection.|
||[insertDateTime(options?: Word.SelectionInsertDateTimeOptions)](/javascript/api/word/word.selection#word-word-selection-insertdatetime-member(1))|Inserts the current date or time, or both, either as text or as a TIME field.|
||[insertFormula(options?: Word.SelectionInsertFormulaOptions)](/javascript/api/word/word.selection#word-word-selection-insertformula-member(1))|Inserts a Formula field at the selection.|
||[insertNewPage()](/javascript/api/word/word.selection#word-word-selection-insertnewpage-member(1))|Inserts a new page at the position of the insertion point.|
||[insertParagraphAfter()](/javascript/api/word/word.selection#word-word-selection-insertparagraphafter-member(1))|Inserts a paragraph mark after the selection.|
||[insertParagraphBefore()](/javascript/api/word/word.selection#word-word-selection-insertparagraphbefore-member(1))|Inserts a new paragraph before the selection or range.|
||[insertParagraphStyleSeparator()](/javascript/api/word/word.selection#word-word-selection-insertparagraphstyleseparator-member(1))|Inserts a special hidden paragraph mark that allows Word to join paragraphs formatted using different paragraph styles.|
||[insertRowsAbove(numRows: number)](/javascript/api/word/word.selection#word-word-selection-insertrowsabove-member(1))|Inserts rows above the current selection.|
||[insertRowsBelow(numRows: number)](/javascript/api/word/word.selection#word-word-selection-insertrowsbelow-member(1))|Inserts rows below the current selection.|
||[insertSymbol(characterNumber: number, options?: Word.SelectionInsertSymbolOptions)](/javascript/api/word/word.selection#word-word-selection-insertsymbol-member(1))|Inserts a symbol in place of the specified selection.|
||[insertText(Text: string)](/javascript/api/word/word.selection#word-word-selection-inserttext-member(1))|Inserts the specified text.|
||[insertXml(xml: string, transform?: string)](/javascript/api/word/word.selection#word-word-selection-insertxml-member(1))|Inserts the specified XML into the document at the cursor, replacing any selected text.|
||[isActive](/javascript/api/word/word.selection#word-word-selection-isactive-member)|Returns whether the selection in the specified window or pane is active.|
||[isColumnSelectModeActive](/javascript/api/word/word.selection#word-word-selection-iscolumnselectmodeactive-member)|Specifies whether column selection mode is active.|
||[isEndOfRowMark](/javascript/api/word/word.selection#word-word-selection-isendofrowmark-member)|Returns whether the selection is at the end-of-row mark in a table.|
||[isEqual(range: Word.Range)](/javascript/api/word/word.selection#word-word-selection-isequal-member(1))|Returns whether the selection is equal to the specified range.|
||[isExtendModeActive](/javascript/api/word/word.selection#word-word-selection-isextendmodeactive-member)|Specifies whether Extend mode is active.|
||[isInRange(range: Word.Range)](/javascript/api/word/word.selection#word-word-selection-isinrange-member(1))|Returns `true` if the selection is contained within the specified range.|
||[isInStory(range: Word.Range)](/javascript/api/word/word.selection#word-word-selection-isinstory-member(1))|Returns whether the selection is in the same story as the specified range.|
||[isInsertionPointAtEndOfLine](/javascript/api/word/word.selection#word-word-selection-isinsertionpointatendofline-member)|Returns whether the insertion point is at the end of a line.|
||[isStartActive](/javascript/api/word/word.selection#word-word-selection-isstartactive-member)|Specifies whether the beginning of the selection is active.|
||[languageDetected](/javascript/api/word/word.selection#word-word-selection-languagedetected-member)|Specifies whether Word has detected the language of the selected text.|
||[languageId](/javascript/api/word/word.selection#word-word-selection-languageid-member)|Returns the language for the selection.|
||[languageIdFarEast](/javascript/api/word/word.selection#word-word-selection-languageidfareast-member)|Returns the East Asian language for the selection.|
||[languageIdOther](/javascript/api/word/word.selection#word-word-selection-languageidother-member)|Returns the language for the selection that isn't classified as an East Asian language.|
||[move(options?: Word.SelectionMoveOptions)](/javascript/api/word/word.selection#word-word-selection-move-member(1))|Collapses the selection to its start or end position and then moves the collapsed object by the specified number of units.|
||[moveDown(options?: Word.SelectionMoveUpDownOptions)](/javascript/api/word/word.selection#word-word-selection-movedown-member(1))|Moves the selection down.|
||[moveEnd(options?: Word.SelectionMoveStartEndOptions)](/javascript/api/word/word.selection#word-word-selection-moveend-member(1))|Moves the ending character position of the range or selection.|
||[moveEndUntil(characters: string, count?: number)](/javascript/api/word/word.selection#word-word-selection-moveenduntil-member(1))|Moves the end position of the selection until any of the specified characters are found in the document.|
||[moveEndWhile(characters: string, count?: number)](/javascript/api/word/word.selection#word-word-selection-moveendwhile-member(1))|Moves the ending character position of the selection while any of the specified characters are found in the document.|
||[moveLeft(options?: Word.SelectionMoveLeftRightOptions)](/javascript/api/word/word.selection#word-word-selection-moveleft-member(1))|Moves the selection to the left.|
||[moveRight(options?: Word.SelectionMoveLeftRightOptions)](/javascript/api/word/word.selection#word-word-selection-moveright-member(1))|Moves the selection to the right.|
||[moveStart(options?: Word.SelectionMoveStartEndOptions)](/javascript/api/word/word.selection#word-word-selection-movestart-member(1))|Moves the start position of the selection.|
||[moveStartUntil(characters: string, count?: number)](/javascript/api/word/word.selection#word-word-selection-movestartuntil-member(1))|Moves the start position of the selection until one of the specified characters is found in the document.|
||[moveStartWhile(characters: string, count?: number)](/javascript/api/word/word.selection#word-word-selection-movestartwhile-member(1))|Moves the start position of the selection while any of the specified characters are found in the document.|
||[moveUntil(characters: string, count?: number)](/javascript/api/word/word.selection#word-word-selection-moveuntil-member(1))|Moves the selection until one of the specified characters is found in the document.|
||[moveUp(options?: Word.SelectionMoveUpDownOptions)](/javascript/api/word/word.selection#word-word-selection-moveup-member(1))|Moves the selection up.|
||[moveWhile(characters: string, count?: number)](/javascript/api/word/word.selection#word-word-selection-movewhile-member(1))|Moves the selection while any of the specified characters are found in the document.|
||[nextField()](/javascript/api/word/word.selection#word-word-selection-nextfield-member(1))|Selects the next field.|
||[nextSubdocument()](/javascript/api/word/word.selection#word-word-selection-nextsubdocument-member(1))|Moves the selection to the next subdocument.|
||[orientation](/javascript/api/word/word.selection#word-word-selection-orientation-member)|Specifies the orientation of text in the selection.|
||[paragraphs](/javascript/api/word/word.selection#word-word-selection-paragraphs-member)|Returns a `ParagraphCollection` object that represents all the paragraphs in the selection.|
||[pasteAndFormat(type: Word.PasteFormatType)](/javascript/api/word/word.selection#word-word-selection-pasteandformat-member(1))|Pastes the content from clipboard and formats them as specified.|
||[pasteExcelTable(linkedToExcel: boolean, wordFormatting: boolean, rtf: boolean)](/javascript/api/word/word.selection#word-word-selection-pasteexceltable-member(1))|Pastes and formats a Microsoft Excel table.|
||[pasteFormat()](/javascript/api/word/word.selection#word-word-selection-pasteformat-member(1))|Applies formatting copied with the `copyFormat` method to the selection.|
||[pasteTableCellsAppendTable()](/javascript/api/word/word.selection#word-word-selection-pastetablecellsappendtable-member(1))|Merges pasted cells into an existing table by inserting the pasted rows between the selected rows.|
||[pasteTableCellsAsNestedTable()](/javascript/api/word/word.selection#word-word-selection-pastetablecellsasnestedtable-member(1))|Pastes a cell or group of cells as a nested table into the selection.|
||[previousField()](/javascript/api/word/word.selection#word-word-selection-previousfield-member(1))|Selects and returns the previous field.|
||[previousSubdocument()](/javascript/api/word/word.selection#word-word-selection-previoussubdocument-member(1))|Moves the selection to the previous subdocument.|
||[range](/javascript/api/word/word.selection#word-word-selection-range-member)|Returns the `Range` object for the portion of the document contained in the selection.|
||[replaceParagraph()](/javascript/api/word/word.selection#word-word-selection-replaceparagraph-member(1))|Replaces the selection with a new paragraph.|
||[sections](/javascript/api/word/word.selection#word-word-selection-sections-member)|Returns the `SectionCollection` object in the selection.|
||[select()](/javascript/api/word/word.selection#word-word-selection-select-member(1))|Selects the current text.|
||[selectCell()](/javascript/api/word/word.selection#word-word-selection-selectcell-member(1))|Selects the entire cell containing the current selection when the selection is in a table.|
||[selectColumn()](/javascript/api/word/word.selection#word-word-selection-selectcolumn-member(1))|Selects the column that contains the insertion point, or selects all columns that contain the selection when the selection is in a table.|
||[selectCurrentAlignment()](/javascript/api/word/word.selection#word-word-selection-selectcurrentalignment-member(1))|Extends the selection forward until text with a different paragraph alignment is encountered.|
||[selectCurrentColor()](/javascript/api/word/word.selection#word-word-selection-selectcurrentcolor-member(1))|Extends the selection forward until text with a different color is encountered.|
||[selectCurrentFont()](/javascript/api/word/word.selection#word-word-selection-selectcurrentfont-member(1))|Extends the selection forward until text in a different font or font size is encountered.|
||[selectCurrentIndent()](/javascript/api/word/word.selection#word-word-selection-selectcurrentindent-member(1))|Extends the selection forward until text with different left or right paragraph indents is encountered.|
||[selectCurrentSpacing()](/javascript/api/word/word.selection#word-word-selection-selectcurrentspacing-member(1))|Extends the selection forward until a paragraph with different line spacing is encountered.|
||[selectCurrentTabs()](/javascript/api/word/word.selection#word-word-selection-selectcurrenttabs-member(1))|Extends the selection forward until a paragraph with different tab stops is encountered.|
||[selectRow()](/javascript/api/word/word.selection#word-word-selection-selectrow-member(1))|Selects the row that contains the insertion point, or selects all rows that contain the selection when the selection is in a table.|
||[sentences](/javascript/api/word/word.selection#word-word-selection-sentences-member)|Returns the `RangeScopedCollection` object for each sentence in the selection.|
||[setRange(start: number, end: number)](/javascript/api/word/word.selection#word-word-selection-setrange-member(1))|Sets the starting and ending character positions for the selection.|
||[shading](/javascript/api/word/word.selection#word-word-selection-shading-member)|Returns the `ShadingUniversal` object for the shading formatting for the selection.|
||[shrink()](/javascript/api/word/word.selection#word-word-selection-shrink-member(1))|Shrinks the selection to the next smaller unit of text.|
||[shrinkDiscontiguousSelection()](/javascript/api/word/word.selection#word-word-selection-shrinkdiscontiguousselection-member(1))|Cancels the selection of all but the most recently selected text when the current selection contains multiple, unconnected selections.|
||[splitTable()](/javascript/api/word/word.selection#word-word-selection-splittable-member(1))|Inserts an empty paragraph above the first row in the selection.|
||[start](/javascript/api/word/word.selection#word-word-selection-start-member)|Specifies the starting character position of the selection.|
||[storyLength](/javascript/api/word/word.selection#word-word-selection-storylength-member)|Returns the number of characters in the story that contains the selection.|
||[storyType](/javascript/api/word/word.selection#word-word-selection-storytype-member)|Returns the story type for the selection.|
||[tables](/javascript/api/word/word.selection#word-word-selection-tables-member)|Returns the `TableCollection` object in the selection.|
||[text](/javascript/api/word/word.selection#word-word-selection-text-member)|Specifies the text in the selection.|
||[toggleCharacterCode()](/javascript/api/word/word.selection#word-word-selection-togglecharactercode-member(1))|Switches the selection between a Unicode character and its corresponding hexadecimal value.|
||[topLevelTables](/javascript/api/word/word.selection#word-word-selection-topleveltables-member)|Returns the tables at the outermost nesting level in the current selection.|
||[type](/javascript/api/word/word.selection#word-word-selection-type-member)|Returns the selection type.|
||[typeBackspace()](/javascript/api/word/word.selection#word-word-selection-typebackspace-member(1))|Deletes the character preceding the selection (if collapsed) or the insertion point.|
||[words](/javascript/api/word/word.selection#word-word-selection-words-member)|Returns the `RangeScopedCollection` object that represents each word in the selection.|
|[SelectionConvertToTableOptions](/javascript/api/word/word.selectionconverttotableoptions)|[applyBorders](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-applyborders-member)|If provided, specifies whether to apply borders to the table of the specified format.|
||[applyColor](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-applycolor-member)|If provided, specifies whether to apply color formatting to the table of the specified format.|
||[applyFirstColumn](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-applyfirstcolumn-member)|If provided, specifies whether to apply special formatting to the first column of the specified format.|
||[applyFont](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-applyfont-member)|If provided, specifies whether to apply font formatting to the table of the specified format.|
||[applyHeadingRows](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-applyheadingrows-member)|If provided, specifies whether to format the first row as a header row of the specified format.|
||[applyLastColumn](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-applylastcolumn-member)|If provided, specifies whether to apply special formatting to the last column of the specified format.|
||[applyLastRow](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-applylastrow-member)|If provided, specifies whether to apply special formatting to the last row of the specified format.|
||[applyShading](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-applyshading-member)|If provided, specifies whether to apply shading to the table of the specified format.|
||[autoFit](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-autofit-member)|If provided, specifies whether to automatically resize the table to fit the contents.|
||[autoFitBehavior](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-autofitbehavior-member)|If provided, specifies the auto-fit behavior for the table.|
||[defaultTableBehavior](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-defaulttablebehavior-member)|If provided, specifies whether Microsoft Word automatically resizes cells in a table to fit the contents.|
||[format](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-format-member)|If provided, specifies a preset format to apply to the table.|
||[initialColumnWidth](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-initialcolumnwidth-member)|If provided, specifies the initial width of each column in the table, in points.|
||[numColumns](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-numcolumns-member)|If provided, specifies the number of columns in the table.|
||[numRows](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-numrows-member)|If provided, specifies the number of rows in the table.|
||[separator](/javascript/api/word/word.selectionconverttotableoptions#word-word-selectionconverttotableoptions-separator-member)|If provided, specifies the character used to separate text into cells.|
|[SelectionDeleteOptions](/javascript/api/word/word.selectiondeleteoptions)|[count](/javascript/api/word/word.selectiondeleteoptions#word-word-selectiondeleteoptions-count-member)|If provided, specifies the number of units to be deleted.|
||[unit](/javascript/api/word/word.selectiondeleteoptions#word-word-selectiondeleteoptions-unit-member)|If provided, specifies the unit by which the collapsed selection is to be deleted.|
|[SelectionGoToOptions](/javascript/api/word/word.selectiongotooptions)|[count](/javascript/api/word/word.selectiongotooptions#word-word-selectiongotooptions-count-member)|If provided, specifies the number of the item in the document.|
||[direction](/javascript/api/word/word.selectiongotooptions#word-word-selectiongotooptions-direction-member)|If provided, specifies the direction the range or selection is moved to.|
||[item](/javascript/api/word/word.selectiongotooptions#word-word-selectiongotooptions-item-member)|If provided, specifies the kind of item the range or selection is moved to.|
||[name](/javascript/api/word/word.selectiongotooptions#word-word-selectiongotooptions-name-member)|If provided, specifies the name if the `item` property is set to Word.GoToItem type `bookmark`, `comment`, `field`, or `object`.|
|[SelectionInsertCrossReferenceOptions](/javascript/api/word/word.selectioninsertcrossreferenceoptions)|[includePosition](/javascript/api/word/word.selectioninsertcrossreferenceoptions#word-word-selectioninsertcrossreferenceoptions-includeposition-member)|If provided, specifies whether to include position.|
||[insertAsHyperlink](/javascript/api/word/word.selectioninsertcrossreferenceoptions#word-word-selectioninsertcrossreferenceoptions-insertashyperlink-member)|If provided, specifies whether to insert the cross-reference as a hyperlink.|
||[separateNumbers](/javascript/api/word/word.selectioninsertcrossreferenceoptions#word-word-selectioninsertcrossreferenceoptions-separatenumbers-member)|If provided, specifies whether to use a separator to separate the numbers from the associated text.|
||[separatorString](/javascript/api/word/word.selectioninsertcrossreferenceoptions#word-word-selectioninsertcrossreferenceoptions-separatorstring-member)|If provided, specifies the string to use as a separator if the `separateNumbers` parameter is set to `true` when the `Selection.insertCrossReference` method is called.|
|[SelectionInsertDateTimeOptions](/javascript/api/word/word.selectioninsertdatetimeoptions)|[calendarType](/javascript/api/word/word.selectioninsertdatetimeoptions#word-word-selectioninsertdatetimeoptions-calendartype-member)|If provided, specifies the calendar type to use when displaying the date or time.|
||[dateLanguage](/javascript/api/word/word.selectioninsertdatetimeoptions#word-word-selectioninsertdatetimeoptions-datelanguage-member)|If provided, specifies the language in which to display the date or time.|
||[dateTimeFormat](/javascript/api/word/word.selectioninsertdatetimeoptions#word-word-selectioninsertdatetimeoptions-datetimeformat-member)|If provided, specifies the format to be used for displaying the date or time, or both.|
||[insertAsField](/javascript/api/word/word.selectioninsertdatetimeoptions#word-word-selectioninsertdatetimeoptions-insertasfield-member)|If provided, specifies whether to insert the specified information as a TIME field.|
||[insertAsFullWidth](/javascript/api/word/word.selectioninsertdatetimeoptions#word-word-selectioninsertdatetimeoptions-insertasfullwidth-member)|If provided, specifies whether to insert the specified information as full-width (double-byte) digits.|
|[SelectionInsertFormulaOptions](/javascript/api/word/word.selectioninsertformulaoptions)|[formula](/javascript/api/word/word.selectioninsertformulaoptions#word-word-selectioninsertformulaoptions-formula-member)|If provided, specifies the mathematical formula you want the = (Formula) field to evaluate.|
||[numberFormat](/javascript/api/word/word.selectioninsertformulaoptions#word-word-selectioninsertformulaoptions-numberformat-member)|If provided, specifies the format for the result of the `= (Formula)` field.|
|[SelectionInsertSymbolOptions](/javascript/api/word/word.selectioninsertsymboloptions)|[bias](/javascript/api/word/word.selectioninsertsymboloptions#word-word-selectioninsertsymboloptions-bias-member)|If provided, specifies the font bias for symbols.|
||[font](/javascript/api/word/word.selectioninsertsymboloptions#word-word-selectioninsertsymboloptions-font-member)|If provided, specifies the name of the font that contains the symbol.|
||[unicode](/javascript/api/word/word.selectioninsertsymboloptions#word-word-selectioninsertsymboloptions-unicode-member)|If provided, specifies whether the character is Unicode.|
|[SelectionMoveLeftRightOptions](/javascript/api/word/word.selectionmoveleftrightoptions)|[count](/javascript/api/word/word.selectionmoveleftrightoptions#word-word-selectionmoveleftrightoptions-count-member)|If provided, specifies the number of units the selection is to be moved.|
||[extend](/javascript/api/word/word.selectionmoveleftrightoptions#word-word-selectionmoveleftrightoptions-extend-member)|If provided, specifies the type of movement.|
||[unit](/javascript/api/word/word.selectionmoveleftrightoptions#word-word-selectionmoveleftrightoptions-unit-member)|If provided, specifies the unit by which the selection is to be moved.|
|[SelectionMoveOptions](/javascript/api/word/word.selectionmoveoptions)|[count](/javascript/api/word/word.selectionmoveoptions#word-word-selectionmoveoptions-count-member)|If provided, specifies the number of units by which the range or selection is to be moved.|
||[unit](/javascript/api/word/word.selectionmoveoptions#word-word-selectionmoveoptions-unit-member)|If provided, specifies the unit by which to move the ending character position.|
|[SelectionMoveStartEndOptions](/javascript/api/word/word.selectionmovestartendoptions)|[count](/javascript/api/word/word.selectionmovestartendoptions#word-word-selectionmovestartendoptions-count-member)|If provided, specifies the number of units to move.|
||[unit](/javascript/api/word/word.selectionmovestartendoptions#word-word-selectionmovestartendoptions-unit-member)|If provided, specifies the unit by which the selection's start or end position (per the calling method) is to be moved.|
|[SelectionMoveUpDownOptions](/javascript/api/word/word.selectionmoveupdownoptions)|[count](/javascript/api/word/word.selectionmoveupdownoptions#word-word-selectionmoveupdownoptions-count-member)|If provided, specifies the number of units the selection is to be moved.|
||[extend](/javascript/api/word/word.selectionmoveupdownoptions#word-word-selectionmoveupdownoptions-extend-member)|If provided, specifies the type of movement.|
||[unit](/javascript/api/word/word.selectionmoveupdownoptions#word-word-selectionmoveupdownoptions-unit-member)|If provided, specifies the unit by which to move the selection.|
|[SelectionNextOptions](/javascript/api/word/word.selectionnextoptions)|[count](/javascript/api/word/word.selectionnextoptions#word-word-selectionnextoptions-count-member)|If provided, specifies the number of units by which you want to move ahead.|
||[unit](/javascript/api/word/word.selectionnextoptions#word-word-selectionnextoptions-unit-member)|If provided, specifies the type of units by which to move the selection.|
|[SelectionPreviousOptions](/javascript/api/word/word.selectionpreviousoptions)|[count](/javascript/api/word/word.selectionpreviousoptions#word-word-selectionpreviousoptions-count-member)|If provided, specifies the number of units by which you want to move.|
||[unit](/javascript/api/word/word.selectionpreviousoptions#word-word-selectionpreviousoptions-unit-member)|If provided, specifies the type of unit by which to move the selection.|
|[SourceCollection](/javascript/api/word/word.sourcecollection)|[getItem(index: number)](/javascript/api/word/word.sourcecollection#word-word-sourcecollection-getitem-member(1))|Gets a `Source` by its index in the collection.|
|[Style](/javascript/api/word/word.style)|[description](/javascript/api/word/word.style#word-word-style-description-member)|Gets the description of the style.|
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
|[Table](/javascript/api/word/word.table)|[applyStyleDirectFormatting(styleName: string)](/javascript/api/word/word.table#word-word-table-applystyledirectformatting-member(1))|Applies the specified style but maintains any formatting that a user directly applies.|
||[autoFitBehavior(behavior: Word.AutoFitBehavior)](/javascript/api/word/word.table#word-word-table-autofitbehavior-member(1))|Determines how Microsoft Word resizes a table when the AutoFit feature is used.|
||[autoFormat(options?: Word.TableAutoFormatOptions)](/javascript/api/word/word.table#word-word-table-autoformat-member(1))|Applies a predefined look to a table.|
||[columns](/javascript/api/word/word.table#word-word-table-columns-member)|Returns the `TableColumnCollection` object that represents the columns in the table.|
||[convertToText(options?: Word.TableConvertToTextOptions)](/javascript/api/word/word.table#word-word-table-converttotext-member(1))|Converts a table to text.|
||[description](/javascript/api/word/word.table#word-word-table-description-member)|Specifies the description of the table.|
||[reapplyAutoFormat()](/javascript/api/word/word.table#word-word-table-reapplyautoformat-member(1))|Updates the table with the characteristics of the predefined table format set when the `autoFormat` method was called.|
||[shading](/javascript/api/word/word.table#word-word-table-shading-member)|Returns the `ShadingUniversal` object that represents the shading of the table.|
||[sort(options?: Word.TableSortOptions)](/javascript/api/word/word.table#word-word-table-sort-member(1))|Sorts the specified table.|
||[title](/javascript/api/word/word.table#word-word-table-title-member)|Specifies the title of the table.|
|[TableAutoFormatOptions](/javascript/api/word/word.tableautoformatoptions)|[applyBorders](/javascript/api/word/word.tableautoformatoptions#word-word-tableautoformatoptions-applyborders-member)|If provided, specifies whether to apply borders of the specified format.|
||[applyColor](/javascript/api/word/word.tableautoformatoptions#word-word-tableautoformatoptions-applycolor-member)|If provided, specifies whether to apply color of the specified format.|
||[applyFirstColumn](/javascript/api/word/word.tableautoformatoptions#word-word-tableautoformatoptions-applyfirstcolumn-member)|If provided, specifies whether to apply first column formatting of the specified format.|
||[applyFont](/javascript/api/word/word.tableautoformatoptions#word-word-tableautoformatoptions-applyfont-member)|If provided, specifies whether to apply font of the specified format.|
||[applyHeadingRows](/javascript/api/word/word.tableautoformatoptions#word-word-tableautoformatoptions-applyheadingrows-member)|If provided, specifies whether to apply heading row formatting of the specified format.|
||[applyLastColumn](/javascript/api/word/word.tableautoformatoptions#word-word-tableautoformatoptions-applylastcolumn-member)|If provided, specifies whether to apply last column formatting of the specified format.|
||[applyLastRow](/javascript/api/word/word.tableautoformatoptions#word-word-tableautoformatoptions-applylastrow-member)|If provided, specifies whether to apply last row formatting of the specified format.|
||[applyShading](/javascript/api/word/word.tableautoformatoptions#word-word-tableautoformatoptions-applyshading-member)|If provided, specifies whether to apply shading of the specified format.|
||[autoFit](/javascript/api/word/word.tableautoformatoptions#word-word-tableautoformatoptions-autofit-member)|If provided, specifies whether to decrease the width of the table columns as much as possible without changing the way text wraps in the cells.|
||[format](/javascript/api/word/word.tableautoformatoptions#word-word-tableautoformatoptions-format-member)|If provided, specifies the format to apply.|
|[TableCell](/javascript/api/word/word.tablecell)|[autoSum()](/javascript/api/word/word.tablecell#word-word-tablecell-autosum-member(1))|Inserts a = (Formula) field that calculates and displays the sum of the values in table cells above or to the left of the cell specified in the expression.|
||[column](/javascript/api/word/word.tablecell#word-word-tablecell-column-member)|Returns the `TableColumn` object that represents the table column that contains this cell.|
||[delete(shiftCells: any)](/javascript/api/word/word.tablecell#word-word-tablecell-delete-member(1))|Deletes the table cell and optionally controls how the remaining cells are shifted.|
||[formula(options?: Word.TableCellFormulaOptions)](/javascript/api/word/word.tablecell#word-word-tablecell-formula-member(1))|Inserts a = (Formula) field that contains the specified formula into a table cell.|
||[merge(mergeTo: Word.TableCell)](/javascript/api/word/word.tablecell#word-word-tablecell-merge-member(1))|Merges this table cell with the specified table cell.|
||[select()](/javascript/api/word/word.tablecell#word-word-tablecell-select-member(1))|Selects the table cell.|
||[shading](/javascript/api/word/word.tablecell#word-word-tablecell-shading-member)|Returns the `ShadingUniversal` object that represents the shading of the table cell.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[autoFit()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-autofit-member(1))|Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.|
||[delete(shiftCells?: Word.DeleteCells)](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-delete-member(1))|Deletes the table cells and optionally controls how the remaining cells are shifted.|
||[distributeHeight()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-distributeheight-member(1))|Adjusts the height of the specified cells so that they are equal.|
||[distributeWidth()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-distributewidth-member(1))|Adjusts the width of the specified cells so that they are equal.|
||[merge()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-merge-member(1))|Merges the specified cells into a single cell.|
||[setHeight(rowHeight: number, heightRule: Word.RowHeightRule)](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-setheight-member(1))|Sets the height of the cells in a table.|
||[setWidth(columnWidth: number, rulerStyle: Word.RulerStyle)](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-setwidth-member(1))|Sets the width of the cells in a table.|
||[split(options?: Word.TableCellCollectionSplitOptions)](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-split-member(1))|Splits this range of table cells.|
|[TableCellCollectionSplitOptions](/javascript/api/word/word.tablecellcollectionsplitoptions)|[mergeBeforeSplit](/javascript/api/word/word.tablecellcollectionsplitoptions#word-word-tablecellcollectionsplitoptions-mergebeforesplit-member)|If provided, specifies whether to merge the cells with one another before splitting them.|
||[numColumns](/javascript/api/word/word.tablecellcollectionsplitoptions#word-word-tablecellcollectionsplitoptions-numcolumns-member)|If provided, specifies the number of columns that the group of cells is to be split into.|
||[numRows](/javascript/api/word/word.tablecellcollectionsplitoptions#word-word-tablecellcollectionsplitoptions-numrows-member)|If provided, specifies the number of rows that the group of cells is to be split into.|
|[TableCellFormulaOptions](/javascript/api/word/word.tablecellformulaoptions)|[formula](/javascript/api/word/word.tablecellformulaoptions#word-word-tablecellformulaoptions-formula-member)|If provided, specifies the mathematical formula you want the = (Formula) field to evaluate.|
||[numFormat](/javascript/api/word/word.tablecellformulaoptions#word-word-tablecellformulaoptions-numformat-member)|If provided, specifies a format for the result of the = (Formula) field.|
|[TableConvertToTextOptions](/javascript/api/word/word.tableconverttotextoptions)|[nestedTables](/javascript/api/word/word.tableconverttotextoptions#word-word-tableconverttotextoptions-nestedtables-member)|If provided, specifies whether to convert the nested tables to text.|
||[separator](/javascript/api/word/word.tableconverttotextoptions#word-word-tableconverttotextoptions-separator-member)|If provided, specifies the character that delimits the converted columns (paragraph marks delimit the converted rows).|
|[TableOfAuthorities](/javascript/api/word/word.tableofauthorities)|[bookmark](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-bookmark-member)|Specifies the name of the bookmark from which to collect table of authorities entries.|
||[category](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-category-member)|Specifies the category of entries to be included in the table of authorities.|
||[delete()](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-delete-member(1))|Deletes this table of authorities.|
||[entrySeparator](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-entryseparator-member)|Specifies a separator of up to five characters.|
||[isCategoryHeaderIncluded](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-iscategoryheaderincluded-member)|Specifies whether the category name for a group of entries appears in the table of authorities.|
||[isEntryFormattingKept](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-isentryformattingkept-member)|Specifies whether the entries in the table of authorities are displayed with their formatting in the table.|
||[isPassimUsed](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-ispassimused-member)|Specifies whether references to the same authority that are repeated on five or more pages are replaced with "Passim".|
||[pageNumberSeparator](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-pagenumberseparator-member)|Specifies a separator of up to five characters.|
||[pageRangeSeparator](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-pagerangeseparator-member)|Specifies a separator of up to five characters.|
||[range](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-range-member)|Gets the portion of a document that is this table of authorities.|
||[sequenceName](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-sequencename-member)|Specifies the Sequence (SEQ) field identifier for the table of authorities.|
||[sequenceSeparator](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-sequenceseparator-member)|Specifies a separator of up to five characters.|
||[tabLeader](/javascript/api/word/word.tableofauthorities#word-word-tableofauthorities-tableader-member)|Specifies the leader character that appears between entries and their associated page numbers in the table of authorities.|
|[TableOfAuthoritiesAddOptions](/javascript/api/word/word.tableofauthoritiesaddoptions)|[bookmark](/javascript/api/word/word.tableofauthoritiesaddoptions#word-word-tableofauthoritiesaddoptions-bookmark-member)|If provided, specifies the string name of the bookmark from which to collect entries for a table of authorities.|
||[category](/javascript/api/word/word.tableofauthoritiesaddoptions#word-word-tableofauthoritiesaddoptions-category-member)|If provided, specifies the category of entries to include in a table of authorities.|
||[entrySeparator](/javascript/api/word/word.tableofauthoritiesaddoptions#word-word-tableofauthoritiesaddoptions-entryseparator-member)|If provided, specifies a separator of up to five characters.|
||[includeCategoryHeader](/javascript/api/word/word.tableofauthoritiesaddoptions#word-word-tableofauthoritiesaddoptions-includecategoryheader-member)|If provided, specifies whether the category name for each group of entries appears in a table of authorities (e.g., "Cases").|
||[keepEntryFormatting](/javascript/api/word/word.tableofauthoritiesaddoptions#word-word-tableofauthoritiesaddoptions-keepentryformatting-member)|If provided, specifies whether the entries in a table of authorities are displayed with their formatting in the table.|
||[pageNumberSeparator](/javascript/api/word/word.tableofauthoritiesaddoptions#word-word-tableofauthoritiesaddoptions-pagenumberseparator-member)|If provided, specifies a separator of up to five characters.|
||[pageRangeSeparator](/javascript/api/word/word.tableofauthoritiesaddoptions#word-word-tableofauthoritiesaddoptions-pagerangeseparator-member)|If provided, specifies a separator of up to five characters.|
||[sequenceName](/javascript/api/word/word.tableofauthoritiesaddoptions#word-word-tableofauthoritiesaddoptions-sequencename-member)|If provided, specifies the string that identifies the Sequence (SEQ) field identifier for a table of authorities.|
||[sequenceSeparator](/javascript/api/word/word.tableofauthoritiesaddoptions#word-word-tableofauthoritiesaddoptions-sequenceseparator-member)|If provided, specifies a separator of up to five characters.|
||[usePassim](/javascript/api/word/word.tableofauthoritiesaddoptions#word-word-tableofauthoritiesaddoptions-usepassim-member)|If provided, specifies whether references to the same authority that are repeated on five or more pages are replaced with "Passim".|
|[TableOfAuthoritiesCategory](/javascript/api/word/word.tableofauthoritiescategory)|[name](/javascript/api/word/word.tableofauthoritiescategory#word-word-tableofauthoritiescategory-name-member)|Specifies the name of this table of authorities category.|
|[TableOfAuthoritiesCategoryCollection](/javascript/api/word/word.tableofauthoritiescategorycollection)|[getCount()](/javascript/api/word/word.tableofauthoritiescategorycollection#word-word-tableofauthoritiescategorycollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItemAt(index: number)](/javascript/api/word/word.tableofauthoritiescategorycollection#word-word-tableofauthoritiescategorycollection-getitemat-member(1))|Returns a `TableOfAuthoritiesCategory` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.tableofauthoritiescategorycollection#word-word-tableofauthoritiescategorycollection-items-member)|Gets the loaded child items in this collection.|
|[TableOfAuthoritiesCollection](/javascript/api/word/word.tableofauthoritiescollection)|[add(range: Word.Range, options?: Word.TableOfAuthoritiesAddOptions)](/javascript/api/word/word.tableofauthoritiescollection#word-word-tableofauthoritiescollection-add-member(1))|Adds a table of authorities to the document at the specified range.|
||[items](/javascript/api/word/word.tableofauthoritiescollection#word-word-tableofauthoritiescollection-items-member)|Gets the loaded child items in this collection.|
||[markAllCitations(shortCitation: string, options?: Word.TableOfAuthoritiesMarkCitationOptions)](/javascript/api/word/word.tableofauthoritiescollection#word-word-tableofauthoritiescollection-markallcitations-member(1))|Inserts a Table of Authorities Entry (TA) field after all instances of the specified citation text.|
||[markCitation(range: Word.Range, shortCitation: string, options?: Word.TableOfAuthoritiesMarkCitationOptions)](/javascript/api/word/word.tableofauthoritiescollection#word-word-tableofauthoritiescollection-markcitation-member(1))|Inserts a Table of Authorities Entry (TA) field at the specified range.|
||[selectNextCitation(shortCitation: string)](/javascript/api/word/word.tableofauthoritiescollection#word-word-tableofauthoritiescollection-selectnextcitation-member(1))|Finds and selects the next instance of the specified citation text.|
|[TableOfAuthoritiesMarkCitationOptions](/javascript/api/word/word.tableofauthoritiesmarkcitationoptions)|[category](/javascript/api/word/word.tableofauthoritiesmarkcitationoptions#word-word-tableofauthoritiesmarkcitationoptions-category-member)|If provided, specifies the category number to be associated with the entry.|
||[longCitation](/javascript/api/word/word.tableofauthoritiesmarkcitationoptions#word-word-tableofauthoritiesmarkcitationoptions-longcitation-member)|If provided, specifies the long citation for the entry as it will appear in a table of authorities.|
||[longCitationAutoText](/javascript/api/word/word.tableofauthoritiesmarkcitationoptions#word-word-tableofauthoritiesmarkcitationoptions-longcitationautotext-member)|If provided, specifies the name of the AutoText entry that contains the text of the long citation as it will appear in a table of authorities.|
|[TableOfContents](/javascript/api/word/word.tableofcontents)|[additionalHeadingStyles](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-additionalheadingstyles-member)|Gets the additional styles used for the table of contents.|
||[areBuiltInHeadingStylesUsed](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-arebuiltinheadingstylesused-member)|Specifies whether built-in heading styles are used for the table of contents.|
||[areFieldsUsed](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-arefieldsused-member)|Specifies whether Table of Contents Entry (TC) fields are included in the table of contents.|
||[areHyperlinksUsedOnWeb](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-arehyperlinksusedonweb-member)|Specifies whether entries in the table of contents should be formatted as hyperlinks when publishing to the web.|
||[arePageNumbersHiddenOnWeb](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-arepagenumbershiddenonweb-member)|Specifies whether the page numbers in the table of contents should be hidden when publishing to the web.|
||[arePageNumbersIncluded](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-arepagenumbersincluded-member)|Specifies whether page numbers are included in the table of contents.|
||[arePageNumbersRightAligned](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-arepagenumbersrightaligned-member)|Specifies whether page numbers are aligned with the right margin in the table of contents.|
||[delete()](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-delete-member(1))|Deletes this table of contents.|
||[lowerHeadingLevel](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-lowerheadinglevel-member)|Specifies the ending heading level for the table of contents.|
||[range](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-range-member)|Gets the portion of a document that is this table of contents.|
||[tabLeader](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-tableader-member)|Specifies the character between entries and their page numbers in the table of contents.|
||[tableId](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-tableid-member)|Specifies a one-letter identifier from TC fields that's used for the table of contents.|
||[updatePageNumbers()](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-updatepagenumbers-member(1))|Updates the entire table of contents.|
||[upperHeadingLevel](/javascript/api/word/word.tableofcontents#word-word-tableofcontents-upperheadinglevel-member)|Specifies the starting heading level for the table of contents.|
|[TableOfContentsAddOptions](/javascript/api/word/word.tableofcontentsaddoptions)|[addedStyles](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-addedstyles-member)|If provided, specifies the string names of additional styles to use for the table of contents.|
||[hidePageNumbersOnWeb](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-hidepagenumbersonweb-member)|If provided, specifies whether the page numbers in a table of contents should be hidden when publishing to the web.|
||[includePageNumbers](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-includepagenumbers-member)|If provided, specifies whether to include page numbers in a table of contents.|
||[lowerHeadingLevel](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-lowerheadinglevel-member)|If provided, specifies the ending heading level for a table of contents and must be a value from 1 to 9.|
||[rightAlignPageNumbers](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-rightalignpagenumbers-member)|If provided, specifies whether page numbers in a table of contents are aligned with the right margin.|
||[tableId](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-tableid-member)|If provided, specifies a one-letter identifier from TC fields that's used for a table of contents.|
||[upperHeadingLevel](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-upperheadinglevel-member)|If provided, specifies the starting heading level for a table of contents and must be a value from 1 to 9.|
||[useBuiltInHeadingStyles](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-usebuiltinheadingstyles-member)|If provided, specifies whether to use built-in heading styles to create a table of contents.|
||[useFields](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-usefields-member)|If provided, specifies whether Table of Contents Entry (TC) fields are used to create a table of contents.|
||[useHyperlinksOnWeb](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-usehyperlinksonweb-member)|If provided, specifies whether entries in a table of contents should be formatted as hyperlinks when the document is published to the web.|
||[useOutlineLevels](/javascript/api/word/word.tableofcontentsaddoptions#word-word-tableofcontentsaddoptions-useoutlinelevels-member)|If provided, specifies whether to use outline levels to create a table of contents.|
|[TableOfContentsCollection](/javascript/api/word/word.tableofcontentscollection)|[add(range: Word.Range, options?: Word.TableOfContentsAddOptions)](/javascript/api/word/word.tableofcontentscollection#word-word-tableofcontentscollection-add-member(1))|Adds a table of contents to the document at the specified range.|
||[items](/javascript/api/word/word.tableofcontentscollection#word-word-tableofcontentscollection-items-member)|Gets the loaded child items in this collection.|
||[markTocEntry(range: Word.Range, options?: Word.TableOfContentsMarkEntryOptions)](/javascript/api/word/word.tableofcontentscollection#word-word-tableofcontentscollection-marktocentry-member(1))|Inserts a Table of Contents Entry (TC) field after the specified range.|
|[TableOfContentsMarkEntryOptions](/javascript/api/word/word.tableofcontentsmarkentryoptions)|[entry](/javascript/api/word/word.tableofcontentsmarkentryoptions#word-word-tableofcontentsmarkentryoptions-entry-member)|If provided, specifies the text that appears in a table of contents or table of figures.|
||[entryAutoText](/javascript/api/word/word.tableofcontentsmarkentryoptions#word-word-tableofcontentsmarkentryoptions-entryautotext-member)|If provided, specifies the AutoText entry name that includes text for the table of figures, or table of contents.|
||[level](/javascript/api/word/word.tableofcontentsmarkentryoptions#word-word-tableofcontentsmarkentryoptions-level-member)|If provided, specifies the level for the entry in a table of contents or table of figures and should be a value from 1 to 9.|
||[tableId](/javascript/api/word/word.tableofcontentsmarkentryoptions#word-word-tableofcontentsmarkentryoptions-tableid-member)|If provided, specifies a one-letter identifier for a table of contents or table of figures (e.g., "i" for an "illustration").|
|[TableOfFigures](/javascript/api/word/word.tableoffigures)|[additionalHeadingStyles](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-additionalheadingstyles-member)|Gets the additional styles used for the table of figures.|
||[areBuiltInHeadingStylesUsed](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-arebuiltinheadingstylesused-member)|Specifies whether built-in heading styles are used for the table of figures.|
||[areFieldsUsed](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-arefieldsused-member)|Specifies whether Table of Contents Entry (TC) fields are included in the table of figures.|
||[areHyperlinksUsedOnWeb](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-arehyperlinksusedonweb-member)|Specifies whether entries in the table of figures should be formatted as hyperlinks when publishing to the web.|
||[arePageNumbersHiddenOnWeb](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-arepagenumbershiddenonweb-member)|Specifies whether the page numbers in the table of figures should be hidden when publishing to the web.|
||[arePageNumbersIncluded](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-arepagenumbersincluded-member)|Specifies whether page numbers are included in the table of figures.|
||[arePageNumbersRightAligned](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-arepagenumbersrightaligned-member)|Specifies whether page numbers are aligned with the right margin in the table of figures.|
||[captionLabel](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-captionlabel-member)|Specifies the label that identifies the items to be included in the table of figures.|
||[delete()](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-delete-member(1))|Deletes this table of figures.|
||[isLabelIncluded](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-islabelincluded-member)|Specifies whether the caption label and caption number are included in the table of figures.|
||[lowerHeadingLevel](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-lowerheadinglevel-member)|Specifies the ending heading level for the table of figures.|
||[range](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-range-member)|Gets the portion of a document that is this table of figures.|
||[tabLeader](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-tableader-member)|Specifies the character between entries and their page numbers in the table of figures.|
||[tableId](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-tableid-member)|Specifies a one-letter identifier from TC fields that's used for the table of figures.|
||[updatePageNumbers()](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-updatepagenumbers-member(1))|Updates the entire table of figures.|
||[upperHeadingLevel](/javascript/api/word/word.tableoffigures#word-word-tableoffigures-upperheadinglevel-member)|Specifies the starting heading level for the table of figures.|
|[TableOfFiguresAddOptions](/javascript/api/word/word.tableoffiguresaddoptions)|[addedStyles](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-addedstyles-member)|If provided, specifies the string names of additional styles to use for the table of figures.|
||[captionLabel](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-captionlabel-member)|If provided, specifies the label that identifies the items to include in a table of figures.|
||[hidePageNumbersOnWeb](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-hidepagenumbersonweb-member)|If provided, specifies whether the page numbers in the table of figures should be hidden when publishing to the web.|
||[includeLabel](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-includelabel-member)|If provided, specifies whether to include the caption label and caption number in a table of figures.|
||[includePageNumbers](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-includepagenumbers-member)|If provided, specifies whether page numbers are included in a table of figures.|
||[lowerHeadingLevel](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-lowerheadinglevel-member)|If provided, specifies the ending heading level for a table of figures when `useBuiltInHeadingStyles` is set to `true`.|
||[rightAlignPageNumbers](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-rightalignpagenumbers-member)|If provided, specifies whether to align page numbers with the right margin in a table of figures.|
||[tableId](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-tableid-member)|If provided, specifies a one-letter identifier from TC fields that's used for a table of figures.|
||[upperHeadingLevel](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-upperheadinglevel-member)|If provided, specifies the starting heading level for a table of figures when `useBuiltInHeadingStyles` is set to `true`.|
||[useBuiltInHeadingStyles](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-usebuiltinheadingstyles-member)|If provided, specifies whether to use built-in heading styles to create a table of figures.|
||[useFields](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-usefields-member)|If provided, specifies whether to use Table of Contents Entry (TC) fields to create a table of figures.|
||[useHyperlinksOnWeb](/javascript/api/word/word.tableoffiguresaddoptions#word-word-tableoffiguresaddoptions-usehyperlinksonweb-member)|If provided, specifies whether entries in a table of figures should be formatted as hyperlinks when the document is published to the web.|
|[TableOfFiguresCollection](/javascript/api/word/word.tableoffigurescollection)|[add(range: Word.Range, options?: Word.TableOfFiguresAddOptions)](/javascript/api/word/word.tableoffigurescollection#word-word-tableoffigurescollection-add-member(1))|Adds a table of figures to the document at the specified range.|
||[items](/javascript/api/word/word.tableoffigurescollection#word-word-tableoffigurescollection-items-member)|Gets the loaded child items in this collection.|
||[markTocEntry(range: Word.Range, options?: Word.TableOfContentsMarkEntryOptions)](/javascript/api/word/word.tableoffigurescollection#word-word-tableoffigurescollection-marktocentry-member(1))|Inserts a Table of Contents Entry (TC) field after the specified range for marking entries in a table of figures.|
|[TableRow](/javascript/api/word/word.tablerow)|[convertToText(options?: Word.TableConvertToTextOptions)](/javascript/api/word/word.tablerow#word-word-tablerow-converttotext-member(1))|Converts the table row to text.|
||[range](/javascript/api/word/word.tablerow#word-word-tablerow-range-member)|Returns the `Range` object that represents the table row.|
||[setHeight(rowHeight: number, heightRule: Word.RowHeightRule)](/javascript/api/word/word.tablerow#word-word-tablerow-setheight-member(1))|Sets the height of the row.|
||[setLeftIndent(leftIndent: number, rulerStyle: Word.RulerStyle)](/javascript/api/word/word.tablerow#word-word-tablerow-setleftindent-member(1))|Sets the left indent for the table row.|
||[shading](/javascript/api/word/word.tablerow#word-word-tablerow-shading-member)|Returns the `ShadingUniversal` object that represents the shading of the table row.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[convertToText(options?: Word.TableConvertToTextOptions)](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-converttotext-member(1))|Converts rows in a table to text.|
||[delete()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-delete-member(1))|Deletes the table rows.|
||[distributeHeight()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-distributeheight-member(1))|Adjusts the height of the rows so that they're equal.|
||[select()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-select-member(1))|Selects the table rows.|
||[setHeight(rowHeight: number, heightRule: Word.RowHeightRule)](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-setheight-member(1))|Sets the height of the cells in a table.|
||[setLeftIndent(leftIndent: number, rulerStyle: Word.RulerStyle)](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-setleftindent-member(1))|Sets the left indent for the table row.|
|[TableSortOptions](/javascript/api/word/word.tablesortoptions)|[bidirectionalSort](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-bidirectionalsort-member)|If provided, specifies whether to use bidirectional sort.|
||[caseSensitive](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-casesensitive-member)|If provided, specifies whether sorting is case-sensitive.|
||[excludeHeader](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-excludeheader-member)|If provided, specifies whether to exclude the header row from the sort operation.|
||[fieldNumber2](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-fieldnumber2-member)|If provided, specifies the second field to sort by.|
||[fieldNumber3](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-fieldnumber3-member)|If provided, specifies the third field to sort by.|
||[fieldNumber](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-fieldnumber-member)|If provided, specifies the first field to sort by.|
||[ignoreArabicThe](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-ignorearabicthe-member)|If provided, specifies whether to ignore Arabic character alef lam when sorting right-to-left language text.|
||[ignoreDiacritics](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-ignorediacritics-member)|If provided, specifies whether to ignore bidirectional control characters when sorting right-to-left language text.|
||[ignoreHebrew](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-ignorehebrew-member)|If provided, specifies whether to ignore Hebrew characters when sorting right-to-left language text.|
||[ignoreKashida](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-ignorekashida-member)|If provided, specifies whether to ignore kashida when sorting right-to-left language text.|
||[languageId](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-languageid-member)|If provided, specifies the sorting language.|
||[sortFieldType2](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-sortfieldtype2-member)|If provided, specifies the type of the second field to sort by.|
||[sortFieldType3](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-sortfieldtype3-member)|If provided, specifies the type of the third field to sort by.|
||[sortFieldType](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-sortfieldtype-member)|If provided, specifies the type of the first field to sort by.|
||[sortOrder2](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-sortorder2-member)|If provided, specifies the sort order of the second field to sort by.|
||[sortOrder3](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-sortorder3-member)|If provided, specifies the sort order of the third field to sort by.|
||[sortOrder](/javascript/api/word/word.tablesortoptions#word-word-tablesortoptions-sortorder-member)|If provided, specifies the sort order of the first field to sort by.|
|[TableStyle](/javascript/api/word/word.tablestyle)|[borders](/javascript/api/word/word.tablestyle#word-word-tablestyle-borders-member)|Returns a `BorderUniversalCollection` that represents all the borders for the table.|
||[columnStripe](/javascript/api/word/word.tablestyle#word-word-tablestyle-columnstripe-member)|Specifies the number of columns in the banding when a style specifies odd- or even-column banding.|
||[condition(conditionCode: Word.ConditionCode)](/javascript/api/word/word.tablestyle#word-word-tablestyle-condition-member(1))|Returns a `ConditionalStyle` object that represents special style formatting for a portion of a table.|
||[isBreakAcrossPagesAllowed](/javascript/api/word/word.tablestyle#word-word-tablestyle-isbreakacrosspagesallowed-member)|Specifies whether Microsoft Word allows to break the specified table across pages.|
||[leftIndent](/javascript/api/word/word.tablestyle#word-word-tablestyle-leftindent-member)|Specifies the left indent value (in points) for the rows in the table style.|
||[rowStripe](/javascript/api/word/word.tablestyle#word-word-tablestyle-rowstripe-member)|Specifies the number of rows to include in the banding when the style specifies odd- or even-row banding.|
||[shading](/javascript/api/word/word.tablestyle#word-word-tablestyle-shading-member)|Returns a `ShadingUniversal` object that refers to the shading formatting for the table style.|
||[tableDirection](/javascript/api/word/word.tablestyle#word-word-tablestyle-tabledirection-member)|Specifies the direction in which Microsoft Word orders cells in the table style.|
|[Template](/javascript/api/word/word.template)|[listTemplates](/javascript/api/word/word.template#word-word-template-listtemplates-member)|Returns a `ListTemplateCollection` object that represents all the list templates in the template.|
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
|[WebSettings](/javascript/api/word/word.websettings)|[allowPng](/javascript/api/word/word.websettings#word-word-websettings-allowpng-member)|Specifies whether PNG is allowed as an image format when saving a document as a webpage.|
||[encoding](/javascript/api/word/word.websettings#word-word-websettings-encoding-member)|Specifies the document encoding (code page or character set) to be used by the web browser when viewing the saved document.|
||[folderSuffix](/javascript/api/word/word.websettings#word-word-websettings-foldersuffix-member)|Gets the folder suffix used when saving a document as a webpage with long file names and supporting files in a separate folder.|
||[organizeInFolder](/javascript/api/word/word.websettings#word-word-websettings-organizeinfolder-member)|Specifies whether supporting files are organized in a separate folder when saving the document as a webpage.|
||[pixelsPerInch](/javascript/api/word/word.websettings#word-word-websettings-pixelsperinch-member)|Specifies the density (pixels per inch) of graphics images and table cells on a webpage.|
||[relyOnCSS](/javascript/api/word/word.websettings#word-word-websettings-relyoncss-member)|Specifies whether cascading style sheets (CSS) are used for font formatting when viewing a saved document in a web browser.|
||[relyOnVectorMarkupLanguage](/javascript/api/word/word.websettings#word-word-websettings-relyonvectormarkuplanguage-member)|Specifies whether image files are not generated from drawing objects when saving a document as a webpage.|
||[screenSize](/javascript/api/word/word.websettings#word-word-websettings-screensize-member)|Specifies the ideal minimum screen size (width by height, in pixels) for viewing the saved document in a web browser.|
||[targetBrowser](/javascript/api/word/word.websettings#word-word-websettings-targetbrowser-member)|Specifies the target browser for documents viewed in a web browser.|
||[useDefaultFolderSuffix()](/javascript/api/word/word.websettings#word-word-websettings-usedefaultfoldersuffix-member(1))|Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.|
||[useLongFileNames](/javascript/api/word/word.websettings#word-word-websettings-uselongfilenames-member)|Specifies whether long file names are used when saving the document as a webpage.|
|[Window](/javascript/api/word/word.window)|[activate()](/javascript/api/word/word.window#word-word-window-activate-member(1))|Activates the window.|
||[areRulersDisplayed](/javascript/api/word/word.window#word-word-window-arerulersdisplayed-member)|Specifies whether rulers are displayed for the window or pane.|
||[areScreenTipsDisplayed](/javascript/api/word/word.window#word-word-window-arescreentipsdisplayed-member)|Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.|
||[areThumbnailsDisplayed](/javascript/api/word/word.window#word-word-window-arethumbnailsdisplayed-member)|Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.|
||[caption](/javascript/api/word/word.window#word-word-window-caption-member)|Specifies the caption text for the window that is displayed in the title bar of the document or application window.|
||[close(options?: Word.WindowCloseOptions)](/javascript/api/word/word.window#word-word-window-close-member(1))|Closes the window.|
||[height](/javascript/api/word/word.window#word-word-window-height-member)|Specifies the height of the window (in points).|
||[horizontalPercentScrolled](/javascript/api/word/word.window#word-word-window-horizontalpercentscrolled-member)|Specifies the horizontal scroll position as a percentage of the document width.|
||[imeMode](/javascript/api/word/word.window#word-word-window-imemode-member)|Specifies the default start-up mode for the Japanese Input Method Editor (IME).|
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
|[WindowCollection](/javascript/api/word/word.windowcollection)||Represents the collection of Word.Window objects.|
|[WindowPageScrollOptions](/javascript/api/word/word.windowpagescrolloptions)|[down](/javascript/api/word/word.windowpagescrolloptions#word-word-windowpagescrolloptions-down-member)|If provided, specifies the number of pages to scroll the window down.|
||[up](/javascript/api/word/word.windowpagescrolloptions#word-word-windowpagescrolloptions-up-member)|If provided, specifies the number of pages to scroll the window up.|
|[WindowScrollOptions](/javascript/api/word/word.windowscrolloptions)|[down](/javascript/api/word/word.windowscrolloptions#word-word-windowscrolloptions-down-member)|If provided, specifies the number of units to scroll the window down.|
||[left](/javascript/api/word/word.windowscrolloptions#word-word-windowscrolloptions-left-member)|If provided, specifies the number of screens to scroll the window to the left.|
||[right](/javascript/api/word/word.windowscrolloptions#word-word-windowscrolloptions-right-member)|If provided, specifies the number of screens to scroll the window to the right.|
||[up](/javascript/api/word/word.windowscrolloptions#word-word-windowscrolloptions-up-member)|If provided, specifies the number of units to scroll the window up.|
|[XmlNode](/javascript/api/word/word.xmlnode)|[attributes](/javascript/api/word/word.xmlnode#word-word-xmlnode-attributes-member)|Gets the attributes for this XML node.|
||[baseName](/javascript/api/word/word.xmlnode#word-word-xmlnode-basename-member)|Gets the name of the element without any prefix.|
||[childNodes](/javascript/api/word/word.xmlnode#word-word-xmlnode-childnodes-member)|Gets the child elements of this XML node.|
||[copy()](/javascript/api/word/word.xmlnode#word-word-xmlnode-copy-member(1))|Copies this XML node, excluding XML markup, to the Clipboard.|
||[cut()](/javascript/api/word/word.xmlnode#word-word-xmlnode-cut-member(1))|Removes this XML node from the document and places it on the Clipboard.|
||[delete()](/javascript/api/word/word.xmlnode#word-word-xmlnode-delete-member(1))|Deletes the XML node from the XML document.|
||[firstChild](/javascript/api/word/word.xmlnode#word-word-xmlnode-firstchild-member)|Gets the first child node if this is a parent node.|
||[hasChildNodes](/javascript/api/word/word.xmlnode#word-word-xmlnode-haschildnodes-member)|Gets whether this XML node has child nodes.|
||[lastChild](/javascript/api/word/word.xmlnode#word-word-xmlnode-lastchild-member)|Gets the last child node if this is a parent node.|
||[level](/javascript/api/word/word.xmlnode#word-word-xmlnode-level-member)|Gets whether this XML element is part of a paragraph, is a paragraph, or is contained within a table cell or contains a table row.|
||[namespaceUri](/javascript/api/word/word.xmlnode#word-word-xmlnode-namespaceuri-member)|Gets the Uniform Resource Identifier (URI) of the schema namespace for this XML node.|
||[nextSibling](/javascript/api/word/word.xmlnode#word-word-xmlnode-nextsibling-member)|Gets the next element in the document that's at the same level as this XML node.|
||[nodeType](/javascript/api/word/word.xmlnode#word-word-xmlnode-nodetype-member)|Gets the type of node.|
||[nodeValue](/javascript/api/word/word.xmlnode#word-word-xmlnode-nodevalue-member)|Specifies the value of this XML node.|
||[ownerDocument](/javascript/api/word/word.xmlnode#word-word-xmlnode-ownerdocument-member)|Gets the parent document of this XML node.|
||[parentNode](/javascript/api/word/word.xmlnode#word-word-xmlnode-parentnode-member)|Gets the parent element of this XML node.|
||[placeholderText](/javascript/api/word/word.xmlnode#word-word-xmlnode-placeholdertext-member)|Specifies the text displayed for this element if it contains no text.|
||[previousSibling](/javascript/api/word/word.xmlnode#word-word-xmlnode-previoussibling-member)|Gets the previous element in the document that's at the same level as this XML node.|
||[range](/javascript/api/word/word.xmlnode#word-word-xmlnode-range-member)|Gets the portion of a document that is contained in this XML node.|
||[removeChild(childElement: Word.XmlNode)](/javascript/api/word/word.xmlnode#word-word-xmlnode-removechild-member(1))|Removes a child element from this XML node.|
||[selectNodes(xPath: string, options?: Word.SelectNodesOptions)](/javascript/api/word/word.xmlnode#word-word-xmlnode-selectnodes-member(1))|Returns all the child elements that match the XPath parameter, in the order in which they appear within this XML node.|
||[selectSingleNode(xPath: string, options?: Word.SelectSingleNodeOptions)](/javascript/api/word/word.xmlnode#word-word-xmlnode-selectsinglenode-member(1))|Returns the first child element that matches the XPath parameter within this XML node.|
||[setValidationError(status: Word.XmlValidationStatus, options?: Word.XmlNodeSetValidationErrorOptions)](/javascript/api/word/word.xmlnode#word-word-xmlnode-setvalidationerror-member(1))|Changes the validation error text displayed to a user for this XML node and whether to force Word to report the node as invalid.|
||[text](/javascript/api/word/word.xmlnode#word-word-xmlnode-text-member)|Specifies the text contained within the XML element.|
||[validate()](/javascript/api/word/word.xmlnode#word-word-xmlnode-validate-member(1))|Validates this XML node against the XML schemas that are attached to the document.|
||[validationErrorText](/javascript/api/word/word.xmlnode#word-word-xmlnode-validationerrortext-member)|Gets the description for a validation error on this `XmlNode` object.|
||[validationStatus](/javascript/api/word/word.xmlnode#word-word-xmlnode-validationstatus-member)|Gets whether this element is valid according to the attached schema.|
|[XmlNodeCollection](/javascript/api/word/word.xmlnodecollection)|[getItem(index: number)](/javascript/api/word/word.xmlnodecollection#word-word-xmlnodecollection-getitem-member(1))|Gets a `XmlNode` object by its index in the collection.|
||[getItemAt(index: number)](/javascript/api/word/word.xmlnodecollection#word-word-xmlnodecollection-getitemat-member(1))|Returns an individual `XmlNode` object in a collection.|
||[items](/javascript/api/word/word.xmlnodecollection#word-word-xmlnodecollection-items-member)|Gets the loaded child items in this collection.|
|[XmlNodeSetValidationErrorOptions](/javascript/api/word/word.xmlnodesetvalidationerroroptions)|[clearedAutomatically](/javascript/api/word/word.xmlnodesetvalidationerroroptions#word-word-xmlnodesetvalidationerroroptions-clearedautomatically-member)|If provided, specifies whether the validation error should be cleared automatically.|
||[errorText](/javascript/api/word/word.xmlnodesetvalidationerroroptions#word-word-xmlnodesetvalidationerroroptions-errortext-member)|If provided, specifies the error text to display for the validation error.|
