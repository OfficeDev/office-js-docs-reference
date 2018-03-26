// Type definitions for Office.js
// Project: http://dev.office.com
// Definitions by: OfficeDev <https://github.com/OfficeDev>, Lance Austin <https://github.com/LanceEA>, Michael Zlatkovsky <https://github.com/Zlatkovsky>, Kim Brandl <https://github.com/kbrandl>, Ricky Kirkham <https://github.com/Rick-Kirkham>
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped

/*
office-js
Copyright (c) Microsoft Corporation
*/

////////////////////////////////////////////////////////////////
/////////////////////// Begin Word APIs ////////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Word {
    /**
     *
     * Represents the body of a document or a section.
     *
     * [Api set: WordApi 1.1]
     */
    export class Body extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the collection of rich text content control objects in the body. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inlinePicture objects in the body. The collection does not include floating images. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of list objects in the body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        lists: Word.ListCollection;
        /**
         *
         * Gets the collection of paragraph objects in the body. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws if there isn't a parent body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBody: Word.Body;
        /**
         *
         * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Returns a null object if there isn't a parent body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBodyOrNullObject: Word.Body;
        /**
         *
         * Gets the content control that contains the body. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the body. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the parent section of the body. Throws if there isn't a parent section. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentSection: Word.Section;
        /**
         *
         * Gets the parent section of the body. Returns a null object if there isn't a parent section. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentSectionOrNullObject: Word.Section;
        /**
         *
         * Gets the collection of table objects in the body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        tables: Word.TableCollection;
        /**
         *
         * Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         *
         * Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: string;
        /**
         *
         * Gets the text of the body. Use the insertText method to insert text. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        text: string;
        /**
         *
         * Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        type: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.BodyUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Body): void;
        /**
         *
         * Clears the contents of the body object. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         *
         * Gets the HTML representation of the body object.
         *
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the OOXML (Office Open XML) representation of the body object.
         *
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the whole body, or the starting or ending point of the body, as a range.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', 'After' or 'Content'.
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add to the body.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the body object with a Rich Text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the document.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the body.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: string, values?: Array<Array<string>>): Word.Table;
        /**
         *
         * Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects.
         *
         * [Api set: WordApi 1.1]
         *
         * @param searchText - Required. The search text.
         * @param searchOptions - Optional. Options for the search.
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects the body and navigates the Word UI to it.
         *
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Body;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Body;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Body;
        toJSON(): {
            "font": Font;
            "style": string;
            "styleBuiltIn": string;
            "text": string;
            "type": string;
        };
    }
    /**
     *
     * Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
     *
     * [Api set: WordApi 1.1]
     */
    export class ContentControl extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the collection of content control objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of list objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        lists: Word.ListCollection;
        /**
         *
         * Get the collection of paragraph objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the parent body of the content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the content control. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the content control. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the table that contains the content control. Throws if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the content control. Throws if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCell: Word.TableCell;
        /**
         *
         * Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCellOrNullObject: Word.TableCell;
        /**
         *
         * Gets the table that contains the content control. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableOrNullObject: Word.Table;
        /**
         *
         * Gets the collection of table objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        tables: Word.TableCollection;
        /**
         *
         * Gets or sets the appearance of the content control. The value can be 'boundingBox', 'tags' or 'hidden'.
         *
         * [Api set: WordApi 1.1]
         */
        appearance: string;
        /**
         *
         * Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
         *
         * [Api set: WordApi 1.1]
         */
        cannotDelete: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the user can edit the contents of the content control.
         *
         * [Api set: WordApi 1.1]
         */
        cannotEdit: boolean;
        /**
         *
         * Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
         *
         * [Api set: WordApi 1.1]
         */
        color: string;
        /**
         *
         * Gets an integer that represents the content control identifier. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        id: number;
        /**
         *
         * Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
         *
         * [Api set: WordApi 1.1]
         */
        placeholderText: string;
        /**
         *
         * Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
         *
         * [Api set: WordApi 1.1]
         */
        removeWhenEdited: boolean;
        /**
         *
         * Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         *
         * Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: string;
        /**
         *
         * Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        subtype: string;
        /**
         *
         * Gets or sets a tag to identify a content control.
         *
         * [Api set: WordApi 1.1]
         */
        tag: string;
        /**
         *
         * Gets the text of the content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        text: string;
        /**
         *
         * Gets or sets the title for a content control.
         *
         * [Api set: WordApi 1.1]
         */
        title: string;
        /**
         *
         * Gets the content control type. Only rich text content controls are supported currently. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        type: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.ContentControlUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: ContentControl): void;
        /**
         *
         * Clears the contents of the content control. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         *
         * Deletes the content control and its content. If keepContent is set to true, the content is not deleted.
         *
         * [Api set: WordApi 1.1]
         *
         * @param keepContent - Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted.
         */
        delete(keepContent: boolean): void;
        /**
         *
         * Gets the HTML representation of the content control object.
         *
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the Office Open XML (OOXML) representation of the content control object.
         *
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the whole content control, or the starting or ending point of the content control, as a range.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Before', 'Start', 'End', 'After' or 'Content'.
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Gets the text ranges in the content control by using punctuation marks and/or other ending marks.
         *
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        getTextRanges(endingMarks: Array<string>, trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before' or 'After'. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. Type of break.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before' or 'After'.
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragrph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before' or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before' or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: string, values?: Array<Array<string>>): Word.Table;
        /**
         *
         * Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. The text to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects.
         *
         * [Api set: WordApi 1.1]
         *
         * @param searchText - Required. The search text.
         * @param searchOptions - Optional. Options for the search.
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects the content control. This causes Word to scroll to the selection.
         *
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        /**
         *
         * Splits the content control into child ranges by using delimiters.
         *
         * [Api set: WordApi 1.3]
         *
         * @param delimiters - Required. The delimiters as an array of strings.
         * @param multiParagraphs - Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.
         * @param trimDelimiters - Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        split(delimiters: Array<string>, multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.ContentControl;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ContentControl;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ContentControl;
        toJSON(): {
            "appearance": string;
            "cannotDelete": boolean;
            "cannotEdit": boolean;
            "color": string;
            "font": Font;
            "id": number;
            "placeholderText": string;
            "removeWhenEdited": boolean;
            "style": string;
            "styleBuiltIn": string;
            "subtype": string;
            "tag": string;
            "text": string;
            "title": string;
            "type": string;
        };
    }
    /**
     *
     * Contains a collection of [contentControl](contentControl.md) objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
     *
     * [Api set: WordApi 1.1]
     */
    export class ContentControlCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        items: Array<Word.ContentControl>;
        /**
         *
         * Gets a content control by its identifier. Throws if there isn't a content control with the identifier in this collection.
         *
         * [Api set: WordApi 1.1]
         *
         * @param id - Required. A content control identifier.
         */
        getById(id: number): Word.ContentControl;
        /**
         *
         * Gets a content control by its identifier. Returns a null object if there isn't a content control with the identifier in this collection.
         *
         * [Api set: WordApi 1.3]
         *
         * @param id - Required. A content control identifier.
         */
        getByIdOrNullObject(id: number): Word.ContentControl;
        /**
         *
         * Gets the content controls that have the specified tag.
         *
         * [Api set: WordApi 1.1]
         *
         * @param tag - Required. A tag set on a content control.
         */
        getByTag(tag: string): Word.ContentControlCollection;
        /**
         *
         * Gets the content controls that have the specified title.
         *
         * [Api set: WordApi 1.1]
         *
         * @param title - Required. The title of a content control.
         */
        getByTitle(title: string): Word.ContentControlCollection;
        /**
         *
         * Gets the content controls that have the specified types and/or subtypes.
         *
         * [Api set: WordApi 1.3]
         *
         * @param types - Required. An array of content control types and/or subtypes.
         */
        getByTypes(types: Array<string>): Word.ContentControlCollection;
        /**
         *
         * Gets the first content control in this collection. Throws if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.ContentControl;
        /**
         *
         * Gets the first content control in this collection. Returns a null object if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.ContentControl;
        /**
         *
         * Gets a content control by its index in the collection.
         *
         * [Api set: WordApi 1.1]
         *
         * @param index - The index.
         */
        getItem(index: number): Word.ContentControl;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.ContentControlCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ContentControlCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ContentControlCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a custom property.
     *
     * [Api set: WordApi 1.3]
     */
    export class CustomProperty extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the key of the custom property. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        key: string;
        /**
         *
         * Gets the value type of the custom property. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        type: string;
        /**
         *
         * Gets or sets the value of the custom property.
         *
         * [Api set: WordApi 1.3]
         */
        value: any;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.CustomPropertyUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: CustomProperty): void;
        /**
         *
         * Deletes the custom property.
         *
         * [Api set: WordApi 1.3]
         */
        delete(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.CustomProperty;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.CustomProperty;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.CustomProperty;
        toJSON(): {
            "key": string;
            "type": string;
            "value": any;
        };
    }
    /**
     *
     * Contains the collection of [customProperty](customProperty.md) objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class CustomPropertyCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        items: Array<Word.CustomProperty>;
        /**
         *
         * Creates a new or sets an existing custom property.
         *
         * [Api set: WordApi 1.3]
         *
         * @param key - Required. The custom property's key, which is case-insensitive.
         * @param value - Required. The custom property's value.
         */
        add(key: string, value: any): Word.CustomProperty;
        /**
         *
         * Deletes all custom properties in this collection.
         *
         * [Api set: WordApi 1.3]
         */
        deleteAll(): void;
        /**
         *
         * Gets the count of custom properties.
         *
         * [Api set: WordApi 1.3]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.
         *
         * [Api set: WordApi 1.3]
         *
         * @param key - The key that identifies the custom property object.
         */
        getItem(key: string): Word.CustomProperty;
        /**
         *
         * Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.
         *
         * [Api set: WordApi 1.3]
         *
         * @param key - Required. The key that identifies the custom property object.
         */
        getItemOrNullObject(key: string): Word.CustomProperty;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.CustomPropertyCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.CustomPropertyCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.CustomPropertyCollection;
        toJSON(): {};
    }
    /**
     *
     * The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
     *
     * [Api set: WordApi 1.1]
     */
    export class Document extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        body: Word.Body;
        /**
         *
         * Gets the collection of content control objects in the current document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the properties of the current document. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        properties: Word.DocumentProperties;
        /**
         *
         * Gets the collection of section objects in the document. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        sections: Word.SectionCollection;
        /**
         *
         * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        saved: boolean;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.DocumentUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Document): void;
        /**
         *
         * Gets the current selection of the document. Multiple selections are not supported.
         *
         * [Api set: WordApi 1.1]
         */
        getSelection(): Word.Range;
        /**
         *
         * Saves the document. This will use the Word default file naming convention if the document has not been saved before.
         *
         * [Api set: WordApi 1.1]
         */
        save(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Document;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Document;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Document;
        toJSON(): {
            "body": Body;
            "properties": DocumentProperties;
            "saved": boolean;
        };
    }
    /**
     *
     * Represents document properties.
     *
     * [Api set: WordApi 1.3]
     */
    export class DocumentProperties extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the collection of custom properties of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        customProperties: Word.CustomPropertyCollection;
        /**
         *
         * Gets the application name of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        applicationName: string;
        /**
         *
         * Gets or sets the author of the document.
         *
         * [Api set: WordApi 1.3]
         */
        author: string;
        /**
         *
         * Gets or sets the category of the document.
         *
         * [Api set: WordApi 1.3]
         */
        category: string;
        /**
         *
         * Gets or sets the comments of the document.
         *
         * [Api set: WordApi 1.3]
         */
        comments: string;
        /**
         *
         * Gets or sets the company of the document.
         *
         * [Api set: WordApi 1.3]
         */
        company: string;
        /**
         *
         * Gets the creation date of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        creationDate: Date;
        /**
         *
         * Gets or sets the format of the document.
         *
         * [Api set: WordApi 1.3]
         */
        format: string;
        /**
         *
         * Gets or sets the keywords of the document.
         *
         * [Api set: WordApi 1.3]
         */
        keywords: string;
        /**
         *
         * Gets the last author of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        lastAuthor: string;
        /**
         *
         * Gets the last print date of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        lastPrintDate: Date;
        /**
         *
         * Gets the last save time of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        lastSaveTime: Date;
        /**
         *
         * Gets or sets the manager of the document.
         *
         * [Api set: WordApi 1.3]
         */
        manager: string;
        /**
         *
         * Gets the revision number of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        revisionNumber: string;
        /**
         *
         * Gets the security of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        security: number;
        /**
         *
         * Gets or sets the subject of the document.
         *
         * [Api set: WordApi 1.3]
         */
        subject: string;
        /**
         *
         * Gets the template of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        template: string;
        /**
         *
         * Gets or sets the title of the document.
         *
         * [Api set: WordApi 1.3]
         */
        title: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.DocumentPropertiesUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: DocumentProperties): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.DocumentProperties;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.DocumentProperties;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.DocumentProperties;
        toJSON(): {
            "applicationName": string;
            "author": string;
            "category": string;
            "comments": string;
            "company": string;
            "creationDate": Date;
            "format": string;
            "keywords": string;
            "lastAuthor": string;
            "lastPrintDate": Date;
            "lastSaveTime": Date;
            "manager": string;
            "revisionNumber": string;
            "security": number;
            "subject": string;
            "template": string;
            "title": string;
        };
    }
    /**
     *
     * Represents a font.
     *
     * [Api set: WordApi 1.1]
     */
    export class Font extends OfficeExtension.ClientObject {
        /**
         *
         * Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        bold: boolean;
        /**
         *
         * Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
         *
         * [Api set: WordApi 1.1]
         */
        color: string;
        /**
         *
         * Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        doubleStrikeThrough: boolean;
        /**
         *
         * Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, or an empty string for mixed highlight colors, or null for no highlight color.
         *
         * [Api set: WordApi 1.1]
         */
        highlightColor: string;
        /**
         *
         * Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        italic: boolean;
        /**
         *
         * Gets or sets a value that represents the name of the font.
         *
         * [Api set: WordApi 1.1]
         */
        name: string;
        /**
         *
         * Gets or sets a value that represents the font size in points.
         *
         * [Api set: WordApi 1.1]
         */
        size: number;
        /**
         *
         * Gets or sets a value that indicates whether the font has a strike through. True if the font is formatted as strikethrough text, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        strikeThrough: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        subscript: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        superscript: boolean;
        /**
         *
         * Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
         *
         * [Api set: WordApi 1.1]
         */
        underline: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.FontUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Font): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Font;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Font;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Font;
        toJSON(): {
            "bold": boolean;
            "color": string;
            "doubleStrikeThrough": boolean;
            "highlightColor": string;
            "italic": boolean;
            "name": string;
            "size": number;
            "strikeThrough": boolean;
            "subscript": boolean;
            "superscript": boolean;
            "underline": string;
        };
    }
    /**
     *
     * Represents an inline picture.
     *
     * [Api set: WordApi 1.1]
     */
    export class InlinePicture extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the parent paragraph that contains the inline image. Read-only.
         *
         * [Api set: WordApi 1.2]
         */
        paragraph: Word.Paragraph;
        /**
         *
         * Gets the content control that contains the inline image. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the table that contains the inline image. Throws if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the inline image. Throws if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCell: Word.TableCell;
        /**
         *
         * Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCellOrNullObject: Word.TableCell;
        /**
         *
         * Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableOrNullObject: Word.Table;
        /**
         *
         * Gets or sets a string that represents the alternative text associated with the inline image
         *
         * [Api set: WordApi 1.1]
         */
        altTextDescription: string;
        /**
         *
         * Gets or sets a string that contains the title for the inline image.
         *
         * [Api set: WordApi 1.1]
         */
        altTextTitle: string;
        /**
         *
         * Gets or sets a number that describes the height of the inline image.
         *
         * [Api set: WordApi 1.1]
         */
        height: number;
        /**
         *
         * Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
         *
         * [Api set: WordApi 1.1]
         */
        hyperlink: string;
        /**
         *
         * Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
         *
         * [Api set: WordApi 1.1]
         */
        lockAspectRatio: boolean;
        /**
         *
         * Gets or sets a number that describes the width of the inline image.
         *
         * [Api set: WordApi 1.1]
         */
        width: number;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.InlinePictureUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: InlinePicture): void;
        /**
         *
         * Deletes the inline picture from the document.
         *
         * [Api set: WordApi 1.2]
         */
        delete(): void;
        /**
         *
         * Gets the base64 encoded string representation of the inline image.
         *
         * [Api set: WordApi 1.1]
         */
        getBase64ImageSrc(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the next inline image. Throws if this inline image is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.InlinePicture;
        /**
         *
         * Gets the next inline image. Returns a null object if this inline image is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.InlinePicture;
        /**
         *
         * Gets the picture, or the starting or ending point of the picture, as a range.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start' or 'End'.
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param breakType - Required. The break type to add.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the inline picture with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Before' or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Selects the inline picture. This causes Word to scroll to the selection.
         *
         * [Api set: WordApi 1.2]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.InlinePicture;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.InlinePicture;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.InlinePicture;
        toJSON(): {
            "altTextDescription": string;
            "altTextTitle": string;
            "height": number;
            "hyperlink": string;
            "lockAspectRatio": boolean;
            "width": number;
        };
    }
    /**
     *
     * Contains a collection of [inlinePicture](inlinePicture.md) objects.
     *
     * [Api set: WordApi 1.1]
     */
    export class InlinePictureCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        items: Array<Word.InlinePicture>;
        /**
         *
         * Gets the first inline image in this collection. Throws if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.InlinePicture;
        /**
         *
         * Gets the first inline image in this collection. Returns a null object if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.InlinePicture;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.InlinePictureCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.InlinePictureCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.InlinePictureCollection;
        toJSON(): {};
    }
    /**
     *
     * Contains a collection of [paragraph](paragraph.md) objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class List extends OfficeExtension.ClientObject {
        /**
         *
         * Gets paragraphs in the list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the list's id.
         *
         * [Api set: WordApi 1.3]
         */
        id: number;
        /**
         *
         * Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        levelExistences: Array<boolean>;
        /**
         *
         * Gets all 9 level types in the list. Each type can be 'Bullet', 'Number' or 'Picture'. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        levelTypes: Array<string>;
        /**
         *
         * Gets the paragraphs that occur at the specified level in the list.
         *
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         */
        getLevelParagraphs(level: number): Word.ParagraphCollection;
        /**
         *
         * Gets the bullet, number or picture at the specified level as a string.
         *
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         */
        getLevelString(level: number): OfficeExtension.ClientResult<string>;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Sets the alignment of the bullet, number or picture at the specified level in the list.
         *
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param alignment - Required. The level alignment that can be 'left', 'centered' or 'right'.
         */
        setLevelAlignment(level: number, alignment: string): void;
        /**
         *
         * Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.
         *
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param listBullet - Required. The bullet.
         * @param charCode - Optional. The bullet character's code value. Used only if the bullet is 'Custom'.
         * @param fontName - Optional. The bullet's font name. Used only if the bullet is 'Custom'.
         */
        setLevelBullet(level: number, listBullet: string, charCode?: number, fontName?: string): void;
        /**
         *
         * Sets the two indents of the specified level in the list.
         *
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param textIndent - Required. The text indent in points. It is the same as paragraph left indent.
         * @param textIndent - Required. The relative indent, in points, of the bullet, number or picture. It is the same as paragraph first line indent.
         */
        setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number): void;
        /**
         *
         * Sets the numbering format at the specified level in the list.
         *
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param listNumbering - Required. The ordinal format.
         * @param formatString - Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.
         */
        setLevelNumbering(level: number, listNumbering: string, formatString?: Array<any>): void;
        /**
         *
         * Sets the starting number at the specified level in the list. Default value is 1.
         *
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param startingNumber - Required. The number to start with.
         */
        setLevelStartingNumber(level: number, startingNumber: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.List;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.List;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.List;
        toJSON(): {
            "id": number;
            "levelExistences": boolean[];
            "levelTypes": string[];
        };
    }
    /**
     *
     * Contains a collection of [list](list.md) objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class ListCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        items: Array<Word.List>;
        /**
         *
         * Gets a list by its identifier. Throws if there isn't a list with the identifier in this collection.
         *
         * [Api set: WordApi 1.3]
         *
         * @param id - Required. A list identifier.
         */
        getById(id: number): Word.List;
        /**
         *
         * Gets a list by its identifier. Returns a null object if there isn't a list with the identifier in this collection.
         *
         * [Api set: WordApi 1.3]
         *
         * @param id - Required. A list identifier.
         */
        getByIdOrNullObject(id: number): Word.List;
        /**
         *
         * Gets the first list in this collection. Throws if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.List;
        /**
         *
         * Gets the first list in this collection. Returns a null object if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.List;
        /**
         *
         * Gets a list object by its index in the collection.
         *
         * [Api set: WordApi 1.3]
         *
         * @param index - A number that identifies the index location of a list object.
         */
        getItem(index: number): Word.List;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.ListCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ListCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ListCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents the paragraph list item format.
     *
     * [Api set: WordApi 1.3]
     */
    export class ListItem extends OfficeExtension.ClientObject {
        /**
         *
         * Gets or sets the level of the item in the list.
         *
         * [Api set: WordApi 1.3]
         */
        level: number;
        /**
         *
         * Gets the list item bullet, number or picture as a string. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        listString: string;
        /**
         *
         * Gets the list item order number in relation to its siblings. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        siblingIndex: number;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.ListItemUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: ListItem): void;
        /**
         *
         * Gets the list item parent, or the closest ancestor if the parent does not exist. Throws if the list item has no ancester.
         *
         * [Api set: WordApi 1.3]
         *
         * @param parentOnly - Optional. Specified only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.
         */
        getAncestor(parentOnly?: boolean): Word.Paragraph;
        /**
         *
         * Gets the list item parent, or the closest ancestor if the parent does not exist. Returns a null object if the list item has no ancester.
         *
         * [Api set: WordApi 1.3]
         *
         * @param parentOnly - Optional. Specified only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.
         */
        getAncestorOrNullObject(parentOnly?: boolean): Word.Paragraph;
        /**
         *
         * Gets all descendant list items of the list item.
         *
         * [Api set: WordApi 1.3]
         *
         * @param directChildrenOnly - Optional. Specified only the list item's direct children will be returned. The default is false that indicates to get all descendant items.
         */
        getDescendants(directChildrenOnly?: boolean): Word.ParagraphCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.ListItem;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ListItem;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ListItem;
        toJSON(): {
            "level": number;
            "listString": string;
            "siblingIndex": number;
        };
    }
    /**
     *
     * Represents a single paragraph in a selection, range, content control, or document body.
     *
     * [Api set: WordApi 1.1]
     */
    export class Paragraph extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the collection of content control objects in the paragraph. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inlinePicture objects in the paragraph. The collection does not include floating images. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the List to which this paragraph belongs. Throws if the paragraph is not in a list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        list: Word.List;
        /**
         *
         * Gets the ListItem for the paragraph. Throws if the paragraph is not part of a list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        listItem: Word.ListItem;
        /**
         *
         * Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        listItemOrNullObject: Word.ListItem;
        /**
         *
         * Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        listOrNullObject: Word.List;
        /**
         *
         * Gets the parent body of the paragraph. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the paragraph. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the table that contains the paragraph. Throws if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the paragraph. Throws if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCell: Word.TableCell;
        /**
         *
         * Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCellOrNullObject: Word.TableCell;
        /**
         *
         * Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableOrNullObject: Word.Table;
        /**
         *
         * Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
         *
         * [Api set: WordApi 1.1]
         */
        alignment: string;
        /**
         *
         * Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
         *
         * [Api set: WordApi 1.1]
         */
        firstLineIndent: number;
        /**
         *
         * Indicates the paragraph is the last one inside its parent body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        isLastParagraph: boolean;
        /**
         *
         * Checks whether the paragraph is a list item. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        isListItem: boolean;
        /**
         *
         * Gets or sets the left indent value, in points, for the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        leftIndent: number;
        /**
         *
         * Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
         *
         * [Api set: WordApi 1.1]
         */
        lineSpacing: number;
        /**
         *
         * Gets or sets the amount of spacing, in grid lines. after the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        lineUnitAfter: number;
        /**
         *
         * Gets or sets the amount of spacing, in grid lines, before the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        lineUnitBefore: number;
        /**
         *
         * Gets or sets the outline level for the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        outlineLevel: number;
        /**
         *
         * Gets or sets the right indent value, in points, for the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        rightIndent: number;
        /**
         *
         * Gets or sets the spacing, in points, after the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        spaceAfter: number;
        /**
         *
         * Gets or sets the spacing, in points, before the paragraph.
         *
         * [Api set: WordApi 1.1]
         */
        spaceBefore: number;
        /**
         *
         * Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         *
         * Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: string;
        /**
         *
         * Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        tableNestingLevel: number;
        /**
         *
         * Gets the text of the paragraph. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        text: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.ParagraphUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Paragraph): void;
        /**
         *
         * Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.
         *
         * [Api set: WordApi 1.3]
         *
         * @param listId - Required. The ID of an existing list.
         * @param level - Required. The level in the list.
         */
        attachToList(listId: number, level: number): Word.List;
        /**
         *
         * Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         *
         * Deletes the paragraph and its content from the document.
         *
         * [Api set: WordApi 1.1]
         */
        delete(): void;
        /**
         *
         * Moves this paragraph out of its list, if the paragraph is a list item.
         *
         * [Api set: WordApi 1.3]
         */
        detachFromList(): void;
        /**
         *
         * Gets the HTML representation of the paragraph object.
         *
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the next paragraph. Throws if the paragraph is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.Paragraph;
        /**
         *
         * Gets the next paragraph. Returns a null object if the paragraph is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.Paragraph;
        /**
         *
         * Gets the Office Open XML (OOXML) representation of the paragraph object.
         *
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets the previous paragraph. Throws if the paragraph is the first one.
         *
         * [Api set: WordApi 1.3]
         */
        getPrevious(): Word.Paragraph;
        /**
         *
         * Gets the previous paragraph. Returns a null object if the paragraph is the first one.
         *
         * [Api set: WordApi 1.3]
         */
        getPreviousOrNullObject(): Word.Paragraph;
        /**
         *
         * Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', 'After' or 'Content'.
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.
         *
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        getTextRanges(endingMarks: Array<string>, trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add to the document.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the paragraph object with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the paragraph.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in the paragraph.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: string, values?: Array<Array<string>>): Word.Table;
        /**
         *
         * Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start' or 'End'.
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects.
         *
         * [Api set: WordApi 1.1]
         *
         * @param searchText - Required. The search text.
         * @param searchOptions - Optional. Options for the search.
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects and navigates the Word UI to the paragraph.
         *
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        /**
         *
         * Splits the paragraph into child ranges by using delimiters.
         *
         * [Api set: WordApi 1.3]
         *
         * @param delimiters - Required. The delimiters as an array of strings.
         * @param trimDelimiters - Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        split(delimiters: Array<string>, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Starts a new list with this paragraph. Fails if the paragraph is already a list item.
         *
         * [Api set: WordApi 1.3]
         */
        startNewList(): Word.List;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Paragraph;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Paragraph;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Paragraph;
        toJSON(): {
            "alignment": string;
            "firstLineIndent": number;
            "font": Font;
            "isLastParagraph": boolean;
            "isListItem": boolean;
            "leftIndent": number;
            "lineSpacing": number;
            "lineUnitAfter": number;
            "lineUnitBefore": number;
            "listItem": ListItem;
            "listItemOrNullObject": ListItem;
            "outlineLevel": number;
            "rightIndent": number;
            "spaceAfter": number;
            "spaceBefore": number;
            "style": string;
            "styleBuiltIn": string;
            "tableNestingLevel": number;
            "text": string;
        };
    }
    /**
     *
     * Contains a collection of [paragraph](paragraph.md) objects.
     *
     * [Api set: WordApi 1.1]
     */
    export class ParagraphCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Paragraph>;
        /**
         *
         * Gets the first paragraph in this collection. Throws if the collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Paragraph;
        /**
         *
         * Gets the first paragraph in this collection. Returns a null object if the collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.Paragraph;
        /**
         *
         * Gets the last paragraph in this collection. Throws if the collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getLast(): Word.Paragraph;
        /**
         *
         * Gets the last paragraph in this collection. Returns a null object if the collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getLastOrNullObject(): Word.Paragraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.ParagraphCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ParagraphCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ParagraphCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a contiguous area in a document.
     *
     * [Api set: WordApi 1.1]
     */
    export class Range extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the collection of content control objects in the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        font: Word.Font;
        /**
         *
         * Gets the collection of inline picture objects in the range. Read-only.
         *
         * [Api set: WordApi 1.2]
         */
        inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of list objects in the range. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        lists: Word.ListCollection;
        /**
         *
         * Gets the collection of paragraph objects in the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the parent body of the range. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the range. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the range. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the table that contains the range. Throws if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the range. Throws if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCell: Word.TableCell;
        /**
         *
         * Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCellOrNullObject: Word.TableCell;
        /**
         *
         * Gets the table that contains the range. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableOrNullObject: Word.Table;
        /**
         *
         * Gets the collection of table objects in the range. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        tables: Word.TableCollection;
        /**
         *
         * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
         *
         * [Api set: WordApi 1.3]
         */
        hyperlink: string;
        /**
         *
         * Checks whether the range length is zero. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        isEmpty: boolean;
        /**
         *
         * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         *
         * Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: string;
        /**
         *
         * Gets the text of the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        text: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.RangeUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Range): void;
        /**
         *
         * Clears the contents of the range object. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         *
         * Compares this range's location with another range's location.
         *
         * [Api set: WordApi 1.3]
         *
         * @param range - Required. The range to compare with this range.
         */
        compareLocationWith(range: Word.Range): OfficeExtension.ClientResult<string>;
        /**
         *
         * Deletes the range and its content from the document.
         *
         * [Api set: WordApi 1.1]
         */
        delete(): void;
        /**
         *
         * Returns a new range that extends from this range in either direction to cover another range. This range is not changed. Throws if the two ranges do not have a union.
         *
         * [Api set: WordApi 1.3]
         *
         * @param range - Required. Another range.
         */
        expandTo(range: Word.Range): Word.Range;
        /**
         *
         * Returns a new range that extends from this range in either direction to cover another range. This range is not changed. Returns a null object if the two ranges do not have a union.
         *
         * [Api set: WordApi 1.3]
         *
         * @param range - Required. Another range.
         */
        expandToOrNullObject(range: Word.Range): Word.Range;
        /**
         *
         * Gets the HTML representation of the range object.
         *
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Gets hyperlink child ranges within the range.
         *
         * [Api set: WordApi 1.3]
         */
        getHyperlinkRanges(): Word.RangeCollection;
        /**
         *
         * Gets the next text range by using punctuation marks and/or other ending marks. Throws if this text range is the last one.
         *
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.
         */
        getNextTextRange(endingMarks: Array<string>, trimSpacing?: boolean): Word.Range;
        /**
         *
         * Gets the next text range by using punctuation marks and/or other ending marks. Returns a null object if this text range is the last one.
         *
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.
         */
        getNextTextRangeOrNullObject(endingMarks: Array<string>, trimSpacing?: boolean): Word.Range;
        /**
         *
         * Gets the OOXML representation of the range object.
         *
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Clones the range, or gets the starting or ending point of the range as a new range.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', 'After' or 'Content'.
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Gets the text child ranges in the range by using punctuation marks and/or other ending marks.
         *
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        getTextRanges(endingMarks: Array<string>, trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: string, insertLocation: string): void;
        /**
         *
         * Wraps the range object with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         */
        insertHtml(html: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: string, values?: Array<Array<string>>): Word.Table;
        /**
         *
         * Inserts text at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'.
         */
        insertText(text: string, insertLocation: string): Word.Range;
        /**
         *
         * Returns a new range as the intersection of this range with another range. This range is not changed. Throws if the two ranges are not overlapped or adjacent.
         *
         * [Api set: WordApi 1.3]
         *
         * @param range - Required. Another range.
         */
        intersectWith(range: Word.Range): Word.Range;
        /**
         *
         * Returns a new range as the intersection of this range with another range. This range is not changed. Returns a null object if the two ranges are not overlapped or adjacent.
         *
         * [Api set: WordApi 1.3]
         *
         * @param range - Required. Another range.
         */
        intersectWithOrNullObject(range: Word.Range): Word.Range;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects.
         *
         * [Api set: WordApi 1.1]
         *
         * @param searchText - Required. The search text.
         * @param searchOptions - Optional. Options for the search.
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects and navigates the Word UI to the range.
         *
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        /**
         *
         * Splits the range into child ranges by using delimiters.
         *
         * [Api set: WordApi 1.3]
         *
         * @param delimiters - Required. The delimiters as an array of strings.
         * @param multiParagraphs - Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.
         * @param trimDelimiters - Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        split(delimiters: Array<string>, multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Range;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Range;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Range;
        toJSON(): {
            "font": Font;
            "hyperlink": string;
            "isEmpty": boolean;
            "style": string;
            "styleBuiltIn": string;
            "text": string;
        };
    }
    /**
     *
     * Contains a collection of [range](range.md) objects.
     *
     * [Api set: WordApi 1.1]
     */
    export class RangeCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Range>;
        /**
         *
         * Gets the first range in this collection. Throws if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Range;
        /**
         *
         * Gets the first range in this collection. Returns a null object if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.RangeCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.RangeCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.RangeCollection;
        toJSON(): {};
    }
    /**
     *
     * Specifies the options to be included in a search operation.
     *
     * [Api set: WordApi 1.1]
     */
    export class SearchOptions extends OfficeExtension.ClientObject {
        matchWildCards: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        ignorePunct: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        ignoreSpace: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box (Edit menu).
         *
         * [Api set: WordApi 1.1]
         */
        matchCase: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        matchPrefix: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        matchSuffix: boolean;
        /**
         *
         * Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        matchWholeWord: boolean;
        /**
         *
         * Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
         *
         * [Api set: WordApi 1.1]
         */
        matchWildcards: boolean;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.SearchOptionsUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: SearchOptions): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.SearchOptions;
        /**
         * Create a new instance of Word.SearchOptions object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): Word.SearchOptions;
        toJSON(): {
            "ignorePunct": boolean;
            "ignoreSpace": boolean;
            "matchCase": boolean;
            "matchPrefix": boolean;
            "matchSuffix": boolean;
            "matchWholeWord": boolean;
            "matchWildcards": boolean;
        };
    }
    /**
     *
     * Represents a section in a Word document.
     *
     * [Api set: WordApi 1.1]
     */
    export class Section extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the body object of the section. This does not include the header/footer and other section metadata. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        body: Word.Body;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.SectionUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Section): void;
        /**
         *
         * Gets one of the section's footers.
         *
         * [Api set: WordApi 1.1]
         *
         * @param type - Required. The type of footer to return. This value can be: 'primary', 'firstPage' or 'evenPages'.
         */
        getFooter(type: string): Word.Body;
        /**
         *
         * Gets one of the section's headers.
         *
         * [Api set: WordApi 1.1]
         *
         * @param type - Required. The type of header to return. This value can be: 'primary', 'firstPage' or 'evenPages'.
         */
        getHeader(type: string): Word.Body;
        /**
         *
         * Gets the next section. Throws if this section is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.Section;
        /**
         *
         * Gets the next section. Returns a null object if this section is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Section;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Section;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Section;
        toJSON(): {
            "body": Body;
        };
    }
    /**
     *
     * Contains the collection of the document's [section](section.md) objects.
     *
     * [Api set: WordApi 1.1]
     */
    export class SectionCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Section>;
        /**
         *
         * Gets the first section in this collection. Throws if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Section;
        /**
         *
         * Gets the first section in this collection. Returns a null object if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.SectionCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.SectionCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.SectionCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a table in a Word document.
     *
     * [Api set: WordApi 1.3]
     */
    export class Table extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        font: Word.Font;
        /**
         *
         * Gets the parent body of the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the table. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the table. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the table that contains this table. Throws if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains this table. Throws if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCell: Word.TableCell;
        /**
         *
         * Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableCellOrNullObject: Word.TableCell;
        /**
         *
         * Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTableOrNullObject: Word.Table;
        /**
         *
         * Gets all of the table rows. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        rows: Word.TableRowCollection;
        /**
         *
         * Gets the child tables nested one level deeper. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        tables: Word.TableCollection;
        /**
         *
         * Gets or sets the alignment of the table against the page column. The value can be 'left', 'centered' or 'right'.
         *
         * [Api set: WordApi 1.3]
         */
        alignment: string;
        /**
         *
         * Gets and sets the number of header rows.
         *
         * [Api set: WordApi 1.3]
         */
        headerRowCount: number;
        /**
         *
         * Gets and sets the horizontal alignment of every cell in the table. The value can be 'left', 'centered', 'right', or 'justified'.
         *
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: string;
        /**
         *
         * Indicates whether all of the table rows are uniform. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        isUniform: boolean;
        /**
         *
         * Gets the nesting level of the table. Top-level tables have level 1. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        nestingLevel: number;
        /**
         *
         * Gets the number of rows in the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        rowCount: number;
        /**
         *
         * Gets and sets the shading color.
         *
         * [Api set: WordApi 1.3]
         */
        shadingColor: string;
        /**
         *
         * Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.3]
         */
        style: string;
        /**
         *
         * Gets and sets whether the table has banded columns.
         *
         * [Api set: WordApi 1.3]
         */
        styleBandedColumns: boolean;
        /**
         *
         * Gets and sets whether the table has banded rows.
         *
         * [Api set: WordApi 1.3]
         */
        styleBandedRows: boolean;
        /**
         *
         * Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: string;
        /**
         *
         * Gets and sets whether the table has a first column with a special style.
         *
         * [Api set: WordApi 1.3]
         */
        styleFirstColumn: boolean;
        /**
         *
         * Gets and sets whether the table has a last column with a special style.
         *
         * [Api set: WordApi 1.3]
         */
        styleLastColumn: boolean;
        /**
         *
         * Gets and sets whether the table has a total (last) row with a special style.
         *
         * [Api set: WordApi 1.3]
         */
        styleTotalRow: boolean;
        /**
         *
         * Gets and sets the text values in the table, as a 2D Javascript array.
         *
         * [Api set: WordApi 1.3]
         */
        values: Array<Array<string>>;
        /**
         *
         * Gets and sets the vertical alignment of every cell in the table. The value can be 'top', 'center' or 'bottom'.
         *
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: string;
        /**
         *
         * Gets and sets the width of the table in points.
         *
         * [Api set: WordApi 1.3]
         */
        width: number;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.TableUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Table): void;
        /**
         *
         * Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
         *
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. It can be 'Start' or 'End', corresponding to the appropriate side of the table.
         * @param columnCount - Required. Number of columns to add.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        addColumns(insertLocation: string, columnCount: number, values?: Array<Array<string>>): void;
        /**
         *
         * Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.
         *
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. It can be 'Start' or 'End'.
         * @param rowCount - Required. Number of rows to add.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        addRows(insertLocation: string, rowCount: number, values?: Array<Array<string>>): Word.TableRowCollection;
        /**
         *
         * Autofits the table columns to the width of the window.
         *
         * [Api set: WordApi 1.3]
         */
        autoFitWindow(): void;
        /**
         *
         * Clears the contents of the table.
         *
         * [Api set: WordApi 1.3]
         */
        clear(): void;
        /**
         *
         * Deletes the entire table.
         *
         * [Api set: WordApi 1.3]
         */
        delete(): void;
        /**
         *
         * Deletes specific columns. This is applicable to uniform tables.
         *
         * [Api set: WordApi 1.3]
         *
         * @param columnIndex - Required. The first column to delete.
         * @param columnCount - Optional. The number of columns to delete. Default 1.
         */
        deleteColumns(columnIndex: number, columnCount?: number): void;
        /**
         *
         * Deletes specific rows.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rowIndex - Required. The first row to delete.
         * @param rowCount - Optional. The number of rows to delete. Default 1.
         */
        deleteRows(rowIndex: number, rowCount?: number): void;
        /**
         *
         * Distributes the column widths evenly. This is applicable to uniform tables.
         *
         * [Api set: WordApi 1.3]
         */
        distributeColumns(): void;
        /**
         *
         * Gets the border style for the specified border.
         *
         * [Api set: WordApi 1.3]
         *
         * @param borderLocation - Required. The border location.
         */
        getBorder(borderLocation: string): Word.TableBorder;
        /**
         *
         * Gets the table cell at a specified row and column. Throws if the specified table cell does not exist.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rowIndex - Required. The index of the row.
         * @param cellIndex - Required. The index of the cell in the row.
         */
        getCell(rowIndex: number, cellIndex: number): Word.TableCell;
        /**
         *
         * Gets the table cell at a specified row and column. Returns a null object if the specified table cell does not exist.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rowIndex - Required. The index of the row.
         * @param cellIndex - Required. The index of the cell in the row.
         */
        getCellOrNullObject(rowIndex: number, cellIndex: number): Word.TableCell;
        /**
         *
         * Gets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         */
        getCellPadding(cellPaddingLocation: string): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets the next table. Throws if this table is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.Table;
        /**
         *
         * Gets the next table. Returns a null object if this table is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.Table;
        /**
         *
         * Gets the paragraph after the table. Throws if there isn't a paragraph after the table.
         *
         * [Api set: WordApi 1.3]
         */
        getParagraphAfter(): Word.Paragraph;
        /**
         *
         * Gets the paragraph after the table. Returns a null object if there isn't a paragraph after the table.
         *
         * [Api set: WordApi 1.3]
         */
        getParagraphAfterOrNullObject(): Word.Paragraph;
        /**
         *
         * Gets the paragraph before the table. Throws if there isn't a paragraph before the table.
         *
         * [Api set: WordApi 1.3]
         */
        getParagraphBefore(): Word.Paragraph;
        /**
         *
         * Gets the paragraph before the table. Returns a null object if there isn't a paragraph before the table.
         *
         * [Api set: WordApi 1.3]
         */
        getParagraphBeforeOrNullObject(): Word.Paragraph;
        /**
         *
         * Gets the range that contains this table, or the range at the start or end of the table.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End' or 'After'.
         */
        getRange(rangeLocation?: string): Word.Range;
        /**
         *
         * Inserts a content control on the table.
         *
         * [Api set: WordApi 1.3]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: string, values?: Array<Array<string>>): Word.Table;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects.
         *
         * [Api set: WordApi 1.3]
         *
         * @param searchText - Required. The search text.
         * @param searchOptions - Optional. Options for the search.
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
         *
         * [Api set: WordApi 1.3]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        /**
         *
         * Sets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: string, cellPadding: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.Table;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Table;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Table;
        toJSON(): {
            "alignment": string;
            "font": Font;
            "headerRowCount": number;
            "horizontalAlignment": string;
            "isUniform": boolean;
            "nestingLevel": number;
            "rowCount": number;
            "shadingColor": string;
            "style": string;
            "styleBandedColumns": boolean;
            "styleBandedRows": boolean;
            "styleBuiltIn": string;
            "styleFirstColumn": boolean;
            "styleLastColumn": boolean;
            "styleTotalRow": boolean;
            "values": string[][];
            "verticalAlignment": string;
            "width": number;
        };
    }
    /**
     *
     * Contains the collection of the document's Table objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Table>;
        /**
         *
         * Gets the first table in this collection. Throws if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Table;
        /**
         *
         * Gets the first table in this collection. Returns a null object if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a row in a Word document.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableRow extends OfficeExtension.ClientObject {
        /**
         *
         * Gets cells. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        cells: Word.TableCellCollection;
        /**
         *
         * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        font: Word.Font;
        /**
         *
         * Gets parent table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the number of cells in the row. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        cellCount: number;
        /**
         *
         * Gets and sets the horizontal alignment of every cell in the row. The value can be 'left', 'centered', 'right', or 'justified'.
         *
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: string;
        /**
         *
         * Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
         *
         * [Api set: WordApi 1.3]
         */
        isHeader: boolean;
        /**
         *
         * Gets and sets the preferred height of the row in points.
         *
         * [Api set: WordApi 1.3]
         */
        preferredHeight: number;
        /**
         *
         * Gets the index of the row in its parent table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        rowIndex: number;
        /**
         *
         * Gets and sets the shading color.
         *
         * [Api set: WordApi 1.3]
         */
        shadingColor: string;
        /**
         *
         * Gets and sets the text values in the row, as a 2D Javascript array.
         *
         * [Api set: WordApi 1.3]
         */
        values: Array<Array<string>>;
        /**
         *
         * Gets and sets the vertical alignment of the cells in the row. The value can be 'top', 'center' or 'bottom'.
         *
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.TableRowUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: TableRow): void;
        /**
         *
         * Clears the contents of the row.
         *
         * [Api set: WordApi 1.3]
         */
        clear(): void;
        /**
         *
         * Deletes the entire row.
         *
         * [Api set: WordApi 1.3]
         */
        delete(): void;
        /**
         *
         * Gets the border style of the cells in the row.
         *
         * [Api set: WordApi 1.3]
         *
         * @param borderLocation - Required. The border location.
         */
        getBorder(borderLocation: string): Word.TableBorder;
        /**
         *
         * Gets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         */
        getCellPadding(cellPaddingLocation: string): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets the next row. Throws if this row is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.TableRow;
        /**
         *
         * Gets the next row. Returns a null object if this row is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.TableRow;
        /**
         *
         * Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.
         *
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. Where the new rows should be inserted, relative to the current row. It can be 'Before' or 'After'.
         * @param rowCount - Required. Number of rows to add
         * @param values - Optional. Strings to insert in the new rows, specified as a 2D array. The number of cells in each row must not exceed the number of cells in the existing row.
         */
        insertRows(insertLocation: string, rowCount: number, values?: Array<Array<string>>): Word.TableRowCollection;
        /**
         *
         * Performs a search with the specified searchOptions on the scope of the row. The search results are a collection of range objects.
         *
         * [Api set: WordApi 1.3]
         *
         * @param searchText - Required. The search text.
         * @param searchOptions - Optional. Options for the search.
         */
        search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
        /**
         *
         * Selects the row and navigates the Word UI to it.
         *
         * [Api set: WordApi 1.3]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.
         */
        select(selectionMode?: string): void;
        /**
         *
         * Sets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: string, cellPadding: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableRow;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableRow;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableRow;
        toJSON(): {
            "cellCount": number;
            "font": Font;
            "horizontalAlignment": string;
            "isHeader": boolean;
            "preferredHeight": number;
            "rowIndex": number;
            "shadingColor": string;
            "values": string[][];
            "verticalAlignment": string;
        };
    }
    /**
     *
     * Contains the collection of the document's TableRow objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableRowCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        items: Array<Word.TableRow>;
        /**
         *
         * Gets the first row in this collection. Throws if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.TableRow;
        /**
         *
         * Gets the first row in this collection. Returns a null object if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableRowCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableRowCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableRowCollection;
        toJSON(): {};
    }
    /**
     *
     * Represents a table cell in a Word document.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableCell extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the body object of the cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        body: Word.Body;
        /**
         *
         * Gets the parent row of the cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentRow: Word.TableRow;
        /**
         *
         * Gets the parent table of the cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        parentTable: Word.Table;
        /**
         *
         * Gets the index of the cell in its row. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        cellIndex: number;
        /**
         *
         * Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
         *
         * [Api set: WordApi 1.3]
         */
        columnWidth: number;
        /**
         *
         * Gets and sets the horizontal alignment of the cell. The value can be 'left', 'centered', 'right', or 'justified'.
         *
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: string;
        /**
         *
         * Gets the index of the cell's row in the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        rowIndex: number;
        /**
         *
         * Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
         *
         * [Api set: WordApi 1.3]
         */
        shadingColor: string;
        /**
         *
         * Gets and sets the text of the cell.
         *
         * [Api set: WordApi 1.3]
         */
        value: string;
        /**
         *
         * Gets and sets the vertical alignment of the cell. The value can be 'top', 'center' or 'bottom'.
         *
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: string;
        /**
         *
         * Gets the width of the cell in points. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        width: number;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.TableCellUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: TableCell): void;
        /**
         *
         * Deletes the column containing this cell. This is applicable to uniform tables.
         *
         * [Api set: WordApi 1.3]
         */
        deleteColumn(): void;
        /**
         *
         * Deletes the row containing this cell.
         *
         * [Api set: WordApi 1.3]
         */
        deleteRow(): void;
        /**
         *
         * Gets the border style for the specified border.
         *
         * [Api set: WordApi 1.3]
         *
         * @param borderLocation - Required. The border location.
         */
        getBorder(borderLocation: string): Word.TableBorder;
        /**
         *
         * Gets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         */
        getCellPadding(cellPaddingLocation: string): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets the next cell. Throws if this cell is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.TableCell;
        /**
         *
         * Gets the next cell. Returns a null object if this cell is the last one.
         *
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.TableCell;
        /**
         *
         * Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
         *
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. It can be 'Before' or 'After'.
         * @param columnCount - Required. Number of columns to add
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertColumns(insertLocation: string, columnCount: number, values?: Array<Array<string>>): void;
        /**
         *
         * Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.
         *
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. It can be 'Before' or 'After'.
         * @param rowCount - Required. Number of rows to add.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertRows(insertLocation: string, rowCount: number, values?: Array<Array<string>>): Word.TableRowCollection;
        /**
         *
         * Sets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: string, cellPadding: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableCell;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableCell;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableCell;
        toJSON(): {
            "body": Body;
            "cellIndex": number;
            "columnWidth": number;
            "horizontalAlignment": string;
            "rowIndex": number;
            "shadingColor": string;
            "value": string;
            "verticalAlignment": string;
            "width": number;
        };
    }
    /**
     *
     * Contains the collection of the document's TableCell objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableCellCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        items: Array<Word.TableCell>;
        /**
         *
         * Gets the first table cell in this collection. Throws if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.TableCell;
        /**
         *
         * Gets the first table cell in this collection. Returns a null object if this collection is empty.
         *
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.TableCell;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableCellCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableCellCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableCellCollection;
        toJSON(): {};
    }
    /**
     *
     * Specifies the border style
     *
     * [Api set: WordApi 1.3]
     */
    export class TableBorder extends OfficeExtension.ClientObject {
        /**
         *
         * Gets or sets the table border color, as a hex value or name.
         *
         * [Api set: WordApi 1.3]
         */
        color: string;
        /**
         *
         * Gets or sets the type of the table border.
         *
         * [Api set: WordApi 1.3]
         */
        type: string;
        /**
         *
         * Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
         *
         * [Api set: WordApi 1.3]
         */
        width: number;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.TableBorderUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: TableBorder): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Word.TableBorder;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableBorder;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableBorder;
        toJSON(): {
            "color": string;
            "type": string;
            "width": number;
        };
    }
    /**
     *
     * Specifies supported content control types and subtypes.
     *
     * [Api set: WordApi]
     */
    export namespace ContentControlType {
        var unknown: string;
        var richTextInline: string;
        var richTextParagraphs: string;
        var richTextTableCell: string;
        var richTextTableRow: string;
        var richTextTable: string;
        var plainTextInline: string;
        var plainTextParagraph: string;
        var picture: string;
        var buildingBlockGallery: string;
        var checkBox: string;
        var comboBox: string;
        var dropDownList: string;
        var datePicker: string;
        var repeatingSection: string;
        var richText: string;
        var plainText: string;
    }
    /**
     *
     * ContentControl appearance
     *
     * [Api set: WordApi]
     */
    export namespace ContentControlAppearance {
        var boundingBox: string;
        var tags: string;
        var hidden: string;
    }
    /**
     *
     * Underline types
     *
     * [Api set: WordApi]
     */
    export namespace UnderlineType {
        var mixed: string;
        var none: string;
        /**
         *
         * @deprecated Hidden is no longer supported.
         */
        var hidden: string;
        /**
         *
         * @deprecated DotLine is no longer supported.
         */
        var dotLine: string;
        var single: string;
        var word: string;
        var double: string;
        var thick: string;
        var dotted: string;
        var dottedHeavy: string;
        var dashLine: string;
        var dashLineHeavy: string;
        var dashLineLong: string;
        var dashLineLongHeavy: string;
        var dotDashLine: string;
        var dotDashLineHeavy: string;
        var twoDotDashLine: string;
        var twoDotDashLineHeavy: string;
        var wave: string;
        var waveHeavy: string;
        var waveDouble: string;
    }
    /**
     *
     * Page break, line break, and four section breaks
     *
     * [Api set: WordApi]
     */
    export namespace BreakType {
        /**
         *
         * Page break.
         *
         */
        var page: string;
        /**
         *
         * @deprecated Use sectionNext instead.
         */
        var next: string;
        /**
         *
         * Section break, with the new section starting on the next page.
         *
         */
        var sectionNext: string;
        /**
         *
         * Section break, with the new section starting on the same page.
         *
         */
        var sectionContinuous: string;
        /**
         *
         * Section break, with the new section starting on the next even-numbered page.
         *
         */
        var sectionEven: string;
        /**
         *
         * Section break, with the new section starting on the next odd-numbered page.
         *
         */
        var sectionOdd: string;
        /**
         *
         * Line break.
         *
         */
        var line: string;
    }
    /**
     *
     * The insertion location types
     *
     * [Api set: WordApi]
     */
    export namespace InsertLocation {
        var before: string;
        var after: string;
        var start: string;
        var end: string;
        var replace: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace Alignment {
        var mixed: string;
        var unknown: string;
        var left: string;
        var centered: string;
        var right: string;
        var justified: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace HeaderFooterType {
        var primary: string;
        var firstPage: string;
        var evenPages: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace BodyType {
        var unknown: string;
        var mainDoc: string;
        var section: string;
        var header: string;
        var footer: string;
        var tableCell: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace SelectionMode {
        var select: string;
        var start: string;
        var end: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace ImageFormat {
        var unsupported: string;
        var undefined: string;
        var bmp: string;
        var jpeg: string;
        var gif: string;
        var tiff: string;
        var png: string;
        var icon: string;
        var exif: string;
        var wmf: string;
        var emf: string;
        var pict: string;
        var pdf: string;
        var svg: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace RangeLocation {
        var whole: string;
        var start: string;
        var end: string;
        var before: string;
        var after: string;
        var content: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace LocationRelation {
        var unrelated: string;
        var equal: string;
        var containsStart: string;
        var containsEnd: string;
        var contains: string;
        var insideStart: string;
        var insideEnd: string;
        var inside: string;
        var adjacentBefore: string;
        var overlapsBefore: string;
        var before: string;
        var adjacentAfter: string;
        var overlapsAfter: string;
        var after: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace BorderLocation {
        var top: string;
        var left: string;
        var bottom: string;
        var right: string;
        var insideHorizontal: string;
        var insideVertical: string;
        var inside: string;
        var outside: string;
        var all: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace CellPaddingLocation {
        var top: string;
        var left: string;
        var bottom: string;
        var right: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace BorderType {
        var mixed: string;
        var none: string;
        var single: string;
        var double: string;
        var dotted: string;
        var dashed: string;
        var dotDashed: string;
        var dot2Dashed: string;
        var triple: string;
        var thinThickSmall: string;
        var thickThinSmall: string;
        var thinThickThinSmall: string;
        var thinThickMed: string;
        var thickThinMed: string;
        var thinThickThinMed: string;
        var thinThickLarge: string;
        var thickThinLarge: string;
        var thinThickThinLarge: string;
        var wave: string;
        var doubleWave: string;
        var dashedSmall: string;
        var dashDotStroked: string;
        var threeDEmboss: string;
        var threeDEngrave: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace VerticalAlignment {
        var mixed: string;
        var top: string;
        var center: string;
        var bottom: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace ListLevelType {
        var bullet: string;
        var number: string;
        var picture: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace ListBullet {
        var custom: string;
        var solid: string;
        var hollow: string;
        var square: string;
        var diamonds: string;
        var arrow: string;
        var checkmark: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace ListNumbering {
        var none: string;
        var arabic: string;
        var upperRoman: string;
        var lowerRoman: string;
        var upperLetter: string;
        var lowerLetter: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace Style {
        /**
         *
         * Mixed styles or other style not in this list.
         *
         */
        var other: string;
        /**
         *
         * Reset character and paragraph style to default.
         *
         */
        var normal: string;
        var heading1: string;
        var heading2: string;
        var heading3: string;
        var heading4: string;
        var heading5: string;
        var heading6: string;
        var heading7: string;
        var heading8: string;
        var heading9: string;
        /**
         *
         * Table-of-content level 1.
         *
         */
        var toc1: string;
        /**
         *
         * Table-of-content level 2.
         *
         */
        var toc2: string;
        /**
         *
         * Table-of-content level 3.
         *
         */
        var toc3: string;
        /**
         *
         * Table-of-content level 4.
         *
         */
        var toc4: string;
        /**
         *
         * Table-of-content level 5.
         *
         */
        var toc5: string;
        /**
         *
         * Table-of-content level 6.
         *
         */
        var toc6: string;
        /**
         *
         * Table-of-content level 7.
         *
         */
        var toc7: string;
        /**
         *
         * Table-of-content level 8.
         *
         */
        var toc8: string;
        /**
         *
         * Table-of-content level 9.
         *
         */
        var toc9: string;
        var footnoteText: string;
        var header: string;
        var footer: string;
        var caption: string;
        var footnoteReference: string;
        var endnoteReference: string;
        var endnoteText: string;
        var title: string;
        var subtitle: string;
        var hyperlink: string;
        var strong: string;
        var emphasis: string;
        var noSpacing: string;
        var listParagraph: string;
        var quote: string;
        var intenseQuote: string;
        var subtleEmphasis: string;
        var intenseEmphasis: string;
        var subtleReference: string;
        var intenseReference: string;
        var bookTitle: string;
        var bibliography: string;
        /**
         *
         * Table-of-content heading.
         *
         */
        var tocHeading: string;
        var tableGrid: string;
        var plainTable1: string;
        var plainTable2: string;
        var plainTable3: string;
        var plainTable4: string;
        var plainTable5: string;
        var tableGridLight: string;
        var gridTable1Light: string;
        var gridTable1Light_Accent1: string;
        var gridTable1Light_Accent2: string;
        var gridTable1Light_Accent3: string;
        var gridTable1Light_Accent4: string;
        var gridTable1Light_Accent5: string;
        var gridTable1Light_Accent6: string;
        var gridTable2: string;
        var gridTable2_Accent1: string;
        var gridTable2_Accent2: string;
        var gridTable2_Accent3: string;
        var gridTable2_Accent4: string;
        var gridTable2_Accent5: string;
        var gridTable2_Accent6: string;
        var gridTable3: string;
        var gridTable3_Accent1: string;
        var gridTable3_Accent2: string;
        var gridTable3_Accent3: string;
        var gridTable3_Accent4: string;
        var gridTable3_Accent5: string;
        var gridTable3_Accent6: string;
        var gridTable4: string;
        var gridTable4_Accent1: string;
        var gridTable4_Accent2: string;
        var gridTable4_Accent3: string;
        var gridTable4_Accent4: string;
        var gridTable4_Accent5: string;
        var gridTable4_Accent6: string;
        var gridTable5Dark: string;
        var gridTable5Dark_Accent1: string;
        var gridTable5Dark_Accent2: string;
        var gridTable5Dark_Accent3: string;
        var gridTable5Dark_Accent4: string;
        var gridTable5Dark_Accent5: string;
        var gridTable5Dark_Accent6: string;
        var gridTable6Colorful: string;
        var gridTable6Colorful_Accent1: string;
        var gridTable6Colorful_Accent2: string;
        var gridTable6Colorful_Accent3: string;
        var gridTable6Colorful_Accent4: string;
        var gridTable6Colorful_Accent5: string;
        var gridTable6Colorful_Accent6: string;
        var gridTable7Colorful: string;
        var gridTable7Colorful_Accent1: string;
        var gridTable7Colorful_Accent2: string;
        var gridTable7Colorful_Accent3: string;
        var gridTable7Colorful_Accent4: string;
        var gridTable7Colorful_Accent5: string;
        var gridTable7Colorful_Accent6: string;
        var listTable1Light: string;
        var listTable1Light_Accent1: string;
        var listTable1Light_Accent2: string;
        var listTable1Light_Accent3: string;
        var listTable1Light_Accent4: string;
        var listTable1Light_Accent5: string;
        var listTable1Light_Accent6: string;
        var listTable2: string;
        var listTable2_Accent1: string;
        var listTable2_Accent2: string;
        var listTable2_Accent3: string;
        var listTable2_Accent4: string;
        var listTable2_Accent5: string;
        var listTable2_Accent6: string;
        var listTable3: string;
        var listTable3_Accent1: string;
        var listTable3_Accent2: string;
        var listTable3_Accent3: string;
        var listTable3_Accent4: string;
        var listTable3_Accent5: string;
        var listTable3_Accent6: string;
        var listTable4: string;
        var listTable4_Accent1: string;
        var listTable4_Accent2: string;
        var listTable4_Accent3: string;
        var listTable4_Accent4: string;
        var listTable4_Accent5: string;
        var listTable4_Accent6: string;
        var listTable5Dark: string;
        var listTable5Dark_Accent1: string;
        var listTable5Dark_Accent2: string;
        var listTable5Dark_Accent3: string;
        var listTable5Dark_Accent4: string;
        var listTable5Dark_Accent5: string;
        var listTable5Dark_Accent6: string;
        var listTable6Colorful: string;
        var listTable6Colorful_Accent1: string;
        var listTable6Colorful_Accent2: string;
        var listTable6Colorful_Accent3: string;
        var listTable6Colorful_Accent4: string;
        var listTable6Colorful_Accent5: string;
        var listTable6Colorful_Accent6: string;
        var listTable7Colorful: string;
        var listTable7Colorful_Accent1: string;
        var listTable7Colorful_Accent2: string;
        var listTable7Colorful_Accent3: string;
        var listTable7Colorful_Accent4: string;
        var listTable7Colorful_Accent5: string;
        var listTable7Colorful_Accent6: string;
    }
    /**
     * [Api set: WordApi]
     */
    export namespace DocumentPropertyType {
        var string: string;
        var number: string;
        var date: string;
        var boolean: string;
    }
    export namespace ErrorCodes {
        var accessDenied: string;
        var generalException: string;
        var invalidArgument: string;
        var itemNotFound: string;
        var notImplemented: string;
    }
    export module Interfaces {
        /** An interface for updating data on the Body object, for use in "body.set({ ... })". */
        export interface BodyUpdateData {
            /**
            *
            * Gets the text format of the body. Use this to get and set font name, size, color and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             *
             * Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             *
             * Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: string;
        }
        /** An interface for updating data on the ContentControl object, for use in "contentControl.set({ ... })". */
        export interface ContentControlUpdateData {
            /**
            *
            * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             *
             * Gets or sets the appearance of the content control. The value can be 'boundingBox', 'tags' or 'hidden'.
             *
             * [Api set: WordApi 1.1]
             */
            appearance?: string;
            /**
             *
             * Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
             *
             * [Api set: WordApi 1.1]
             */
            cannotDelete?: boolean;
            /**
             *
             * Gets or sets a value that indicates whether the user can edit the contents of the content control.
             *
             * [Api set: WordApi 1.1]
             */
            cannotEdit?: boolean;
            /**
             *
             * Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
             *
             * [Api set: WordApi 1.1]
             */
            color?: string;
            /**
             *
             * Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             *
             * [Api set: WordApi 1.1]
             */
            placeholderText?: string;
            /**
             *
             * Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
             *
             * [Api set: WordApi 1.1]
             */
            removeWhenEdited?: boolean;
            /**
             *
             * Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             *
             * Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: string;
            /**
             *
             * Gets or sets a tag to identify a content control.
             *
             * [Api set: WordApi 1.1]
             */
            tag?: string;
            /**
             *
             * Gets or sets the title for a content control.
             *
             * [Api set: WordApi 1.1]
             */
            title?: string;
        }
        /** An interface for updating data on the CustomProperty object, for use in "customProperty.set({ ... })". */
        export interface CustomPropertyUpdateData {
            /**
             *
             * Gets or sets the value of the custom property.
             *
             * [Api set: WordApi 1.3]
             */
            value?: any;
        }
        /** An interface for updating data on the Document object, for use in "document.set({ ... })". */
        export interface DocumentUpdateData {
            /**
            *
            * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc..
            *
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyUpdateData;
            /**
            *
            * Gets the properties of the current document.
            *
            * [Api set: WordApi 1.3]
            */
            properties?: Word.Interfaces.DocumentPropertiesUpdateData;
        }
        /** An interface for updating data on the DocumentProperties object, for use in "documentProperties.set({ ... })". */
        export interface DocumentPropertiesUpdateData {
            /**
             *
             * Gets or sets the author of the document.
             *
             * [Api set: WordApi 1.3]
             */
            author?: string;
            /**
             *
             * Gets or sets the category of the document.
             *
             * [Api set: WordApi 1.3]
             */
            category?: string;
            /**
             *
             * Gets or sets the comments of the document.
             *
             * [Api set: WordApi 1.3]
             */
            comments?: string;
            /**
             *
             * Gets or sets the company of the document.
             *
             * [Api set: WordApi 1.3]
             */
            company?: string;
            /**
             *
             * Gets or sets the format of the document.
             *
             * [Api set: WordApi 1.3]
             */
            format?: string;
            /**
             *
             * Gets or sets the keywords of the document.
             *
             * [Api set: WordApi 1.3]
             */
            keywords?: string;
            /**
             *
             * Gets or sets the manager of the document.
             *
             * [Api set: WordApi 1.3]
             */
            manager?: string;
            /**
             *
             * Gets or sets the subject of the document.
             *
             * [Api set: WordApi 1.3]
             */
            subject?: string;
            /**
             *
             * Gets or sets the title of the document.
             *
             * [Api set: WordApi 1.3]
             */
            title?: string;
        }
        /** An interface for updating data on the Font object, for use in "font.set({ ... })". */
        export interface FontUpdateData {
            /**
             *
             * Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
             *
             * [Api set: WordApi 1.1]
             */
            bold?: boolean;
            /**
             *
             * Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
             *
             * [Api set: WordApi 1.1]
             */
            color?: string;
            /**
             *
             * Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false.
             *
             * [Api set: WordApi 1.1]
             */
            doubleStrikeThrough?: boolean;
            /**
             *
             * Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, or an empty string for mixed highlight colors, or null for no highlight color.
             *
             * [Api set: WordApi 1.1]
             */
            highlightColor?: string;
            /**
             *
             * Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
             *
             * [Api set: WordApi 1.1]
             */
            italic?: boolean;
            /**
             *
             * Gets or sets a value that represents the name of the font.
             *
             * [Api set: WordApi 1.1]
             */
            name?: string;
            /**
             *
             * Gets or sets a value that represents the font size in points.
             *
             * [Api set: WordApi 1.1]
             */
            size?: number;
            /**
             *
             * Gets or sets a value that indicates whether the font has a strike through. True if the font is formatted as strikethrough text, otherwise, false.
             *
             * [Api set: WordApi 1.1]
             */
            strikeThrough?: boolean;
            /**
             *
             * Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
             *
             * [Api set: WordApi 1.1]
             */
            subscript?: boolean;
            /**
             *
             * Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
             *
             * [Api set: WordApi 1.1]
             */
            superscript?: boolean;
            /**
             *
             * Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
             *
             * [Api set: WordApi 1.1]
             */
            underline?: string;
        }
        /** An interface for updating data on the InlinePicture object, for use in "inlinePicture.set({ ... })". */
        export interface InlinePictureUpdateData {
            /**
             *
             * Gets or sets a string that represents the alternative text associated with the inline image
             *
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: string;
            /**
             *
             * Gets or sets a string that contains the title for the inline image.
             *
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: string;
            /**
             *
             * Gets or sets a number that describes the height of the inline image.
             *
             * [Api set: WordApi 1.1]
             */
            height?: number;
            /**
             *
             * Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * [Api set: WordApi 1.1]
             */
            hyperlink?: string;
            /**
             *
             * Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
             *
             * [Api set: WordApi 1.1]
             */
            lockAspectRatio?: boolean;
            /**
             *
             * Gets or sets a number that describes the width of the inline image.
             *
             * [Api set: WordApi 1.1]
             */
            width?: number;
        }
        /** An interface for updating data on the ListItem object, for use in "listItem.set({ ... })". */
        export interface ListItemUpdateData {
            /**
             *
             * Gets or sets the level of the item in the list.
             *
             * [Api set: WordApi 1.3]
             */
            level?: number;
        }
        /** An interface for updating data on the Paragraph object, for use in "paragraph.set({ ... })". */
        export interface ParagraphUpdateData {
            /**
            *
            * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
            *
            * Gets the ListItem for the paragraph. Throws if the paragraph is not part of a list.
            *
            * [Api set: WordApi 1.3]
            */
            listItem?: Word.Interfaces.ListItemUpdateData;
            /**
            *
            * Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list.
            *
            * [Api set: WordApi 1.3]
            */
            listItemOrNullObject?: Word.Interfaces.ListItemUpdateData;
            /**
             *
             * Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * [Api set: WordApi 1.1]
             */
            alignment?: string;
            /**
             *
             * Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: number;
            /**
             *
             * Gets or sets the left indent value, in points, for the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            leftIndent?: number;
            /**
             *
             * Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: number;
            /**
             *
             * Gets or sets the amount of spacing, in grid lines. after the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: number;
            /**
             *
             * Gets or sets the amount of spacing, in grid lines, before the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: number;
            /**
             *
             * Gets or sets the outline level for the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: number;
            /**
             *
             * Gets or sets the right indent value, in points, for the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            rightIndent?: number;
            /**
             *
             * Gets or sets the spacing, in points, after the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: number;
            /**
             *
             * Gets or sets the spacing, in points, before the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: number;
            /**
             *
             * Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             *
             * Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: string;
        }
        /** An interface for updating data on the Range object, for use in "range.set({ ... })". */
        export interface RangeUpdateData {
            /**
            *
            * Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             *
             * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
             *
             * [Api set: WordApi 1.3]
             */
            hyperlink?: string;
            /**
             *
             * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             *
             * Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: string;
        }
        /** An interface for updating data on the SearchOptions object, for use in "searchOptions.set({ ... })". */
        export interface SearchOptionsUpdateData {
            /**
             *
             * Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
             *
             * [Api set: WordApi 1.1]
             */
            ignorePunct?: boolean;
            /**
             *
             * Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
             *
             * [Api set: WordApi 1.1]
             */
            ignoreSpace?: boolean;
            /**
             *
             * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box (Edit menu).
             *
             * [Api set: WordApi 1.1]
             */
            matchCase?: boolean;
            /**
             *
             * Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
             *
             * [Api set: WordApi 1.1]
             */
            matchPrefix?: boolean;
            /**
             *
             * Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
             *
             * [Api set: WordApi 1.1]
             */
            matchSuffix?: boolean;
            /**
             *
             * Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
             *
             * [Api set: WordApi 1.1]
             */
            matchWholeWord?: boolean;
            /**
             *
             * Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
             *
             * [Api set: WordApi 1.1]
             */
            matchWildcards?: boolean;
        }
        /** An interface for updating data on the Section object, for use in "section.set({ ... })". */
        export interface SectionUpdateData {
            /**
            *
            * Gets the body object of the section. This does not include the header/footer and other section metadata.
            *
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyUpdateData;
        }
        /** An interface for updating data on the Table object, for use in "table.set({ ... })". */
        export interface TableUpdateData {
            /**
            *
            * Gets the font. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             *
             * Gets or sets the alignment of the table against the page column. The value can be 'left', 'centered' or 'right'.
             *
             * [Api set: WordApi 1.3]
             */
            alignment?: string;
            /**
             *
             * Gets and sets the number of header rows.
             *
             * [Api set: WordApi 1.3]
             */
            headerRowCount?: number;
            /**
             *
             * Gets and sets the horizontal alignment of every cell in the table. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: string;
            /**
             *
             * Gets and sets the shading color.
             *
             * [Api set: WordApi 1.3]
             */
            shadingColor?: string;
            /**
             *
             * Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.3]
             */
            style?: string;
            /**
             *
             * Gets and sets whether the table has banded columns.
             *
             * [Api set: WordApi 1.3]
             */
            styleBandedColumns?: boolean;
            /**
             *
             * Gets and sets whether the table has banded rows.
             *
             * [Api set: WordApi 1.3]
             */
            styleBandedRows?: boolean;
            /**
             *
             * Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: string;
            /**
             *
             * Gets and sets whether the table has a first column with a special style.
             *
             * [Api set: WordApi 1.3]
             */
            styleFirstColumn?: boolean;
            /**
             *
             * Gets and sets whether the table has a last column with a special style.
             *
             * [Api set: WordApi 1.3]
             */
            styleLastColumn?: boolean;
            /**
             *
             * Gets and sets whether the table has a total (last) row with a special style.
             *
             * [Api set: WordApi 1.3]
             */
            styleTotalRow?: boolean;
            /**
             *
             * Gets and sets the text values in the table, as a 2D Javascript array.
             *
             * [Api set: WordApi 1.3]
             */
            values?: Array<Array<string>>;
            /**
             *
             * Gets and sets the vertical alignment of every cell in the table. The value can be 'top', 'center' or 'bottom'.
             *
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: string;
            /**
             *
             * Gets and sets the width of the table in points.
             *
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /** An interface for updating data on the TableRow object, for use in "tableRow.set({ ... })". */
        export interface TableRowUpdateData {
            /**
            *
            * Gets the font. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             *
             * Gets and sets the horizontal alignment of every cell in the row. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: string;
            /**
             *
             * Gets and sets the preferred height of the row in points.
             *
             * [Api set: WordApi 1.3]
             */
            preferredHeight?: number;
            /**
             *
             * Gets and sets the shading color.
             *
             * [Api set: WordApi 1.3]
             */
            shadingColor?: string;
            /**
             *
             * Gets and sets the text values in the row, as a 2D Javascript array.
             *
             * [Api set: WordApi 1.3]
             */
            values?: Array<Array<string>>;
            /**
             *
             * Gets and sets the vertical alignment of the cells in the row. The value can be 'top', 'center' or 'bottom'.
             *
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: string;
        }
        /** An interface for updating data on the TableCell object, for use in "tableCell.set({ ... })". */
        export interface TableCellUpdateData {
            /**
            *
            * Gets the body object of the cell.
            *
            * [Api set: WordApi 1.3]
            */
            body?: Word.Interfaces.BodyUpdateData;
            /**
             *
             * Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
             *
             * [Api set: WordApi 1.3]
             */
            columnWidth?: number;
            /**
             *
             * Gets and sets the horizontal alignment of the cell. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: string;
            /**
             *
             * Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * [Api set: WordApi 1.3]
             */
            shadingColor?: string;
            /**
             *
             * Gets and sets the text of the cell.
             *
             * [Api set: WordApi 1.3]
             */
            value?: string;
            /**
             *
             * Gets and sets the vertical alignment of the cell. The value can be 'top', 'center' or 'bottom'.
             *
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: string;
        }
        /** An interface for updating data on the TableBorder object, for use in "tableBorder.set({ ... })". */
        export interface TableBorderUpdateData {
            /**
             *
             * Gets or sets the table border color, as a hex value or name.
             *
             * [Api set: WordApi 1.3]
             */
            color?: string;
            /**
             *
             * Gets or sets the type of the table border.
             *
             * [Api set: WordApi 1.3]
             */
            type?: string;
            /**
             *
             * Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
             *
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
    }
}
export declare module Word {
    /**
     * The RequestContext object facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the request context is required to get access to the Word object model from the add-in.
     */
    export class RequestContext extends OfficeExtension.ClientRequestContext {
        constructor(url?: string);
        document: Document;
    }
    /**
     * Executes a batch script that performs actions on the Word object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    export function run<T>(batch: (context: Word.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the Word object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    export function run<T>(object: OfficeExtension.ClientObject, batch: (context: Word.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the Word object model, using the RequestContext of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Word.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
}


////////////////////////////////////////////////////////////////
//////////////////////// End Word APIs /////////////////////////
////////////////////////////////////////////////////////////////
