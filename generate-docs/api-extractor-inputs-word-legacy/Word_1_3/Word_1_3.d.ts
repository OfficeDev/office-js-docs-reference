////////////////////////////////////////////////////////////////
/////////////////////// Begin Word APIs ////////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Word {
    /**
     *
     * Represents the application object.
     *
     * [Api set: WordApi 1.3]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Creates a new document by using an optional base64 encoded .docx file.
         *
         * [Api set: WordApi 1.3]
         *
         * @param base64File - Optional. The base64 encoded .docx file. The default value is null.
         */
        createDocument(base64File?: string): Word.DocumentCreated;
        /**
         * Create a new instance of Word.Application object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): Word.Application;
        toJSON(): {
            [key: string]: string;
        };
    }
    /**
     *
     * Represents the body of a document or a section.
     *
     * [Api set: WordApi 1.1]
     */
    export class Body extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the collection of rich text content control objects in the body. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        /**
         *
         * Gets the collection of InlinePicture objects in the body. The collection does not include floating images. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of list objects in the body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly lists: Word.ListCollection;
        /**
         *
         * Gets the collection of paragraph objects in the body. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws if there isn't a parent body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentBody: Word.Body;
        /**
         *
         * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Returns a null object if there isn't a parent body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentBodyOrNullObject: Word.Body;
        /**
         *
         * Gets the content control that contains the body. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the body. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the parent section of the body. Throws if there isn't a parent section. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentSection: Word.Section;
        /**
         *
         * Gets the parent section of the body. Returns a null object if there isn't a parent section. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentSectionOrNullObject: Word.Section;
        /**
         *
         * Gets the collection of table objects in the body. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly tables: Word.TableCollection;
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
        styleBuiltIn: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        /**
         *
         * Gets the text of the body. Use the insertText method to insert text. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        /**
         *
         * Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly type: Word.BodyType | "Unknown" | "MainDoc" | "Section" | "Header" | "Footer" | "TableCell";
        
        /**
         *
         * Clears the contents of the body object. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         *
         * Gets an HTML representation of the body object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word Online, etc.). If you need exact fidelity, or consistency across platforms, use `Body.getOoxml()` and convert the returned XML to HTML.
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
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: Word.RangeLocation): Word.Range;
        /**
         *
         * Gets the whole body, or the starting or ending point of the body, as a range.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: "Whole" | "Start" | "End" | "Before" | "After" | "Content"): Word.Range;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add to the body.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation): void;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add to the body.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertBreak(breakType: "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): void;
        /**
         *
         * Wraps the body object with a Rich Text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the document.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the document.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertHtml(html: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the body.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation): Word.InlinePicture;
        /**
         *
         * Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the body.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertParagraph(paragraphText: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
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
        insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][]): Word.Table;
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
        insertTable(rowCount: number, columnCount: number, insertLocation: "Before" | "After" | "Start" | "End" | "Replace", values?: string[][]): Word.Table;
        /**
         *
         * Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertText(text: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Performs a search with the specified SearchOptions on the scope of the body object. The search results are a collection of range objects.
         *
         * [Api set: WordApi 1.1]
         *
         * @param searchText - Required. The search text. Can be a maximum of 255 characters.
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
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         *
         * Selects the body and navigates the Word UI to it.
         *
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Body` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Body` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.Body;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.Body;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Body;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Body;
        toJSON(): Word.Interfaces.BodyData;
    }
    /**
     *
     * Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
     *
     * [Api set: WordApi 1.1]
     */
    export class ContentControl extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the collection of content control objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        /**
         *
         * Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of list objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly lists: Word.ListCollection;
        /**
         *
         * Get the collection of paragraph objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the parent body of the content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the content control. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the content control. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the table that contains the content control. Throws if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the content control. Throws if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCell: Word.TableCell;
        /**
         *
         * Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         *
         * Gets the table that contains the content control. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         *
         * Gets the collection of table objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly tables: Word.TableCollection;
        /**
         *
         * Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
         *
         * [Api set: WordApi 1.1]
         */
        appearance: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
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
        readonly id: number;
        /**
         *
         * Gets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
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
        styleBuiltIn: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        /**
         *
         * Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly subtype: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText";
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
        readonly text: string;
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
        readonly type: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText";
        
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
         * Gets an HTML representation of the content control object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word Online, etc.). If you need exact fidelity, or consistency across platforms, use `ContentControl.getOoxml()` and convert the returned XML to HTML.
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
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Before', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: Word.RangeLocation): Word.Range;
        /**
         *
         * Gets the whole content control, or the starting or ending point of the content control, as a range.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Before', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: "Whole" | "Start" | "End" | "Before" | "After" | "Content"): Word.Range;
        /**
         *
         * Gets the text ranges in the content control by using punctuation marks and/or other ending marks.
         *
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        getTextRanges(endingMarks: string[], trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before', or 'After'. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. Type of break.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'.
         */
        insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation): void;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before', or 'After'. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. Type of break.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'.
         */
        insertBreak(breakType: "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): void;
        /**
         *
         * Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertFileFromBase64(base64File: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertHtml(html: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation): Word.InlinePicture;
        /**
         *
         * Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.InlinePicture;
        /**
         *
         * Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertOoxml(ooxml: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before', or 'After'. This method is only supported if the content control encompasses one or more paragraphs in entirety.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before', or 'After'. This method is only supported if the content control encompasses one or more paragraphs in entirety.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         */
        insertParagraph(paragraphText: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
        /**
         *
         * Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][]): Word.Table;
        /**
         *
         * Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: "Before" | "After" | "Start" | "End" | "Replace", values?: string[][]): Word.Table;
        /**
         *
         * Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. The text to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertText(text: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. The text to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertText(text: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Performs a search with the specified SearchOptions on the scope of the content control object. The search results are a collection of range objects.
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
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         *
         * Selects the content control. This causes Word to scroll to the selection.
         *
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
        /**
         *
         * Splits the content control into child ranges by using delimiters.
         *
         * [Api set: WordApi 1.3]
         *
         * @param delimiters - Required. The delimiters as an array of strings.
         * @param multiParagraphs - Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.
         * @param trimDelimiters - Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.ContentControl` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.ContentControl` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.ContentControl;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.ContentControl;
        
        
        
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ContentControl;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ContentControl;
        toJSON(): Word.Interfaces.ContentControlData;
    }
    /**
     *
     * Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
     *
     * [Api set: WordApi 1.1]
     */
    export class ContentControlCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Word.ContentControl[];
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
        getByTypes(types: Word.ContentControlType[]): Word.ContentControlCollection;
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.ContentControlCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.ContentControlCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.ContentControlCollection;
        load(option?: OfficeExtension.LoadOption): Word.ContentControlCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ContentControlCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ContentControlCollection;
        toJSON(): Word.Interfaces.ContentControlCollectionData;
    }
    /**
     *
     * Represents a custom property.
     *
     * [Api set: WordApi 1.3]
     */
    export class CustomProperty extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the key of the custom property. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly key: string;
        /**
         *
         * Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly type: Word.DocumentPropertyType | "String" | "Number" | "Date" | "Boolean";
        /**
         *
         * Gets or sets the value of the custom property. Note that even though Word Online and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
         *
         * [Api set: WordApi 1.3]
         */
        value: any;
        
        /**
         *
         * Deletes the custom property.
         *
         * [Api set: WordApi 1.3]
         */
        delete(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.CustomProperty` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.CustomProperty` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.CustomProperty;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.CustomProperty;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.CustomProperty;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.CustomProperty;
        toJSON(): Word.Interfaces.CustomPropertyData;
    }
    /**
     *
     * Contains the collection of {@link Word.CustomProperty} objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class CustomPropertyCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Word.CustomProperty[];
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.CustomPropertyCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.CustomPropertyCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.CustomPropertyCollection;
        load(option?: OfficeExtension.LoadOption): Word.CustomPropertyCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.CustomPropertyCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.CustomPropertyCollection;
        toJSON(): Word.Interfaces.CustomPropertyCollectionData;
    }
    
    
    
    /**
     *
     * The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
     *
     * [Api set: WordApi 1.1]
     */
    export class Document extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly body: Word.Body;
        /**
         *
         * Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        
        /**
         *
         * Gets the properties of the document. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly properties: Word.DocumentProperties;
        /**
         *
         * Gets the collection of section objects in the document. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly sections: Word.SectionCollection;
        
        /**
         *
         * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly saved: boolean;
        
        
        
        
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Document` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Document` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.Document;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.Document;
        
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Document;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Document;
        toJSON(): Word.Interfaces.DocumentData;
    }
    /**
     *
     * The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.
     *
     * [Api set: WordApi 1.3]
     */
    export class DocumentCreated extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
         *
         * [Api set: WordApiHiddenDocument 1.3]
         */
        readonly body: Word.Body;
        /**
         *
         * Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
         *
         * [Api set: WordApiHiddenDocument 1.3]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the custom XML parts in the document. Read-only.
         *
         * [Api set: WordApiHiddenDocument 1.4]
         */
        readonly customXmlParts: Word.CustomXmlPartCollection;
        /**
         *
         * Gets the properties of the document. Read-only.
         *
         * [Api set: WordApiHiddenDocument 1.3]
         */
        readonly properties: Word.DocumentProperties;
        /**
         *
         * Gets the collection of section objects in the document. Read-only.
         *
         * [Api set: WordApiHiddenDocument 1.3]
         */
        readonly sections: Word.SectionCollection;
        /**
         *
         * Gets the add-in's settings in the document. Read-only.
         *
         * [Api set: WordApiHiddenDocument 1.4]
         */
        readonly settings: Word.SettingCollection;
        /**
         *
         * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
         *
         * [Api set: WordApiHiddenDocument 1.3]
         */
        readonly saved: boolean;
        
        /**
         *
         * Deletes a bookmark, if exists, from the document.
         *
         * [Api set: WordApiHiddenDocument 1.4]
         *
         * @param name - Required. The bookmark name, which is case-insensitive.
         */
        deleteBookmark(name: string): void;
        /**
         *
         * Gets a bookmark's range. Throws if the bookmark does not exist.
         *
         * [Api set: WordApiHiddenDocument 1.4]
         *
         * @param name - Required. The bookmark name, which is case-insensitive.
         */
        getBookmarkRange(name: string): Word.Range;
        /**
         *
         * Gets a bookmark's range. Returns a null object if the bookmark does not exist.
         *
         * [Api set: WordApiHiddenDocument 1.4]
         *
         * @param name - Required. The bookmark name, which is case-insensitive.
         */
        getBookmarkRangeOrNullObject(name: string): Word.Range;
        /**
         *
         * Opens the document.
         *
         * [Api set: WordApi 1.3]
         */
        open(): void;
        /**
         *
         * Saves the document. This will use the Word default file naming convention if the document has not been saved before.
         *
         * [Api set: WordApiHiddenDocument 1.3]
         */
        save(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.DocumentCreated` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.DocumentCreated` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.DocumentCreated;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.DocumentCreated;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.DocumentCreated;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.DocumentCreated;
        toJSON(): Word.Interfaces.DocumentCreatedData;
    }
    /**
     *
     * Represents document properties.
     *
     * [Api set: WordApi 1.3]
     */
    export class DocumentProperties extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the collection of custom properties of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly customProperties: Word.CustomPropertyCollection;
        /**
         *
         * Gets the application name of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly applicationName: string;
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
        readonly creationDate: Date;
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
        readonly lastAuthor: string;
        /**
         *
         * Gets the last print date of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly lastPrintDate: Date;
        /**
         *
         * Gets the last save time of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly lastSaveTime: Date;
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
        readonly revisionNumber: string;
        /**
         *
         * Gets the security of the document. Read only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly security: number;
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
        readonly template: string;
        /**
         *
         * Gets or sets the title of the document.
         *
         * [Api set: WordApi 1.3]
         */
        title: string;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.DocumentProperties` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.DocumentProperties` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.DocumentProperties;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.DocumentProperties;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.DocumentProperties;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.DocumentProperties;
        toJSON(): Word.Interfaces.DocumentPropertiesData;
    }
    /**
     *
     * Represents a font.
     *
     * [Api set: WordApi 1.1]
     */
    export class Font extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
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
         * Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
         *
         * [Api set: WordApi 1.1]
         */
        doubleStrikeThrough: boolean;
        /**
         *
         * Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
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
         * Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
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
        underline: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
        
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Font` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Font` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.Font;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.Font;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Font;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Font;
        toJSON(): Word.Interfaces.FontData;
    }
    /**
     *
     * Represents an inline picture.
     *
     * [Api set: WordApi 1.1]
     */
    export class InlinePicture extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the parent paragraph that contains the inline image. Read-only.
         *
         * [Api set: WordApi 1.2]
         */
        readonly paragraph: Word.Paragraph;
        /**
         *
         * Gets the content control that contains the inline image. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the table that contains the inline image. Throws if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the inline image. Throws if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCell: Word.TableCell;
        /**
         *
         * Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         *
         * Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         *
         * Gets or sets a string that represents the alternative text associated with the inline image.
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
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', or 'End'.
         */
        getRange(rangeLocation?: Word.RangeLocation): Word.Range;
        /**
         *
         * Gets the picture, or the starting or ending point of the picture, as a range.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', or 'End'.
         */
        getRange(rangeLocation?: "Whole" | "Start" | "End" | "Before" | "After" | "Content"): Word.Range;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param breakType - Required. The break type to add.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation): void;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param breakType - Required. The break type to add.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): void;
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
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertHtml(html: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Before', or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation): Word.InlinePicture;
        /**
         *
         * Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Before', or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
        /**
         *
         * Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertText(text: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Selects the inline picture. This causes Word to scroll to the selection.
         *
         * [Api set: WordApi 1.2]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         *
         * Selects the inline picture. This causes Word to scroll to the selection.
         *
         * [Api set: WordApi 1.2]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.InlinePicture` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.InlinePicture` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.InlinePicture;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.InlinePicture;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.InlinePicture;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.InlinePicture;
        toJSON(): Word.Interfaces.InlinePictureData;
    }
    /**
     *
     * Contains a collection of {@link Word.InlinePicture} objects.
     *
     * [Api set: WordApi 1.1]
     */
    export class InlinePictureCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Word.InlinePicture[];
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.InlinePictureCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.InlinePictureCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.InlinePictureCollection;
        load(option?: OfficeExtension.LoadOption): Word.InlinePictureCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.InlinePictureCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.InlinePictureCollection;
        toJSON(): Word.Interfaces.InlinePictureCollectionData;
    }
    /**
     *
     * Contains a collection of {@link Word.Paragraph} objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class List extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets paragraphs in the list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the list's id.
         *
         * [Api set: WordApi 1.3]
         */
        readonly id: number;
        /**
         *
         * Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly levelExistences: boolean[];
        /**
         *
         * Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly levelTypes: Word.ListLevelType[];
        
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
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
        
        /**
         *
         * Sets the alignment of the bullet, number or picture at the specified level in the list.
         *
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param alignment - Required. The level alignment that can be 'Left', 'Centered', or 'Right'.
         */
        setLevelAlignment(level: number, alignment: Word.Alignment): void;
        /**
         *
         * Sets the alignment of the bullet, number or picture at the specified level in the list.
         *
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param alignment - Required. The level alignment that can be 'Left', 'Centered', or 'Right'.
         */
        setLevelAlignment(level: number, alignment: "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"): void;
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
        setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string): void;
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
        setLevelBullet(level: number, listBullet: "Custom" | "Solid" | "Hollow" | "Square" | "Diamonds" | "Arrow" | "Checkmark", charCode?: number, fontName?: string): void;
        /**
         *
         * Sets the two indents of the specified level in the list.
         *
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param textIndent - Required. The text indent in points. It is the same as paragraph left indent.
         * @param bulletNumberPictureIndent - Required. The relative indent, in points, of the bullet, number or picture. It is the same as paragraph first line indent.
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
        setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: any[]): void;
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
        setLevelNumbering(level: number, listNumbering: "None" | "Arabic" | "UpperRoman" | "LowerRoman" | "UpperLetter" | "LowerLetter", formatString?: any[]): void;
        
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.List` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.List` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.List;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.List;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.List;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.List;
        toJSON(): Word.Interfaces.ListData;
    }
    /**
     *
     * Contains a collection of {@link Word.List} objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class ListCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Word.List[];
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.ListCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.ListCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.ListCollection;
        load(option?: OfficeExtension.LoadOption): Word.ListCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ListCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ListCollection;
        toJSON(): Word.Interfaces.ListCollectionData;
    }
    /**
     *
     * Represents the paragraph list item format.
     *
     * [Api set: WordApi 1.3]
     */
    export class ListItem extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets or sets the level of the item in the list.
         *
         * [Api set: WordApi 1.3]
         */
        level: number;
        /**
         *
         * Gets the list item bullet, number, or picture as a string. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly listString: string;
        /**
         *
         * Gets the list item order number in relation to its siblings. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly siblingIndex: number;
        
        /**
         *
         * Gets the list item parent, or the closest ancestor if the parent does not exist. Throws if the list item has no ancestor.
         *
         * [Api set: WordApi 1.3]
         *
         * @param parentOnly - Optional. Specifies only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.
         */
        getAncestor(parentOnly?: boolean): Word.Paragraph;
        /**
         *
         * Gets the list item parent, or the closest ancestor if the parent does not exist. Returns a null object if the list item has no ancestor.
         *
         * [Api set: WordApi 1.3]
         *
         * @param parentOnly - Optional. Specifies only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.
         */
        getAncestorOrNullObject(parentOnly?: boolean): Word.Paragraph;
        /**
         *
         * Gets all descendant list items of the list item.
         *
         * [Api set: WordApi 1.3]
         *
         * @param directChildrenOnly - Optional. Specifies only the list item's direct children will be returned. The default is false that indicates to get all descendant items.
         */
        getDescendants(directChildrenOnly?: boolean): Word.ParagraphCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.ListItem` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.ListItem` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.ListItem;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.ListItem;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ListItem;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ListItem;
        toJSON(): Word.Interfaces.ListItemData;
    }
    /**
     *
     * Represents a single paragraph in a selection, range, content control, or document body.
     *
     * [Api set: WordApi 1.1]
     */
    export class Paragraph extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the collection of content control objects in the paragraph. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        /**
         *
         * Gets the collection of InlinePicture objects in the paragraph. The collection does not include floating images. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the List to which this paragraph belongs. Throws if the paragraph is not in a list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly list: Word.List;
        /**
         *
         * Gets the ListItem for the paragraph. Throws if the paragraph is not part of a list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly listItem: Word.ListItem;
        /**
         *
         * Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly listItemOrNullObject: Word.ListItem;
        /**
         *
         * Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly listOrNullObject: Word.List;
        /**
         *
         * Gets the parent body of the paragraph. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the paragraph. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the table that contains the paragraph. Throws if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the paragraph. Throws if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCell: Word.TableCell;
        /**
         *
         * Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         *
         * Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         *
         * Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
         *
         * [Api set: WordApi 1.1]
         */
        alignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
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
        readonly isLastParagraph: boolean;
        /**
         *
         * Checks whether the paragraph is a list item. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly isListItem: boolean;
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
         * Gets or sets the amount of spacing, in grid lines, after the paragraph.
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
        styleBuiltIn: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        /**
         *
         * Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly tableNestingLevel: number;
        /**
         *
         * Gets the text of the paragraph. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        
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
         * Gets an HTML representation of the paragraph object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word Online, etc.). If you need exact fidelity, or consistency across platforms, use `Paragraph.getOoxml()` and convert the returned XML to HTML.
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
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: Word.RangeLocation): Word.Range;
        /**
         *
         * Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: "Whole" | "Start" | "End" | "Before" | "After" | "Content"): Word.Range;
        /**
         *
         * Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.
         *
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        getTextRanges(endingMarks: string[], trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add to the document.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation): void;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add to the document.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): void;
        /**
         *
         * Wraps the paragraph object with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the paragraph.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the paragraph.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertHtml(html: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation): Word.InlinePicture;
        /**
         *
         * Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.InlinePicture;
        /**
         *
         * Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in the paragraph.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in the paragraph.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
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
        insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][]): Word.Table;
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
        insertTable(rowCount: number, columnCount: number, insertLocation: "Before" | "After" | "Start" | "End" | "Replace", values?: string[][]): Word.Table;
        /**
         *
         * Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start', or 'End'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertText(text: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Performs a search with the specified SearchOptions on the scope of the paragraph object. The search results are a collection of range objects.
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
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         *
         * Selects and navigates the Word UI to the paragraph.
         *
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
        /**
         *
         * Splits the paragraph into child ranges by using delimiters.
         *
         * [Api set: WordApi 1.3]
         *
         * @param delimiters - Required. The delimiters as an array of strings.
         * @param trimDelimiters - Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        /**
         *
         * Starts a new list with this paragraph. Fails if the paragraph is already a list item.
         *
         * [Api set: WordApi 1.3]
         */
        startNewList(): Word.List;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Paragraph` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Paragraph` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.Paragraph;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.Paragraph;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Paragraph;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Paragraph;
        toJSON(): Word.Interfaces.ParagraphData;
    }
    /**
     *
     * Contains a collection of {@link Word.Paragraph} objects.
     *
     * [Api set: WordApi 1.1]
     */
    export class ParagraphCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Paragraph[];
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.ParagraphCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.ParagraphCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.ParagraphCollection;
        load(option?: OfficeExtension.LoadOption): Word.ParagraphCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ParagraphCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ParagraphCollection;
        toJSON(): Word.Interfaces.ParagraphCollectionData;
    }
    /**
     *
     * Represents a contiguous area in a document.
     *
     * [Api set: WordApi 1.1]
     */
    export class Range extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the collection of content control objects in the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         *
         * Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        /**
         *
         * Gets the collection of inline picture objects in the range. Read-only.
         *
         * [Api set: WordApi 1.2]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        /**
         *
         * Gets the collection of list objects in the range. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly lists: Word.ListCollection;
        /**
         *
         * Gets the collection of paragraph objects in the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly paragraphs: Word.ParagraphCollection;
        /**
         *
         * Gets the parent body of the range. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the range. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the range. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the table that contains the range. Throws if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains the range. Throws if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCell: Word.TableCell;
        /**
         *
         * Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         *
         * Gets the table that contains the range. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         *
         * Gets the collection of table objects in the range. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly tables: Word.TableCollection;
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
        readonly isEmpty: boolean;
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
        styleBuiltIn: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        /**
         *
         * Gets the text of the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        
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
        compareLocationWith(range: Word.Range): OfficeExtension.ClientResult<Word.LocationRelation>;
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
         * Gets an HTML representation of the range object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word Online, etc.). If you need exact fidelity, or consistency across platforms, use `Range.getOoxml()` and convert the returned XML to HTML.
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
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.
         */
        getNextTextRange(endingMarks: string[], trimSpacing?: boolean): Word.Range;
        /**
         *
         * Gets the next text range by using punctuation marks and/or other ending marks. Returns a null object if this text range is the last one.
         *
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.
         */
        getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean): Word.Range;
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
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: Word.RangeLocation): Word.Range;
        /**
         *
         * Clones the range, or gets the starting or ending point of the range as a new range.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: "Whole" | "Start" | "End" | "Before" | "After" | "Content"): Word.Range;
        /**
         *
         * Gets the text child ranges in the range by using punctuation marks and/or other ending marks.
         *
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        getTextRanges(endingMarks: string[], trimSpacing?: boolean): Word.RangeCollection;
        
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation): void;
        /**
         *
         * Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): void;
        /**
         *
         * Wraps the range object with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts a document at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertHtml(html: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation): Word.InlinePicture;
        /**
         *
         * Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
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
        insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][]): Word.Table;
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
        insertTable(rowCount: number, columnCount: number, insertLocation: "Before" | "After" | "Start" | "End" | "Replace", values?: string[][]): Word.Table;
        /**
         *
         * Inserts text at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts text at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertText(text: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
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
         * Performs a search with the specified SearchOptions on the scope of the range object. The search results are a collection of range objects.
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
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         *
         * Selects and navigates the Word UI to the range.
         *
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
        /**
         *
         * Splits the range into child ranges by using delimiters.
         *
         * [Api set: WordApi 1.3]
         *
         * @param delimiters - Required. The delimiters as an array of strings.
         * @param multiParagraphs - Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.
         * @param trimDelimiters - Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Range` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Range` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.Range;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.Range;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Range;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Range;
        toJSON(): Word.Interfaces.RangeData;
    }
    /**
     *
     * Contains a collection of {@link Word.Range} objects.
     *
     * [Api set: WordApi 1.1]
     */
    export class RangeCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Range[];
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.RangeCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.RangeCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.RangeCollection;
        load(option?: OfficeExtension.LoadOption): Word.RangeCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.RangeCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.RangeCollection;
        toJSON(): Word.Interfaces.RangeCollectionData;
    }
    /**
     *
     * Specifies the options to be included in a search operation.
     *
     * [Api set: WordApi 1.1]
     */
    export class SearchOptions extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
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
         * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
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
        
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.SearchOptions` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.SearchOptions` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.SearchOptions;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.SearchOptions;
        /**
         * Create a new instance of Word.SearchOptions object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): Word.SearchOptions;
        toJSON(): Word.Interfaces.SearchOptionsData;
    }
    /**
     *
     * Represents a section in a Word document.
     *
     * [Api set: WordApi 1.1]
     */
    export class Section extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the body object of the section. This does not include the header/footer and other section metadata. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly body: Word.Body;
        
        /**
         *
         * Gets one of the section's footers.
         *
         * [Api set: WordApi 1.1]
         *
         * @param type - Required. The type of footer to return. This value can be: 'Primary', 'FirstPage', or 'EvenPages'.
         */
        getFooter(type: Word.HeaderFooterType): Word.Body;
        /**
         *
         * Gets one of the section's footers.
         *
         * [Api set: WordApi 1.1]
         *
         * @param type - Required. The type of footer to return. This value can be: 'Primary', 'FirstPage', or 'EvenPages'.
         */
        getFooter(type: "Primary" | "FirstPage" | "EvenPages"): Word.Body;
        /**
         *
         * Gets one of the section's headers.
         *
         * [Api set: WordApi 1.1]
         *
         * @param type - Required. The type of header to return. This value can be: 'Primary', 'FirstPage', or 'EvenPages'.
         */
        getHeader(type: Word.HeaderFooterType): Word.Body;
        /**
         *
         * Gets one of the section's headers.
         *
         * [Api set: WordApi 1.1]
         *
         * @param type - Required. The type of header to return. This value can be: 'Primary', 'FirstPage', or 'EvenPages'.
         */
        getHeader(type: "Primary" | "FirstPage" | "EvenPages"): Word.Body;
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Section` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Section` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.Section;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.Section;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Section;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Section;
        toJSON(): Word.Interfaces.SectionData;
    }
    /**
     *
     * Contains the collection of the document's {@link Word.Section} objects.
     *
     * [Api set: WordApi 1.1]
     */
    export class SectionCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Section[];
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.SectionCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.SectionCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.SectionCollection;
        load(option?: OfficeExtension.LoadOption): Word.SectionCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.SectionCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.SectionCollection;
        toJSON(): Word.Interfaces.SectionCollectionData;
    }
    
    
    /**
     *
     * Represents a table in a Word document.
     *
     * [Api set: WordApi 1.3]
     */
    export class Table extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly font: Word.Font;
        /**
         *
         * Gets the parent body of the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentBody: Word.Body;
        /**
         *
         * Gets the content control that contains the table. Throws if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         *
         * Gets the content control that contains the table. Returns a null object if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         *
         * Gets the table that contains this table. Throws if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         *
         * Gets the table cell that contains this table. Throws if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCell: Word.TableCell;
        /**
         *
         * Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         *
         * Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         *
         * Gets all of the table rows. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly rows: Word.TableRowCollection;
        /**
         *
         * Gets the child tables nested one level deeper. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly tables: Word.TableCollection;
        /**
         *
         * Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
         *
         * [Api set: WordApi 1.3]
         */
        alignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
        /**
         *
         * Gets and sets the number of header rows.
         *
         * [Api set: WordApi 1.3]
         */
        headerRowCount: number;
        /**
         *
         * Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
         *
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
        /**
         *
         * Indicates whether all of the table rows are uniform. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly isUniform: boolean;
        /**
         *
         * Gets the nesting level of the table. Top-level tables have level 1. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly nestingLevel: number;
        /**
         *
         * Gets the number of rows in the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly rowCount: number;
        /**
         *
         * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
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
        styleBuiltIn: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
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
        values: string[][];
        /**
         *
         * Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
         *
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        /**
         *
         * Gets and sets the width of the table in points.
         *
         * [Api set: WordApi 1.3]
         */
        width: number;
        
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
        addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][]): void;
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
        addColumns(insertLocation: "Before" | "After" | "Start" | "End" | "Replace", columnCount: number, values?: string[][]): void;
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
        addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][]): Word.TableRowCollection;
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
        addRows(insertLocation: "Before" | "After" | "Start" | "End" | "Replace", rowCount: number, values?: string[][]): Word.TableRowCollection;
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
        getBorder(borderLocation: Word.BorderLocation): Word.TableBorder;
        /**
         *
         * Gets the border style for the specified border.
         *
         * [Api set: WordApi 1.3]
         *
         * @param borderLocation - Required. The border location.
         */
        getBorder(borderLocation: "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"): Word.TableBorder;
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
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocation: Word.CellPaddingLocation): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right"): OfficeExtension.ClientResult<number>;
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
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', or 'After'.
         */
        getRange(rangeLocation?: Word.RangeLocation): Word.Range;
        /**
         *
         * Gets the range that contains this table, or the range at the start or end of the table.
         *
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location can be 'Whole', 'Start', 'End', or 'After'.
         */
        getRange(rangeLocation?: "Whole" | "Start" | "End" | "Before" | "After" | "Content"): Word.Range;
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
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
         *
         * [Api set: WordApi 1.3]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
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
        insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][]): Word.Table;
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
        insertTable(rowCount: number, columnCount: number, insertLocation: "Before" | "After" | "Start" | "End" | "Replace", values?: string[][]): Word.Table;
        
        /**
         *
         * Performs a search with the specified SearchOptions on the scope of the table object. The search results are a collection of range objects.
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
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         *
         * Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
         *
         * [Api set: WordApi 1.3]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
        /**
         *
         * Sets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number): void;
        /**
         *
         * Sets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right", cellPadding: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Table` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Table` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.Table;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.Table;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.Table;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.Table;
        toJSON(): Word.Interfaces.TableData;
    }
    /**
     *
     * Contains the collection of the document's Table objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Table[];
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.TableCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.TableCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.TableCollection;
        load(option?: OfficeExtension.LoadOption): Word.TableCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableCollection;
        toJSON(): Word.Interfaces.TableCollectionData;
    }
    /**
     *
     * Represents a row in a Word document.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableRow extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets cells. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly cells: Word.TableCellCollection;
        /**
         *
         * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly font: Word.Font;
        /**
         *
         * Gets parent table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         *
         * Gets the number of cells in the row. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly cellCount: number;
        /**
         *
         * Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
         *
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
        /**
         *
         * Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
         *
         * [Api set: WordApi 1.3]
         */
        readonly isHeader: boolean;
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
        readonly rowIndex: number;
        /**
         *
         * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
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
        values: string[][];
        /**
         *
         * Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
         *
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        
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
        getBorder(borderLocation: Word.BorderLocation): Word.TableBorder;
        /**
         *
         * Gets the border style of the cells in the row.
         *
         * [Api set: WordApi 1.3]
         *
         * @param borderLocation - Required. The border location.
         */
        getBorder(borderLocation: "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"): Word.TableBorder;
        /**
         *
         * Gets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocation: Word.CellPaddingLocation): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right"): OfficeExtension.ClientResult<number>;
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
        insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][]): Word.TableRowCollection;
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
        insertRows(insertLocation: "Before" | "After" | "Start" | "End" | "Replace", rowCount: number, values?: string[][]): Word.TableRowCollection;
        
        /**
         *
         * Performs a search with the specified SearchOptions on the scope of the row. The search results are a collection of range objects.
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
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         *
         * Selects the row and navigates the Word UI to it.
         *
         * [Api set: WordApi 1.3]
         *
         * @param selectionMode - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
        /**
         *
         * Sets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number): void;
        /**
         *
         * Sets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right", cellPadding: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.TableRow` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.TableRow` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.TableRow;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.TableRow;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableRow;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableRow;
        toJSON(): Word.Interfaces.TableRowData;
    }
    /**
     *
     * Contains the collection of the document's TableRow objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableRowCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Word.TableRow[];
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.TableRowCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.TableRowCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.TableRowCollection;
        load(option?: OfficeExtension.LoadOption): Word.TableRowCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableRowCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableRowCollection;
        toJSON(): Word.Interfaces.TableRowCollectionData;
    }
    /**
     *
     * Represents a table cell in a Word document.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableCell extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the body object of the cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly body: Word.Body;
        /**
         *
         * Gets the parent row of the cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentRow: Word.TableRow;
        /**
         *
         * Gets the parent table of the cell. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         *
         * Gets the index of the cell in its row. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly cellIndex: number;
        /**
         *
         * Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
         *
         * [Api set: WordApi 1.3]
         */
        columnWidth: number;
        /**
         *
         * Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
         *
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
        /**
         *
         * Gets the index of the cell's row in the table. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly rowIndex: number;
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
         * Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
         *
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        /**
         *
         * Gets the width of the cell in points. Read-only.
         *
         * [Api set: WordApi 1.3]
         */
        readonly width: number;
        
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
        getBorder(borderLocation: Word.BorderLocation): Word.TableBorder;
        /**
         *
         * Gets the border style for the specified border.
         *
         * [Api set: WordApi 1.3]
         *
         * @param borderLocation - Required. The border location.
         */
        getBorder(borderLocation: "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"): Word.TableBorder;
        /**
         *
         * Gets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocation: Word.CellPaddingLocation): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right"): OfficeExtension.ClientResult<number>;
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
         * @param columnCount - Required. Number of columns to add.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][]): void;
        /**
         *
         * Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
         *
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. It can be 'Before' or 'After'.
         * @param columnCount - Required. Number of columns to add.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertColumns(insertLocation: "Before" | "After" | "Start" | "End" | "Replace", columnCount: number, values?: string[][]): void;
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
        insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][]): Word.TableRowCollection;
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
        insertRows(insertLocation: "Before" | "After" | "Start" | "End" | "Replace", rowCount: number, values?: string[][]): Word.TableRowCollection;
        /**
         *
         * Sets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number): void;
        /**
         *
         * Sets cell padding in points.
         *
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location can be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right", cellPadding: number): void;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.TableCell` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.TableCell` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.TableCell;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.TableCell;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableCell;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableCell;
        toJSON(): Word.Interfaces.TableCellData;
    }
    /**
     *
     * Contains the collection of the document's TableCell objects.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableCellCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Word.TableCell[];
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
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.TableCellCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.TableCellCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.TableCellCollection;
        load(option?: OfficeExtension.LoadOption): Word.TableCellCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableCellCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableCellCollection;
        toJSON(): Word.Interfaces.TableCellCollectionData;
    }
    /**
     *
     * Specifies the border style.
     *
     * [Api set: WordApi 1.3]
     */
    export class TableBorder extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets or sets the table border color.
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
        type: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
        /**
         *
         * Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
         *
         * [Api set: WordApi 1.3]
         */
        width: number;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         
         *
         * `load(option?: { select?: string; expand?: string; }): Word.TableBorder` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.TableBorder` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         * @param option - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(option?: string | string[]): Word.TableBorder;
        load(option?: {
            select?: string;
            expand?: string;
        }): Word.TableBorder;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.TableBorder;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.TableBorder;
        toJSON(): Word.Interfaces.TableBorderData;
    }
    /**
     *
     * Provides information about the type of a raised event. For each object type, please keep the order of: deleted, selection changed, data changed, added.
     *
     * [Api set: WordApi]
     */
    enum EventType {
        /**
         *
         * ContentControlDeleted represent the event that the content control has been deleted.
         *
         */
        contentControlDeleted = "ContentControlDeleted",
        /**
         *
         * ContentControlSelectionChanged represents the event that the selection in the content control has been changed.
         *
         */
        contentControlSelectionChanged = "ContentControlSelectionChanged",
        /**
         *
         * ContentControlDataChanged represents the event that the data in the content control have been changed.
         *
         */
        contentControlDataChanged = "ContentControlDataChanged",
        /**
         *
         * ContentControlAdded represents the event a content control has been added to the document.
         *
         */
        contentControlAdded = "ContentControlAdded",
    }
    
    /**
     *
     * Specifies supported content control types and subtypes.
     *
     * [Api set: WordApi]
     */
    enum ContentControlType {
        unknown = "Unknown",
        richTextInline = "RichTextInline",
        richTextParagraphs = "RichTextParagraphs",
        /**
         *
         * Contains a whole cell.
         *
         */
        richTextTableCell = "RichTextTableCell",
        /**
         *
         * Contains a whole row.
         *
         */
        richTextTableRow = "RichTextTableRow",
        /**
         *
         * Contains a whole table.
         *
         */
        richTextTable = "RichTextTable",
        plainTextInline = "PlainTextInline",
        plainTextParagraph = "PlainTextParagraph",
        picture = "Picture",
        buildingBlockGallery = "BuildingBlockGallery",
        checkBox = "CheckBox",
        comboBox = "ComboBox",
        dropDownList = "DropDownList",
        datePicker = "DatePicker",
        repeatingSection = "RepeatingSection",
        /**
         *
         * Identifies a rich text content control.
         *
         */
        richText = "RichText",
        plainText = "PlainText",
    }
    /**
     *
     * ContentControl appearance
     *
     * [Api set: WordApi]
     */
    enum ContentControlAppearance {
        /**
         *
         * Represents a content control shown as a shaded rectangle or bounding box (with optional title).
         *
         */
        boundingBox = "BoundingBox",
        /**
         *
         * Represents a content control shown as start and end markers.
         *
         */
        tags = "Tags",
        /**
         *
         * Represents a content control that is not shown.
         *
         */
        hidden = "Hidden",
    }
    /**
     *
     * Underline types
     *
     * [Api set: WordApi]
     */
    enum UnderlineType {
        mixed = "Mixed",
        /**
         *
         * No underline.
         *
         */
        none = "None",
        /**
         *
         * @deprecated Hidden is no longer supported.
         */
        hidden = "Hidden",
        /**
         *
         * @deprecated DotLine is no longer supported.
         */
        dotLine = "DotLine",
        /**
         *
         * A single underline. This is the default value.
         *
         */
        single = "Single",
        /**
         *
         * Only underline individual words.
         *
         */
        word = "Word",
        /**
         *
         * A double underline.
         *
         */
        double = "Double",
        /**
         *
         * A single thick underline.
         *
         */
        thick = "Thick",
        /**
         *
         * A dotted underline.
         *
         */
        dotted = "Dotted",
        dottedHeavy = "DottedHeavy",
        /**
         *
         * A single dash underline.
         *
         */
        dashLine = "DashLine",
        dashLineHeavy = "DashLineHeavy",
        dashLineLong = "DashLineLong",
        dashLineLongHeavy = "DashLineLongHeavy",
        /**
         *
         * An alternating dot-dash underline.
         *
         */
        dotDashLine = "DotDashLine",
        dotDashLineHeavy = "DotDashLineHeavy",
        /**
         *
         * An alternating dot-dot-dash underline.
         *
         */
        twoDotDashLine = "TwoDotDashLine",
        twoDotDashLineHeavy = "TwoDotDashLineHeavy",
        /**
         *
         * A single wavy underline.
         *
         */
        wave = "Wave",
        waveHeavy = "WaveHeavy",
        waveDouble = "WaveDouble",
    }
    /**
     *
     * Specifies the form of a break.
     *
     * [Api set: WordApi]
     */
    enum BreakType {
        /**
         *
         * Page break at the insertion point.
         *
         */
        page = "Page",
        /**
         *
         * @deprecated Use sectionNext instead.
         */
        next = "Next",
        /**
         *
         * Section break on next page.
         *
         */
        sectionNext = "SectionNext",
        /**
         *
         * New section without a corresponding page break.
         *
         */
        sectionContinuous = "SectionContinuous",
        /**
         *
         * Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.
         *
         */
        sectionEven = "SectionEven",
        /**
         *
         * Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.
         *
         */
        sectionOdd = "SectionOdd",
        /**
         *
         * Line break.
         *
         */
        line = "Line",
    }
    /**
     *
     * The insertion location types
     *
     * [Api set: WordApi]
     */
    enum InsertLocation {
        /**
         *
         * Add content before the contents of the calling object.
         *
         */
        before = "Before",
        /**
         *
         * Add content after the contents of the calling object.
         *
         */
        after = "After",
        /**
         *
         * Prepend content to the contents of the calling object.
         *
         */
        start = "Start",
        /**
         *
         * Append content to the contents of the calling object.
         *
         */
        end = "End",
        /**
         *
         * Replace the contents of the current object.
         *
         */
        replace = "Replace",
    }
    /**
     * [Api set: WordApi]
     */
    enum Alignment {
        mixed = "Mixed",
        /**
         *
         * Unknown alignment.
         *
         */
        unknown = "Unknown",
        /**
         *
         * Alignment to the left.
         *
         */
        left = "Left",
        /**
         *
         * Alignment to the center.
         *
         */
        centered = "Centered",
        /**
         *
         * Alignment to the right.
         *
         */
        right = "Right",
        /**
         *
         * Fully justified alignment.
         *
         */
        justified = "Justified",
    }
    /**
     * [Api set: WordApi]
     */
    enum HeaderFooterType {
        /**
         *
         * Returns the header or footer on all pages of a section, with the first page or odd pages excluded if they are different.
         *
         */
        primary = "Primary",
        /**
         *
         * Returns the header or footer on the first page of a section.
         *
         */
        firstPage = "FirstPage",
        /**
         *
         * Returns all headers or footers on even-numbered pages of a section.
         *
         */
        evenPages = "EvenPages",
    }
    /**
     * [Api set: WordApi]
     */
    enum BodyType {
        unknown = "Unknown",
        mainDoc = "MainDoc",
        section = "Section",
        header = "Header",
        footer = "Footer",
        tableCell = "TableCell",
    }
    /**
     * [Api set: WordApi]
     */
    enum SelectionMode {
        select = "Select",
        start = "Start",
        end = "End",
    }
    /**
     * [Api set: WordApi]
     */
    enum ImageFormat {
        unsupported = "Unsupported",
        undefined = "Undefined",
        bmp = "Bmp",
        jpeg = "Jpeg",
        gif = "Gif",
        tiff = "Tiff",
        png = "Png",
        icon = "Icon",
        exif = "Exif",
        wmf = "Wmf",
        emf = "Emf",
        pict = "Pict",
        pdf = "Pdf",
        svg = "Svg",
    }
    /**
     * [Api set: WordApi]
     */
    enum RangeLocation {
        /**
         *
         * The object's whole range. If the object is a paragraph content control or table content control, the EOP or Table characters after the content control are also included.
         *
         */
        whole = "Whole",
        /**
         *
         * The starting point of the object. For content control, it is the point after the opening tag.
         *
         */
        start = "Start",
        /**
         *
         * The ending point of the object. For paragraph, it is the point before the EOP. For content control, it is the point before the closing tag.
         *
         */
        end = "End",
        /**
         *
         * For content control only. It is the point before the opening tag.
         *
         */
        before = "Before",
        /**
         *
         * The point after the object. If the object is a paragraph content control or table content control, it is the point after the EOP or Table characters.
         *
         */
        after = "After",
        /**
         *
         * The range between 'Start' and 'End'.
         *
         */
        content = "Content",
    }
    /**
     * [Api set: WordApi]
     */
    enum LocationRelation {
        /**
         *
         * Indicates that this instance and the range are in different sub-documents.
         *
         */
        unrelated = "Unrelated",
        /**
         *
         * Indicates that this instance and the range represent the same range.
         *
         */
        equal = "Equal",
        /**
         *
         * Indicates that this instance contains the range and that it shares the same start character. The range does not share the same end character as this instance.
         *
         */
        containsStart = "ContainsStart",
        /**
         *
         * Indicates that this instance contains the range and that it shares the same end character. The range does not share the same start character as this instance.
         *
         */
        containsEnd = "ContainsEnd",
        /**
         *
         * Indicates that this instance contains the range, with the exception of the start and end character of this instance.
         *
         */
        contains = "Contains",
        /**
         *
         * Indicates that this instance is inside the range and that it shares the same start character. The range does not share the same end character as this instance.
         *
         */
        insideStart = "InsideStart",
        /**
         *
         * Indicates that this instance is inside the range and that it shares the same end character. The range does not share the same start character as this instance.
         *
         */
        insideEnd = "InsideEnd",
        /**
         *
         * Indicates that this instance is inside the range. The range does not share the same start and end characters as this instance.
         *
         */
        inside = "Inside",
        /**
         *
         * Indicates that this instance occurs before, and is adjacent to, the range.
         *
         */
        adjacentBefore = "AdjacentBefore",
        /**
         *
         * Indicates that this instance starts before the range and overlaps the range’s first character.
         *
         */
        overlapsBefore = "OverlapsBefore",
        /**
         *
         * Indicates that this instance occurs before the range.
         *
         */
        before = "Before",
        /**
         *
         * Indicates that this instance occurs after, and is adjacent to, the range.
         *
         */
        adjacentAfter = "AdjacentAfter",
        /**
         *
         * Indicates that this instance starts inside the range and overlaps the range’s last character.
         *
         */
        overlapsAfter = "OverlapsAfter",
        /**
         *
         * Indicates that this instance occurs after the range.
         *
         */
        after = "After",
    }
    /**
     * [Api set: WordApi]
     */
    enum BorderLocation {
        top = "Top",
        left = "Left",
        bottom = "Bottom",
        right = "Right",
        insideHorizontal = "InsideHorizontal",
        insideVertical = "InsideVertical",
        inside = "Inside",
        outside = "Outside",
        all = "All",
    }
    /**
     * [Api set: WordApi]
     */
    enum CellPaddingLocation {
        top = "Top",
        left = "Left",
        bottom = "Bottom",
        right = "Right",
    }
    /**
     * [Api set: WordApi]
     */
    enum BorderType {
        mixed = "Mixed",
        none = "None",
        single = "Single",
        double = "Double",
        dotted = "Dotted",
        dashed = "Dashed",
        dotDashed = "DotDashed",
        dot2Dashed = "Dot2Dashed",
        triple = "Triple",
        thinThickSmall = "ThinThickSmall",
        thickThinSmall = "ThickThinSmall",
        thinThickThinSmall = "ThinThickThinSmall",
        thinThickMed = "ThinThickMed",
        thickThinMed = "ThickThinMed",
        thinThickThinMed = "ThinThickThinMed",
        thinThickLarge = "ThinThickLarge",
        thickThinLarge = "ThickThinLarge",
        thinThickThinLarge = "ThinThickThinLarge",
        wave = "Wave",
        doubleWave = "DoubleWave",
        dashedSmall = "DashedSmall",
        dashDotStroked = "DashDotStroked",
        threeDEmboss = "ThreeDEmboss",
        threeDEngrave = "ThreeDEngrave",
    }
    /**
     * [Api set: WordApi]
     */
    enum VerticalAlignment {
        mixed = "Mixed",
        top = "Top",
        center = "Center",
        bottom = "Bottom",
    }
    /**
     * [Api set: WordApi]
     */
    enum ListLevelType {
        bullet = "Bullet",
        number = "Number",
        picture = "Picture",
    }
    /**
     * [Api set: WordApi]
     */
    enum ListBullet {
        custom = "Custom",
        solid = "Solid",
        hollow = "Hollow",
        square = "Square",
        diamonds = "Diamonds",
        arrow = "Arrow",
        checkmark = "Checkmark",
    }
    /**
     * [Api set: WordApi]
     */
    enum ListNumbering {
        none = "None",
        arabic = "Arabic",
        upperRoman = "UpperRoman",
        lowerRoman = "LowerRoman",
        upperLetter = "UpperLetter",
        lowerLetter = "LowerLetter",
    }
    /**
     * [Api set: WordApi]
     */
    enum Style {
        /**
         *
         * Mixed styles or other style not in this list.
         *
         */
        other = "Other",
        /**
         *
         * Reset character and paragraph style to default.
         *
         */
        normal = "Normal",
        heading1 = "Heading1",
        heading2 = "Heading2",
        heading3 = "Heading3",
        heading4 = "Heading4",
        heading5 = "Heading5",
        heading6 = "Heading6",
        heading7 = "Heading7",
        heading8 = "Heading8",
        heading9 = "Heading9",
        /**
         *
         * Table-of-content level 1.
         *
         */
        toc1 = "Toc1",
        /**
         *
         * Table-of-content level 2.
         *
         */
        toc2 = "Toc2",
        /**
         *
         * Table-of-content level 3.
         *
         */
        toc3 = "Toc3",
        /**
         *
         * Table-of-content level 4.
         *
         */
        toc4 = "Toc4",
        /**
         *
         * Table-of-content level 5.
         *
         */
        toc5 = "Toc5",
        /**
         *
         * Table-of-content level 6.
         *
         */
        toc6 = "Toc6",
        /**
         *
         * Table-of-content level 7.
         *
         */
        toc7 = "Toc7",
        /**
         *
         * Table-of-content level 8.
         *
         */
        toc8 = "Toc8",
        /**
         *
         * Table-of-content level 9.
         *
         */
        toc9 = "Toc9",
        footnoteText = "FootnoteText",
        header = "Header",
        footer = "Footer",
        caption = "Caption",
        footnoteReference = "FootnoteReference",
        endnoteReference = "EndnoteReference",
        endnoteText = "EndnoteText",
        title = "Title",
        subtitle = "Subtitle",
        hyperlink = "Hyperlink",
        strong = "Strong",
        emphasis = "Emphasis",
        noSpacing = "NoSpacing",
        listParagraph = "ListParagraph",
        quote = "Quote",
        intenseQuote = "IntenseQuote",
        subtleEmphasis = "SubtleEmphasis",
        intenseEmphasis = "IntenseEmphasis",
        subtleReference = "SubtleReference",
        intenseReference = "IntenseReference",
        bookTitle = "BookTitle",
        bibliography = "Bibliography",
        /**
         *
         * Table-of-content heading.
         *
         */
        tocHeading = "TocHeading",
        tableGrid = "TableGrid",
        plainTable1 = "PlainTable1",
        plainTable2 = "PlainTable2",
        plainTable3 = "PlainTable3",
        plainTable4 = "PlainTable4",
        plainTable5 = "PlainTable5",
        tableGridLight = "TableGridLight",
        gridTable1Light = "GridTable1Light",
        gridTable1Light_Accent1 = "GridTable1Light_Accent1",
        gridTable1Light_Accent2 = "GridTable1Light_Accent2",
        gridTable1Light_Accent3 = "GridTable1Light_Accent3",
        gridTable1Light_Accent4 = "GridTable1Light_Accent4",
        gridTable1Light_Accent5 = "GridTable1Light_Accent5",
        gridTable1Light_Accent6 = "GridTable1Light_Accent6",
        gridTable2 = "GridTable2",
        gridTable2_Accent1 = "GridTable2_Accent1",
        gridTable2_Accent2 = "GridTable2_Accent2",
        gridTable2_Accent3 = "GridTable2_Accent3",
        gridTable2_Accent4 = "GridTable2_Accent4",
        gridTable2_Accent5 = "GridTable2_Accent5",
        gridTable2_Accent6 = "GridTable2_Accent6",
        gridTable3 = "GridTable3",
        gridTable3_Accent1 = "GridTable3_Accent1",
        gridTable3_Accent2 = "GridTable3_Accent2",
        gridTable3_Accent3 = "GridTable3_Accent3",
        gridTable3_Accent4 = "GridTable3_Accent4",
        gridTable3_Accent5 = "GridTable3_Accent5",
        gridTable3_Accent6 = "GridTable3_Accent6",
        gridTable4 = "GridTable4",
        gridTable4_Accent1 = "GridTable4_Accent1",
        gridTable4_Accent2 = "GridTable4_Accent2",
        gridTable4_Accent3 = "GridTable4_Accent3",
        gridTable4_Accent4 = "GridTable4_Accent4",
        gridTable4_Accent5 = "GridTable4_Accent5",
        gridTable4_Accent6 = "GridTable4_Accent6",
        gridTable5Dark = "GridTable5Dark",
        gridTable5Dark_Accent1 = "GridTable5Dark_Accent1",
        gridTable5Dark_Accent2 = "GridTable5Dark_Accent2",
        gridTable5Dark_Accent3 = "GridTable5Dark_Accent3",
        gridTable5Dark_Accent4 = "GridTable5Dark_Accent4",
        gridTable5Dark_Accent5 = "GridTable5Dark_Accent5",
        gridTable5Dark_Accent6 = "GridTable5Dark_Accent6",
        gridTable6Colorful = "GridTable6Colorful",
        gridTable6Colorful_Accent1 = "GridTable6Colorful_Accent1",
        gridTable6Colorful_Accent2 = "GridTable6Colorful_Accent2",
        gridTable6Colorful_Accent3 = "GridTable6Colorful_Accent3",
        gridTable6Colorful_Accent4 = "GridTable6Colorful_Accent4",
        gridTable6Colorful_Accent5 = "GridTable6Colorful_Accent5",
        gridTable6Colorful_Accent6 = "GridTable6Colorful_Accent6",
        gridTable7Colorful = "GridTable7Colorful",
        gridTable7Colorful_Accent1 = "GridTable7Colorful_Accent1",
        gridTable7Colorful_Accent2 = "GridTable7Colorful_Accent2",
        gridTable7Colorful_Accent3 = "GridTable7Colorful_Accent3",
        gridTable7Colorful_Accent4 = "GridTable7Colorful_Accent4",
        gridTable7Colorful_Accent5 = "GridTable7Colorful_Accent5",
        gridTable7Colorful_Accent6 = "GridTable7Colorful_Accent6",
        listTable1Light = "ListTable1Light",
        listTable1Light_Accent1 = "ListTable1Light_Accent1",
        listTable1Light_Accent2 = "ListTable1Light_Accent2",
        listTable1Light_Accent3 = "ListTable1Light_Accent3",
        listTable1Light_Accent4 = "ListTable1Light_Accent4",
        listTable1Light_Accent5 = "ListTable1Light_Accent5",
        listTable1Light_Accent6 = "ListTable1Light_Accent6",
        listTable2 = "ListTable2",
        listTable2_Accent1 = "ListTable2_Accent1",
        listTable2_Accent2 = "ListTable2_Accent2",
        listTable2_Accent3 = "ListTable2_Accent3",
        listTable2_Accent4 = "ListTable2_Accent4",
        listTable2_Accent5 = "ListTable2_Accent5",
        listTable2_Accent6 = "ListTable2_Accent6",
        listTable3 = "ListTable3",
        listTable3_Accent1 = "ListTable3_Accent1",
        listTable3_Accent2 = "ListTable3_Accent2",
        listTable3_Accent3 = "ListTable3_Accent3",
        listTable3_Accent4 = "ListTable3_Accent4",
        listTable3_Accent5 = "ListTable3_Accent5",
        listTable3_Accent6 = "ListTable3_Accent6",
        listTable4 = "ListTable4",
        listTable4_Accent1 = "ListTable4_Accent1",
        listTable4_Accent2 = "ListTable4_Accent2",
        listTable4_Accent3 = "ListTable4_Accent3",
        listTable4_Accent4 = "ListTable4_Accent4",
        listTable4_Accent5 = "ListTable4_Accent5",
        listTable4_Accent6 = "ListTable4_Accent6",
        listTable5Dark = "ListTable5Dark",
        listTable5Dark_Accent1 = "ListTable5Dark_Accent1",
        listTable5Dark_Accent2 = "ListTable5Dark_Accent2",
        listTable5Dark_Accent3 = "ListTable5Dark_Accent3",
        listTable5Dark_Accent4 = "ListTable5Dark_Accent4",
        listTable5Dark_Accent5 = "ListTable5Dark_Accent5",
        listTable5Dark_Accent6 = "ListTable5Dark_Accent6",
        listTable6Colorful = "ListTable6Colorful",
        listTable6Colorful_Accent1 = "ListTable6Colorful_Accent1",
        listTable6Colorful_Accent2 = "ListTable6Colorful_Accent2",
        listTable6Colorful_Accent3 = "ListTable6Colorful_Accent3",
        listTable6Colorful_Accent4 = "ListTable6Colorful_Accent4",
        listTable6Colorful_Accent5 = "ListTable6Colorful_Accent5",
        listTable6Colorful_Accent6 = "ListTable6Colorful_Accent6",
        listTable7Colorful = "ListTable7Colorful",
        listTable7Colorful_Accent1 = "ListTable7Colorful_Accent1",
        listTable7Colorful_Accent2 = "ListTable7Colorful_Accent2",
        listTable7Colorful_Accent3 = "ListTable7Colorful_Accent3",
        listTable7Colorful_Accent4 = "ListTable7Colorful_Accent4",
        listTable7Colorful_Accent5 = "ListTable7Colorful_Accent5",
        listTable7Colorful_Accent6 = "ListTable7Colorful_Accent6",
    }
    /**
     * [Api set: WordApi]
     */
    enum DocumentPropertyType {
        string = "String",
        number = "Number",
        date = "Date",
        boolean = "Boolean",
    }
    /**
     * [Api set: WordApi]
     */
    enum TapObjectType {
        chart = "Chart",
        smartArt = "SmartArt",
        table = "Table",
        image = "Image",
        slide = "Slide",
        ole = "OLE",
        text = "Text",
    }
    /**
     * [Api set: WordApi]
     */
    enum FileContentFormat {
        base64 = "Base64",
        html = "Html",
        ooxml = "Ooxml",
    }
    enum ErrorCodes {
        accessDenied = "AccessDenied",
        generalException = "GeneralException",
        invalidArgument = "InvalidArgument",
        itemNotFound = "ItemNotFound",
        notImplemented = "NotImplemented",
    }
    export module Interfaces {
        /**
        * Provides ways to load properties of only a subset of members of a collection.
        */
        
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
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
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
             * Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
             *
             * [Api set: WordApi 1.1]
             */
            appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
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
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
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
        /** An interface for updating data on the ContentControlCollection object, for use in "contentControlCollection.set({ ... })". */
        export interface ContentControlCollectionUpdateData {
            items?: Word.Interfaces.ContentControlData[];
        }
        /** An interface for updating data on the CustomProperty object, for use in "customProperty.set({ ... })". */
        export interface CustomPropertyUpdateData {
            /**
             *
             * Gets or sets the value of the custom property. Note that even though Word Online and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
             *
             * [Api set: WordApi 1.3]
             */
            value?: any;
        }
        /** An interface for updating data on the CustomPropertyCollection object, for use in "customPropertyCollection.set({ ... })". */
        export interface CustomPropertyCollectionUpdateData {
            items?: Word.Interfaces.CustomPropertyData[];
        }
        /** An interface for updating data on the CustomXmlPartCollection object, for use in "customXmlPartCollection.set({ ... })". */
        export interface CustomXmlPartCollectionUpdateData {
            items?: Word.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the CustomXmlPartScopedCollection object, for use in "customXmlPartScopedCollection.set({ ... })". */
        export interface CustomXmlPartScopedCollectionUpdateData {
            items?: Word.Interfaces.CustomXmlPartData[];
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
            * Gets the properties of the document.
            *
            * [Api set: WordApi 1.3]
            */
            properties?: Word.Interfaces.DocumentPropertiesUpdateData;
        }
        /** An interface for updating data on the DocumentCreated object, for use in "documentCreated.set({ ... })". */
        export interface DocumentCreatedUpdateData {
            /**
            *
            * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc..
            *
            * [Api set: WordApiHiddenDocument 1.3]
            */
            body?: Word.Interfaces.BodyUpdateData;
            /**
            *
            * Gets the properties of the document.
            *
            * [Api set: WordApiHiddenDocument 1.3]
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
             * Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
             *
             * [Api set: WordApi 1.1]
             */
            doubleStrikeThrough?: boolean;
            /**
             *
             * Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
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
             * Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
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
            underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
        }
        /** An interface for updating data on the InlinePicture object, for use in "inlinePicture.set({ ... })". */
        export interface InlinePictureUpdateData {
            /**
             *
             * Gets or sets a string that represents the alternative text associated with the inline image.
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
        /** An interface for updating data on the InlinePictureCollection object, for use in "inlinePictureCollection.set({ ... })". */
        export interface InlinePictureCollectionUpdateData {
            items?: Word.Interfaces.InlinePictureData[];
        }
        /** An interface for updating data on the ListCollection object, for use in "listCollection.set({ ... })". */
        export interface ListCollectionUpdateData {
            items?: Word.Interfaces.ListData[];
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
            alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
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
             * Gets or sets the amount of spacing, in grid lines, after the paragraph.
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
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        }
        /** An interface for updating data on the ParagraphCollection object, for use in "paragraphCollection.set({ ... })". */
        export interface ParagraphCollectionUpdateData {
            items?: Word.Interfaces.ParagraphData[];
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
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        }
        /** An interface for updating data on the RangeCollection object, for use in "rangeCollection.set({ ... })". */
        export interface RangeCollectionUpdateData {
            items?: Word.Interfaces.RangeData[];
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
             * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
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
        /** An interface for updating data on the SectionCollection object, for use in "sectionCollection.set({ ... })". */
        export interface SectionCollectionUpdateData {
            items?: Word.Interfaces.SectionData[];
        }
        /** An interface for updating data on the Setting object, for use in "setting.set({ ... })". */
        export interface SettingUpdateData {
            
        }
        /** An interface for updating data on the SettingCollection object, for use in "settingCollection.set({ ... })". */
        export interface SettingCollectionUpdateData {
            items?: Word.Interfaces.SettingData[];
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
             * Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
             *
             * [Api set: WordApi 1.3]
             */
            alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             *
             * Gets and sets the number of header rows.
             *
             * [Api set: WordApi 1.3]
             */
            headerRowCount?: number;
            /**
             *
             * Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             *
             * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
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
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
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
            values?: string[][];
            /**
             *
             * Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
            /**
             *
             * Gets and sets the width of the table in points.
             *
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /** An interface for updating data on the TableCollection object, for use in "tableCollection.set({ ... })". */
        export interface TableCollectionUpdateData {
            items?: Word.Interfaces.TableData[];
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
             * Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             *
             * Gets and sets the preferred height of the row in points.
             *
             * [Api set: WordApi 1.3]
             */
            preferredHeight?: number;
            /**
             *
             * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
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
            values?: string[][];
            /**
             *
             * Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        }
        /** An interface for updating data on the TableRowCollection object, for use in "tableRowCollection.set({ ... })". */
        export interface TableRowCollectionUpdateData {
            items?: Word.Interfaces.TableRowData[];
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
             * Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
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
             * Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        }
        /** An interface for updating data on the TableCellCollection object, for use in "tableCellCollection.set({ ... })". */
        export interface TableCellCollectionUpdateData {
            items?: Word.Interfaces.TableCellData[];
        }
        /** An interface for updating data on the TableBorder object, for use in "tableBorder.set({ ... })". */
        export interface TableBorderUpdateData {
            /**
             *
             * Gets or sets the table border color.
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
            type?: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
            /**
             *
             * Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
             *
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling "body.toJSON()". */
        export interface BodyData {
            /**
            *
            * Gets the collection of rich text content control objects in the body. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            /**
            *
            * Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            /**
            *
            * Gets the collection of InlinePicture objects in the body. The collection does not include floating images. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            /**
            *
            * Gets the collection of list objects in the body. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            lists?: Word.Interfaces.ListData[];
            /**
            *
            * Gets the collection of paragraph objects in the body. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            /**
            *
            * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws if there isn't a parent body. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyData;
            /**
            *
            * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Returns a null object if there isn't a parent body. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentBodyOrNullObject?: Word.Interfaces.BodyData;
            /**
            *
            * Gets the content control that contains the body. Throws if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the content control that contains the body. Returns a null object if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the parent section of the body. Throws if there isn't a parent section. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentSection?: Word.Interfaces.SectionData;
            /**
            *
            * Gets the parent section of the body. Returns a null object if there isn't a parent section. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentSectionOrNullObject?: Word.Interfaces.SectionData;
            /**
            *
            * Gets the collection of table objects in the body. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            tables?: Word.Interfaces.TableData[];
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
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             *
             * Gets the text of the body. Use the insertText method to insert text. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: string;
            /**
             *
             * Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            type?: Word.BodyType | "Unknown" | "MainDoc" | "Section" | "Header" | "Footer" | "TableCell";
        }
        /** An interface describing the data returned by calling "contentControl.toJSON()". */
        export interface ContentControlData {
            /**
            *
            * Gets the collection of content control objects in the content control. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            /**
            *
            * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            /**
            *
            * Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            /**
            *
            * Gets the collection of list objects in the content control. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            lists?: Word.Interfaces.ListData[];
            /**
            *
            * Get the collection of paragraph objects in the content control. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            /**
            *
            * Gets the parent body of the content control. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyData;
            /**
            *
            * Gets the content control that contains the content control. Throws if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the content control that contains the content control. Returns a null object if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the table that contains the content control. Throws if it is not contained in a table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableData;
            /**
            *
            * Gets the table cell that contains the content control. Throws if it is not contained in a table cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellData;
            /**
            *
            * Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellData;
            /**
            *
            * Gets the table that contains the content control. Returns a null object if it is not contained in a table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableData;
            /**
            *
            * Gets the collection of table objects in the content control. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            tables?: Word.Interfaces.TableData[];
            /**
             *
             * Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
             *
             * [Api set: WordApi 1.1]
             */
            appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
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
             * Gets an integer that represents the content control identifier. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            id?: number;
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
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             *
             * Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            subtype?: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText";
            /**
             *
             * Gets or sets a tag to identify a content control.
             *
             * [Api set: WordApi 1.1]
             */
            tag?: string;
            /**
             *
             * Gets the text of the content control. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: string;
            /**
             *
             * Gets or sets the title for a content control.
             *
             * [Api set: WordApi 1.1]
             */
            title?: string;
            /**
             *
             * Gets the content control type. Only rich text content controls are supported currently. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            type?: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText";
        }
        /** An interface describing the data returned by calling "contentControlCollection.toJSON()". */
        export interface ContentControlCollectionData {
            items?: Word.Interfaces.ContentControlData[];
        }
        /** An interface describing the data returned by calling "customProperty.toJSON()". */
        export interface CustomPropertyData {
            /**
             *
             * Gets the key of the custom property. Read only.
             *
             * [Api set: WordApi 1.3]
             */
            key?: string;
            /**
             *
             * Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.
             *
             * [Api set: WordApi 1.3]
             */
            type?: Word.DocumentPropertyType | "String" | "Number" | "Date" | "Boolean";
            /**
             *
             * Gets or sets the value of the custom property. Note that even though Word Online and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
             *
             * [Api set: WordApi 1.3]
             */
            value?: any;
        }
        /** An interface describing the data returned by calling "customPropertyCollection.toJSON()". */
        export interface CustomPropertyCollectionData {
            items?: Word.Interfaces.CustomPropertyData[];
        }
        /** An interface describing the data returned by calling "customXmlPart.toJSON()". */
        export interface CustomXmlPartData {
            
            
        }
        /** An interface describing the data returned by calling "customXmlPartCollection.toJSON()". */
        export interface CustomXmlPartCollectionData {
            items?: Word.Interfaces.CustomXmlPartData[];
        }
        /** An interface describing the data returned by calling "customXmlPartScopedCollection.toJSON()". */
        export interface CustomXmlPartScopedCollectionData {
            items?: Word.Interfaces.CustomXmlPartData[];
        }
        /** An interface describing the data returned by calling "document.toJSON()". */
        export interface DocumentData {
            /**
            *
            * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyData;
            /**
            *
            * Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            
            /**
            *
            * Gets the properties of the document. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            properties?: Word.Interfaces.DocumentPropertiesData;
            /**
            *
            * Gets the collection of section objects in the document. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            sections?: Word.Interfaces.SectionData[];
            
            /**
             *
             * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            saved?: boolean;
        }
        /** An interface describing the data returned by calling "documentCreated.toJSON()". */
        export interface DocumentCreatedData {
            /**
            *
            * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.
            *
            * [Api set: WordApiHiddenDocument 1.3]
            */
            body?: Word.Interfaces.BodyData;
            /**
            *
            * Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.
            *
            * [Api set: WordApiHiddenDocument 1.3]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            /**
            *
            * Gets the custom XML parts in the document. Read-only.
            *
            * [Api set: WordApiHiddenDocument 1.4]
            */
            customXmlParts?: Word.Interfaces.CustomXmlPartData[];
            /**
            *
            * Gets the properties of the document. Read-only.
            *
            * [Api set: WordApiHiddenDocument 1.3]
            */
            properties?: Word.Interfaces.DocumentPropertiesData;
            /**
            *
            * Gets the collection of section objects in the document. Read-only.
            *
            * [Api set: WordApiHiddenDocument 1.3]
            */
            sections?: Word.Interfaces.SectionData[];
            /**
            *
            * Gets the add-in's settings in the document. Read-only.
            *
            * [Api set: WordApiHiddenDocument 1.4]
            */
            settings?: Word.Interfaces.SettingData[];
            /**
             *
             * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
             *
             * [Api set: WordApiHiddenDocument 1.3]
             */
            saved?: boolean;
        }
        /** An interface describing the data returned by calling "documentProperties.toJSON()". */
        export interface DocumentPropertiesData {
            /**
            *
            * Gets the collection of custom properties of the document. Read only.
            *
            * [Api set: WordApi 1.3]
            */
            customProperties?: Word.Interfaces.CustomPropertyData[];
            /**
             *
             * Gets the application name of the document. Read only.
             *
             * [Api set: WordApi 1.3]
             */
            applicationName?: string;
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
             * Gets the creation date of the document. Read only.
             *
             * [Api set: WordApi 1.3]
             */
            creationDate?: Date;
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
             * Gets the last author of the document. Read only.
             *
             * [Api set: WordApi 1.3]
             */
            lastAuthor?: string;
            /**
             *
             * Gets the last print date of the document. Read only.
             *
             * [Api set: WordApi 1.3]
             */
            lastPrintDate?: Date;
            /**
             *
             * Gets the last save time of the document. Read only.
             *
             * [Api set: WordApi 1.3]
             */
            lastSaveTime?: Date;
            /**
             *
             * Gets or sets the manager of the document.
             *
             * [Api set: WordApi 1.3]
             */
            manager?: string;
            /**
             *
             * Gets the revision number of the document. Read only.
             *
             * [Api set: WordApi 1.3]
             */
            revisionNumber?: string;
            /**
             *
             * Gets the security of the document. Read only.
             *
             * [Api set: WordApi 1.3]
             */
            security?: number;
            /**
             *
             * Gets or sets the subject of the document.
             *
             * [Api set: WordApi 1.3]
             */
            subject?: string;
            /**
             *
             * Gets the template of the document. Read only.
             *
             * [Api set: WordApi 1.3]
             */
            template?: string;
            /**
             *
             * Gets or sets the title of the document.
             *
             * [Api set: WordApi 1.3]
             */
            title?: string;
        }
        /** An interface describing the data returned by calling "font.toJSON()". */
        export interface FontData {
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
             * Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
             *
             * [Api set: WordApi 1.1]
             */
            doubleStrikeThrough?: boolean;
            /**
             *
             * Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
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
             * Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
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
            underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
        }
        /** An interface describing the data returned by calling "inlinePicture.toJSON()". */
        export interface InlinePictureData {
            /**
            *
            * Gets the parent paragraph that contains the inline image. Read-only.
            *
            * [Api set: WordApi 1.2]
            */
            paragraph?: Word.Interfaces.ParagraphData;
            /**
            *
            * Gets the content control that contains the inline image. Throws if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the table that contains the inline image. Throws if it is not contained in a table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableData;
            /**
            *
            * Gets the table cell that contains the inline image. Throws if it is not contained in a table cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellData;
            /**
            *
            * Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellData;
            /**
            *
            * Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableData;
            /**
             *
             * Gets or sets a string that represents the alternative text associated with the inline image.
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
        /** An interface describing the data returned by calling "inlinePictureCollection.toJSON()". */
        export interface InlinePictureCollectionData {
            items?: Word.Interfaces.InlinePictureData[];
        }
        /** An interface describing the data returned by calling "list.toJSON()". */
        export interface ListData {
            /**
            *
            * Gets paragraphs in the list. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            /**
             *
             * Gets the list's id.
             *
             * [Api set: WordApi 1.3]
             */
            id?: number;
            /**
             *
             * Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            levelExistences?: boolean[];
            /**
             *
             * Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            levelTypes?: Word.ListLevelType[];
        }
        /** An interface describing the data returned by calling "listCollection.toJSON()". */
        export interface ListCollectionData {
            items?: Word.Interfaces.ListData[];
        }
        /** An interface describing the data returned by calling "listItem.toJSON()". */
        export interface ListItemData {
            /**
             *
             * Gets or sets the level of the item in the list.
             *
             * [Api set: WordApi 1.3]
             */
            level?: number;
            /**
             *
             * Gets the list item bullet, number, or picture as a string. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            listString?: string;
            /**
             *
             * Gets the list item order number in relation to its siblings. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            siblingIndex?: number;
        }
        /** An interface describing the data returned by calling "paragraph.toJSON()". */
        export interface ParagraphData {
            /**
            *
            * Gets the collection of content control objects in the paragraph. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            /**
            *
            * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            /**
            *
            * Gets the collection of InlinePicture objects in the paragraph. The collection does not include floating images. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            /**
            *
            * Gets the List to which this paragraph belongs. Throws if the paragraph is not in a list. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            list?: Word.Interfaces.ListData;
            /**
            *
            * Gets the ListItem for the paragraph. Throws if the paragraph is not part of a list. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            listItem?: Word.Interfaces.ListItemData;
            /**
            *
            * Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            listItemOrNullObject?: Word.Interfaces.ListItemData;
            /**
            *
            * Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            listOrNullObject?: Word.Interfaces.ListData;
            /**
            *
            * Gets the parent body of the paragraph. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyData;
            /**
            *
            * Gets the content control that contains the paragraph. Throws if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the table that contains the paragraph. Throws if it is not contained in a table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableData;
            /**
            *
            * Gets the table cell that contains the paragraph. Throws if it is not contained in a table cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellData;
            /**
            *
            * Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellData;
            /**
            *
            * Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableData;
            /**
             *
             * Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * [Api set: WordApi 1.1]
             */
            alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             *
             * Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: number;
            /**
             *
             * Indicates the paragraph is the last one inside its parent body. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            isLastParagraph?: boolean;
            /**
             *
             * Checks whether the paragraph is a list item. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            isListItem?: boolean;
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
             * Gets or sets the amount of spacing, in grid lines, after the paragraph.
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
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             *
             * Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            tableNestingLevel?: number;
            /**
             *
             * Gets the text of the paragraph. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling "paragraphCollection.toJSON()". */
        export interface ParagraphCollectionData {
            items?: Word.Interfaces.ParagraphData[];
        }
        /** An interface describing the data returned by calling "range.toJSON()". */
        export interface RangeData {
            /**
            *
            * Gets the collection of content control objects in the range. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            /**
            *
            * Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            /**
            *
            * Gets the collection of inline picture objects in the range. Read-only.
            *
            * [Api set: WordApi 1.2]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            /**
            *
            * Gets the collection of list objects in the range. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            lists?: Word.Interfaces.ListData[];
            /**
            *
            * Gets the collection of paragraph objects in the range. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            /**
            *
            * Gets the parent body of the range. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyData;
            /**
            *
            * Gets the content control that contains the range. Throws if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the content control that contains the range. Returns a null object if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the table that contains the range. Throws if it is not contained in a table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableData;
            /**
            *
            * Gets the table cell that contains the range. Throws if it is not contained in a table cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellData;
            /**
            *
            * Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellData;
            /**
            *
            * Gets the table that contains the range. Returns a null object if it is not contained in a table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableData;
            /**
            *
            * Gets the collection of table objects in the range. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            tables?: Word.Interfaces.TableData[];
            /**
             *
             * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
             *
             * [Api set: WordApi 1.3]
             */
            hyperlink?: string;
            /**
             *
             * Checks whether the range length is zero. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            isEmpty?: boolean;
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
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             *
             * Gets the text of the range. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling "rangeCollection.toJSON()". */
        export interface RangeCollectionData {
            items?: Word.Interfaces.RangeData[];
        }
        /** An interface describing the data returned by calling "searchOptions.toJSON()". */
        export interface SearchOptionsData {
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
             * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
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
        /** An interface describing the data returned by calling "section.toJSON()". */
        export interface SectionData {
            /**
            *
            * Gets the body object of the section. This does not include the header/footer and other section metadata. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyData;
        }
        /** An interface describing the data returned by calling "sectionCollection.toJSON()". */
        export interface SectionCollectionData {
            items?: Word.Interfaces.SectionData[];
        }
        /** An interface describing the data returned by calling "setting.toJSON()". */
        export interface SettingData {
            
            
        }
        /** An interface describing the data returned by calling "settingCollection.toJSON()". */
        export interface SettingCollectionData {
            items?: Word.Interfaces.SettingData[];
        }
        /** An interface describing the data returned by calling "table.toJSON()". */
        export interface TableData {
            /**
            *
            * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontData;
            /**
            *
            * Gets the parent body of the table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyData;
            /**
            *
            * Gets the content control that contains the table. Throws if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentContentControl?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the content control that contains the table. Returns a null object if there isn't a parent content control. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlData;
            /**
            *
            * Gets the table that contains this table. Throws if it is not contained in a table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableData;
            /**
            *
            * Gets the table cell that contains this table. Throws if it is not contained in a table cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellData;
            /**
            *
            * Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellData;
            /**
            *
            * Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableData;
            /**
            *
            * Gets all of the table rows. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            rows?: Word.Interfaces.TableRowData[];
            /**
            *
            * Gets the child tables nested one level deeper. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            tables?: Word.Interfaces.TableData[];
            /**
             *
             * Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
             *
             * [Api set: WordApi 1.3]
             */
            alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             *
             * Gets and sets the number of header rows.
             *
             * [Api set: WordApi 1.3]
             */
            headerRowCount?: number;
            /**
             *
             * Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             *
             * Indicates whether all of the table rows are uniform. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            isUniform?: boolean;
            /**
             *
             * Gets the nesting level of the table. Top-level tables have level 1. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            nestingLevel?: number;
            /**
             *
             * Gets the number of rows in the table. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            rowCount?: number;
            /**
             *
             * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
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
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
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
            values?: string[][];
            /**
             *
             * Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
            /**
             *
             * Gets and sets the width of the table in points.
             *
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling "tableCollection.toJSON()". */
        export interface TableCollectionData {
            items?: Word.Interfaces.TableData[];
        }
        /** An interface describing the data returned by calling "tableRow.toJSON()". */
        export interface TableRowData {
            /**
            *
            * Gets cells. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            cells?: Word.Interfaces.TableCellData[];
            /**
            *
            * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontData;
            /**
            *
            * Gets parent table. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableData;
            /**
             *
             * Gets the number of cells in the row. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            cellCount?: number;
            /**
             *
             * Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             *
             * Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
             *
             * [Api set: WordApi 1.3]
             */
            isHeader?: boolean;
            /**
             *
             * Gets and sets the preferred height of the row in points.
             *
             * [Api set: WordApi 1.3]
             */
            preferredHeight?: number;
            /**
             *
             * Gets the index of the row in its parent table. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            rowIndex?: number;
            /**
             *
             * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
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
            values?: string[][];
            /**
             *
             * Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        }
        /** An interface describing the data returned by calling "tableRowCollection.toJSON()". */
        export interface TableRowCollectionData {
            items?: Word.Interfaces.TableRowData[];
        }
        /** An interface describing the data returned by calling "tableCell.toJSON()". */
        export interface TableCellData {
            /**
            *
            * Gets the body object of the cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            body?: Word.Interfaces.BodyData;
            /**
            *
            * Gets the parent row of the cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentRow?: Word.Interfaces.TableRowData;
            /**
            *
            * Gets the parent table of the cell. Read-only.
            *
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableData;
            /**
             *
             * Gets the index of the cell in its row. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            cellIndex?: number;
            /**
             *
             * Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
             *
             * [Api set: WordApi 1.3]
             */
            columnWidth?: number;
            /**
             *
             * Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             *
             * Gets the index of the cell's row in the table. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            rowIndex?: number;
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
             * Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
            /**
             *
             * Gets the width of the cell in points. Read-only.
             *
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling "tableCellCollection.toJSON()". */
        export interface TableCellCollectionData {
            items?: Word.Interfaces.TableCellData[];
        }
        /** An interface describing the data returned by calling "tableBorder.toJSON()". */
        export interface TableBorderData {
            /**
             *
             * Gets or sets the table border color.
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
            type?: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
            /**
             *
             * Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
             *
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /**
         *
         * Represents the body of a document or a section.
         *
         * [Api set: WordApi 1.1]
         */
        
        /**
         *
         * Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
         *
         * [Api set: WordApi 1.1]
         */
        
        /**
         *
         * Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
         *
         * [Api set: WordApi 1.1]
         */
        
        /**
         *
         * Represents a custom property.
         *
         * [Api set: WordApi 1.3]
         */
        
        /**
         *
         * Contains the collection of {@link Word.CustomProperty} objects.
         *
         * [Api set: WordApi 1.3]
         */
        
        
        readonly document: Document;
        readonly application: Application;
    }
    /**
     * Executes a batch script that performs actions on the Word object model, using the RequestContext of previously created API objects.
     * @param objects - An array of previously created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Word object model, using the RequestContext of a previously created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    export function run<T>(object: OfficeExtension.ClientObject, batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Word object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     *
     * @remarks
     *
     * In addition to this signature, the method also has the following signatures, which allow you to resume using the request context of previously created objects:
     *
     * run<T>(object: OfficeExtension.ClientObject, batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;
     *
     * run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;
     */
    export function run<T>(batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;
}


////////////////////////////////////////////////////////////////
//////////////////////// End Word APIs /////////////////////////
////////////////////////////////////////////////////////////////