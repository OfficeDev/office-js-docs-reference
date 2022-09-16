import { OfficeExtension } from "../api-extractor-inputs-office/office"
import { Office as Outlook} from "../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
/////////////////////// Begin Word APIs ////////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Word {
    /**
     * Represents the application object.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Creates a new document by using an optional base64 encoded .docx file.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param base64File - Optional. The base64 encoded .docx file. The default value is null.
         */
        createDocument(base64File?: string): Word.DocumentCreated;
        /**
         * Create a new instance of Word.Application object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): Word.Application;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Application object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): {
            [key: string]: string;
        };
    }
    /**
     * Represents the body of a document or a section.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class Body extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the collection of rich text content control objects in the body. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         * Gets the collection of endnotes in the body. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly endnotes: Word.NoteItemCollection;
        /**
         * Gets the collection of field objects in the body. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly fields: Word.FieldCollection;
        /**
         * Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        /**
         * Gets the collection of footnotes in the body. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly footnotes: Word.NoteItemCollection;
        /**
         * Gets the collection of InlinePicture objects in the body. The collection does not include floating images. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        /**
         * Gets the collection of list objects in the body. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly lists: Word.ListCollection;
        /**
         * Gets the collection of paragraph objects in the body. Read-only. **Important**: Paragraphs in tables are not returned for requirement sets 1.1 and 1.2. From requirement set 1.3, paragraphs in tables are also returned.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly paragraphs: Word.ParagraphCollection;
        /**
         * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws an error if there isn't a parent body. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentBody: Word.Body;
        /**
         * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Returns a null object if there isn't a parent body. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentBodyOrNullObject: Word.Body;
        /**
         * Gets the content control that contains the body. Throws an error if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         * Gets the content control that contains the body. Returns a null object if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         * Gets the parent section of the body. Throws an error if there isn't a parent section. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentSection: Word.Section;
        /**
         * Gets the parent section of the body. Returns a null object if there isn't a parent section. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentSectionOrNullObject: Word.Section;
        /**
         * Gets the collection of table objects in the body. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly tables: Word.TableCollection;
        /**
         * Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         * Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        /**
         * Gets the text of the body. Use the insertText method to insert text. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        /**
         * Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly type: Word.BodyType | "Unknown" | "MainDoc" | "Section" | "Header" | "Footer" | "TableCell" | "Footnote" | "Endnote" | "NoteItem";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.BodyUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Body): void;
        /**
         * Clears the contents of the body object. The user can perform the undo operation on the cleared content.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         * Gets comments associated with the body.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getComments(): Word.CommentCollection;
        /**
         * Gets an HTML representation of the body object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Body.getOoxml()` and convert the returned XML to HTML.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         * Gets the OOXML (Office Open XML) representation of the body object.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         * Gets the whole body, or the starting or ending point of the body, as a range.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location must be 'Whole', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | Word.RangeLocation.after | Word.RangeLocation.content | "Whole" | "Start" | "End" | "After" | "Content"): Word.Range;
        /**
         * Gets reviewed text based on ChangeTrackingVersion selection.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param changeTrackingVersion - Optional. The value must be 'Original' or 'Current'. The default is 'Current'.
         */
        getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion): OfficeExtension.ClientResult<string>;
        /**
         * Gets reviewed text based on ChangeTrackingVersion selection.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param changeTrackingVersionString - Optional. The value must be 'Original' or 'Current'. The default is 'Current'.
         */
        getReviewedText(changeTrackingVersionString?: "Original" | "Current"): OfficeExtension.ClientResult<string>;
        /**
         * Inserts a break at the specified location in the main document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add to the body.
         * @param insertLocation - Required. The value must be 'Start' or 'End'.
         */
        insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End"): void;
        /**
         * Wraps the body object with a Rich Text content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         * Inserts a document into the body at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts HTML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the document.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts a picture into the body at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the body.
         * @param insertLocation - Required. The value must be 'Start' or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End"): Word.InlinePicture;
        /**
         * Inserts OOXML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value must be 'Start' or 'End'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End"): Word.Paragraph;
        /**
         * Inserts a table with the specified number of rows and columns.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value must be 'Start' or 'End'.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End", values?: string[][]): Word.Table;
        /**
         * Inserts text into the body at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Performs a search with the specified SearchOptions on the scope of the body object. The search results are a collection of range objects.
         *
         * @remarks
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
         * Selects the body and navigates the Word UI to it.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects the body and navigates the Word UI to it.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionModeString - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.BodyLoadOptions): Word.Body;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Body;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Body;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Body;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Body;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Body object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BodyData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.BodyData;
    }
    /**
     * Represents a comment in the document.
     *
     * @remarks
     * [Api set: WordApiOnline 1.1]
     */
    export class Comment extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets or sets the comment's content range.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        contentRange: Word.CommentContentRange;
        /**
         * Gets the collection of reply objects associated with the comment.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly replies: Word.CommentReplyCollection;
        /**
         * Gets the email of the comment's author.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly authorEmail: string;
        /**
         * Gets the name of the comment's author.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly authorName: string;
        /**
         * Gets or sets the comment's content as plain text.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        content: string;
        /**
         * Gets the creation date of the comment.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly creationDate: Date;
        /**
         * Gets the ID of the comment. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly id: string;
        /**
         * Gets or sets the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        resolved: boolean;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.CommentUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Comment): void;
        /**
         * Deletes the comment and its replies.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        delete(): void;
        /**
         * Gets the range in the main document where the comment is on.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getRange(): Word.Range;
        /**
         * Adds a new reply to the end of the comment thread.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param replyText - Required. Reply text.
         */
        reply(replyText: string): Word.CommentReply;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.CommentLoadOptions): Word.Comment;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Comment;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Comment;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Comment;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Comment;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Comment object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CommentData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.CommentData;
    }
    /**
     * Contains a collection of {@link Word.Comment} objects.
     *
     * @remarks
     * [Api set: WordApiOnline 1.1]
     */
    export class CommentCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Comment[];
        /**
         * Gets the first comment in the collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getFirst(): Word.Comment;
        /**
         * Gets the first comment in the collection. If the collection is empty, returns a null object.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getFirstOrNullObject(): Word.Comment;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.CommentCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CommentCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.CommentCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CommentCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.CommentCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.CommentCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.CommentCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CommentCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.CommentCollectionData;
    }
    /**
     * @remarks
     * [Api set: WordApiOnline 1.1]
     */
    export class CommentContentRange extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets or sets a value that indicates whether the comment text is bold.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        bold: boolean;
        /**
         * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        hyperlink: string;
        /**
         * Checks whether the range length is zero. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly isEmpty: boolean;
        /**
         * Gets or sets a value that indicates whether the comment text is italicized.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        italic: boolean;
        /**
         * Gets or sets a value that indicates whether the comment text has a strikethrough.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        strikeThrough: boolean;
        /**
         * Gets the text of the comment range. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly text: string;
        /**
         * Gets or sets a value that indicates the comment text's underline type. 'None' if the comment text is not underlined.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        underline: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.CommentContentRangeUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.CommentContentRange): void;
        /**
         * Inserts text into at the specified location. **Note**: For the modern comment, the content range tracked across context turns to empty if any revision to the comment is posted through the UI.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param text - Required. The text to be inserted in to the CommentContentRange.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.CommentContentRange;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.CommentContentRangeLoadOptions): Word.CommentContentRange;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.CommentContentRange;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.CommentContentRange;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.CommentContentRange;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.CommentContentRange;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.CommentContentRange object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CommentContentRangeData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.CommentContentRangeData;
    }
    /**
     * Represents a comment reply in the document.
     *
     * @remarks
     * [Api set: WordApiOnline 1.1]
     */
    export class CommentReply extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets or sets the commentReply's content range.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        contentRange: Word.CommentContentRange;
        /**
         * Gets the parent comment of this reply.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly parentComment: Word.Comment;
        /**
         * Gets the email of the comment reply's author.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly authorEmail: string;
        /**
         * Gets the name of the comment reply's author.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly authorName: string;
        /**
         * Gets or sets the comment reply's content. The string is plain text.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        content: string;
        /**
         * Gets the creation date of the comment reply.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly creationDate: Date;
        /**
         * Gets the ID of the comment reply. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly id: string;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.CommentReplyUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.CommentReply): void;
        /**
         * Deletes the comment reply.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        delete(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.CommentReplyLoadOptions): Word.CommentReply;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.CommentReply;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.CommentReply;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.CommentReply;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.CommentReply;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.CommentReply object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CommentReplyData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.CommentReplyData;
    }
    /**
     * Contains a collection of {@link Word.CommentReply} objects. Represents all comment replies in one comment thread.
     *
     * @remarks
     * [Api set: WordApiOnline 1.1]
     */
    export class CommentReplyCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.CommentReply[];
        /**
         * Gets the first comment reply in the collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getFirst(): Word.CommentReply;
        /**
         * Gets the first comment reply in the collection. If the collection is empty, returns a null object.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getFirstOrNullObject(): Word.CommentReply;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.CommentReplyCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CommentReplyCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.CommentReplyCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CommentReplyCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.CommentReplyCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.CommentReplyCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.CommentReplyCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CommentReplyCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.CommentReplyCollectionData;
    }
    /**
     * Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class ContentControl extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the collection of content control objects in the content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         * Gets the collection of endnotes in the contentcontrol. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly endnotes: Word.NoteItemCollection;
        /**
         * Gets the collection of field objects in the contentcontrol. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly fields: Word.FieldCollection;
        /**
         * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        /**
         * Gets the collection of footnotes in the contentcontrol. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly footnotes: Word.NoteItemCollection;
        /**
         * Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        /**
         * Gets the collection of list objects in the content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly lists: Word.ListCollection;
        /**
         * Gets the collection of paragraph objects in the content control. Read-only. **Important**: For requirement sets 1.1 and 1.2, paragraphs in tables wholly contained within this content control are not returned. From requirement set 1.3, paragraphs in such tables are also returned.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly paragraphs: Word.ParagraphCollection;
        /**
         * Gets the parent body of the content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentBody: Word.Body;
        /**
         * Gets the content control that contains the content control. Throws an error if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         * Gets the content control that contains the content control. Returns a null object if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         * Gets the table that contains the content control. Throws an error if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         * Gets the table cell that contains the content control. Throws an error if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCell: Word.TableCell;
        /**
         * Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         * Gets the table that contains the content control. Returns a null object if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         * Gets the collection of table objects in the content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly tables: Word.TableCollection;
        /**
         * Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        appearance: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
        /**
         * Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        cannotDelete: boolean;
        /**
         * Gets or sets a value that indicates whether the user can edit the contents of the content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        cannotEdit: boolean;
        /**
         * Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        color: string;
        /**
         * Gets an integer that represents the content control identifier. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly id: number;
        /**
         * Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
         * **Note**: The set operation for this property is not supported in Word on the web.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        placeholderText: string;
        /**
         * Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        removeWhenEdited: boolean;
        /**
         * Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         * Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        /**
         * Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly subtype: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText";
        /**
         * Gets or sets a tag to identify a content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        tag: string;
        /**
         * Gets the text of the content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        /**
         * Gets or sets the title for a content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        title: string;
        /**
         * Gets the content control type. Only rich text content controls are supported currently. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly type: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ContentControlUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.ContentControl): void;
        /**
         * Clears the contents of the content control. The user can perform the undo operation on the cleared content.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         * Deletes the content control and its content. If keepContent is set to true, the content is not deleted.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param keepContent - Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted.
         */
        delete(keepContent: boolean): void;
        /**
         * Gets comments associated with the content control.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getComments(): Word.CommentCollection;
        /**
         * Gets an HTML representation of the content control object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `ContentControl.getOoxml()` and convert the returned XML to HTML.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         * Gets the Office Open XML (OOXML) representation of the content control object.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         * Gets the whole content control, or the starting or ending point of the content control, as a range.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location must be 'Whole', 'Start', 'End', 'Before', 'After', or 'Content'.
         */
        getRange(rangeLocation?: Word.RangeLocation | "Whole" | "Start" | "End" | "Before" | "After" | "Content"): Word.Range;
        /**
         * Gets reviewed text based on ChangeTrackingVersion selection.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param changeTrackingVersion - Optional. The value must be 'Original' or 'Current'. The default is 'Current'.
         */
        getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion): OfficeExtension.ClientResult<string>;
        /**
         * Gets reviewed text based on ChangeTrackingVersion selection.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param changeTrackingVersionString - Optional. The value must be 'Original' or 'Current'. The default is 'Current'.
         */
        getReviewedText(changeTrackingVersionString?: "Original" | "Current"): OfficeExtension.ClientResult<string>;
        /**
         * Gets the text ranges in the content control by using punctuation marks and/or other ending marks.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        getTextRanges(endingMarks: string[], trimSpacing?: boolean): Word.RangeCollection;
        /**
         * Inserts a break at the specified location in the main document. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. Type of break.
         * @param insertLocation - Required. The value must be 'Start', 'End', 'Before', or 'After'.
         */
        insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"): void;
        /**
         * Inserts a document into the content control at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts HTML into the content control at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in to the content control.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts an inline picture into the content control at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the content control.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.InlinePicture;
        /**
         * Inserts OOXML into the content control at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in to the content control.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value must be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"): Word.Paragraph;
        /**
         * Inserts a table with the specified number of rows and columns into, or next to, a content control.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value must be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After", values?: string[][]): Word.Table;
        /**
         * Inserts text into the content control at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. The text to be inserted in to the content control.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertText(text: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Performs a search with the specified SearchOptions on the scope of the content control object. The search results are a collection of range objects.
         *
         * @remarks
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
         * Selects the content control. This causes Word to scroll to the selection.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects the content control. This causes Word to scroll to the selection.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionModeString - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        /**
         * Splits the content control into child ranges by using delimiters.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param delimiters - Required. The delimiters as an array of strings.
         * @param multiParagraphs - Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.
         * @param trimDelimiters - Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.ContentControlLoadOptions): Word.ContentControl;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.ContentControl;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.ContentControl;
        /**
         * Occurs when data within the content control are changed. To get the new text, load this content control in the handler. To get the old text, do not load it.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * @eventproperty
         * @beta
         */
        readonly onDataChanged: OfficeExtension.EventHandlers<Word.ContentControlEventArgs>;
        /**
         * Occurs when the content control is deleted. Do not load this content control in the handler, otherwise you won't be able to get its original properties.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * @eventproperty
         * @beta
         */
        readonly onDeleted: OfficeExtension.EventHandlers<Word.ContentControlEventArgs>;
        /**
         * Occurs when selection within the content control is changed.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * @eventproperty
         * @beta
         */
        readonly onSelectionChanged: OfficeExtension.EventHandlers<Word.ContentControlEventArgs>;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.ContentControl;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.ContentControl;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.ContentControl object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.ContentControlData;
    }
    /**
     * Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class ContentControlCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.ContentControl[];
        /**
         * Gets a content control by its identifier. Throws an error if there isn't a content control with the identifier in this collection.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param id - Required. A content control identifier.
         */
        getById(id: number): Word.ContentControl;
        /**
         * Gets a content control by its identifier. Returns a null object if there isn't a content control with the identifier in this collection.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param id - Required. A content control identifier.
         */
        getByIdOrNullObject(id: number): Word.ContentControl;
        /**
         * Gets the content controls that have the specified tag.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param tag - Required. A tag set on a content control.
         */
        getByTag(tag: string): Word.ContentControlCollection;
        /**
         * Gets the content controls that have the specified title.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param title - Required. The title of a content control.
         */
        getByTitle(title: string): Word.ContentControlCollection;
        /**
         * Gets the content controls that have the specified types and/or subtypes.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param types - Required. An array of content control types and/or subtypes.
         */
        getByTypes(types: Word.ContentControlType[]): Word.ContentControlCollection;
        /**
         * Gets the first content control in this collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.ContentControl;
        /**
         * Gets the first content control in this collection. Returns a null object if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.ContentControl;
        /**
         * Gets a content control by its ID.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param id - The content control's ID.
         */
        getItem(id: number): Word.ContentControl;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.ContentControlCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.ContentControlCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.ContentControlCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.ContentControlCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.ContentControlCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.ContentControlCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.ContentControlCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.ContentControlCollectionData;
    }
    /**
     * Represents a custom property.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class CustomProperty extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the key of the custom property. Read only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly key: string;
        /**
         * Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly type: Word.DocumentPropertyType | "String" | "Number" | "Date" | "Boolean";
        /**
         * Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        value: any;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.CustomPropertyUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.CustomProperty): void;
        /**
         * Deletes the custom property.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        delete(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.CustomPropertyLoadOptions): Word.CustomProperty;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.CustomProperty;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.CustomProperty;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.CustomProperty;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.CustomProperty;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.CustomProperty object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomPropertyData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.CustomPropertyData;
    }
    /**
     * Contains the collection of {@link Word.CustomProperty} objects.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class CustomPropertyCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.CustomProperty[];
        /**
         * Creates a new or sets an existing custom property.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param key - Required. The custom property's key, which is case-insensitive.
         * @param value - Required. The custom property's value.
         */
        add(key: string, value: any): Word.CustomProperty;
        /**
         * Deletes all custom properties in this collection.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        deleteAll(): void;
        /**
         * Gets the count of custom properties.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a custom property object by its key, which is case-insensitive. Throws an error if the custom property does not exist.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param key - The key that identifies the custom property object.
         */
        getItem(key: string): Word.CustomProperty;
        /**
         * Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param key - Required. The key that identifies the custom property object.
         */
        getItemOrNullObject(key: string): Word.CustomProperty;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.CustomPropertyCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CustomPropertyCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.CustomPropertyCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CustomPropertyCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.CustomPropertyCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.CustomPropertyCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.CustomPropertyCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomPropertyCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.CustomPropertyCollectionData;
    }
    /**
     * Represents a custom XML part.
     *
     * @remarks
     * [Api set: WordApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class CustomXmlPart extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the ID of the custom XML part. Read only.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly id: string;
        /**
         * Gets the namespace URI of the custom XML part. Read only.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly namespaceUri: string;
        /**
         * Deletes the custom XML part.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        delete(): void;
        /**
         * Deletes an attribute with the given name from the element identified by xpath.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI),
         * an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose
         * the tree is the following:
         *
         * ```xml
         * <Day>
         *   <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
         *     <Week>something</Week>
         *   </Month>
         * </Day>
         * ```
         *
         * The xpath to `<Week>` must be \/Day\/greg:Month\/Week, where greg is an alias that is mapped to
         * "http:\/\/calendartypes.org\/xsds\/GregorianCalendar" in the namespaceMappings parameter.
         * @beta
         *
         * @param xpath - Required. Absolute path to the single element in XPath notation.
         * @param namespaceMappings - Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that does not used reserved XPath characters, such as the forward slash "/".
         * @param name - Required. Name of the attribute.
         */
        deleteAttribute(xpath: string, namespaceMappings: {
            [key: string]: string;
        }, name: string): void;
        /**
         * Deletes the element identified by xpath.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI),
         * an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose
         * the tree is the following:
         *
         * ```xml
         * <Day>
         *   <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
         *     <Week>something</Week>
         *   </Month>
         * </Day>
         * ```
         *
         * The xpath to `<Week>` must be \/Day\/greg:Month\/Week, where greg is an alias that is mapped to
         * "http:\/\/calendartypes.org\/xsds\/GregorianCalendar" in the namespaceMappings parameter.
         * @beta
         *
         * @param xpath - Required. Absolute path to the single element in XPath notation.
         * @param namespaceMappings - Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that does not used reserved XPath characters, such as the forward slash "/".
         */
        deleteElement(xpath: string, namespaceMappings: {
            [key: string]: string;
        }): void;
        /**
         * Gets the full XML content of the custom XML part.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        getXml(): OfficeExtension.ClientResult<string>;
        /**
         * Inserts an attribute with the given name and value to the element identified by xpath.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI),
         * an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose
         * the tree is the following:
         *
         * ```xml
         * <Day>
         *   <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
         *     <Week>something</Week>
         *   </Month>
         * </Day>
         * ```
         *
         * The xpath to `<Week>` must be \/Day\/greg:Month\/Week, where greg is an alias that is mapped to
         * "http:\/\/calendartypes.org\/xsds\/GregorianCalendar" in the namespaceMappings parameter.
         * @beta
         *
         * @param xpath - Required. Absolute path to the single element in XPath notation.
         * @param namespaceMappings - Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that does not used reserved XPath characters, such as the forward slash "/".
         * @param name - Required. Name of the attribute.
         * @param value - Required. Value of the attribute.
         */
        insertAttribute(xpath: string, namespaceMappings: {
            [key: string]: string;
        }, name: string, value: string): void;
        /**
         * Inserts the given XML under the parent element identified by xpath at child position index.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI),
         * an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose
         * the tree is the following:
         *
         * ```xml
         * <Day>
         *   <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
         *     <Week>something</Week>
         *   </Month>
         * </Day>
         * ```
         *
         * The xpath to `<Week>` must be \/Day\/greg:Month\/Week, where greg is an alias that is mapped to
         * "http:\/\/calendartypes.org\/xsds\/GregorianCalendar" in the namespaceMappings parameter.
         * @beta
         *
         * @param xpath - Required. Absolute path to the single parent element in XPath notation.
         * @param xml - Required. XML content to be inserted.
         * @param namespaceMappings - Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that does not used reserved XPath characters, such as the forward slash "/".
         * @param index - Optional. Zero-based position at which the new XML to be inserted. If omitted, the XML will be appended as the last child of this parent.
         */
        insertElement(xpath: string, xml: string, namespaceMappings: {
            [key: string]: string;
        }, index?: number): void;
        /**
         * Queries the XML content of the custom XML part.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI),
         * an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose
         * the tree is the following:
         *
         * ```xml
         * <Day>
         *   <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
         *     <Week>something</Week>
         *   </Month>
         * </Day>
         * ```
         *
         * The xpath to `<Week>` must be \/Day\/greg:Month\/Week, where greg is an alias that is mapped to
         * "http:\/\/calendartypes.org\/xsds\/GregorianCalendar" in the namespaceMappings parameter.
         * @beta
         *
         * @param xpath - Required. An XPath query.
         * @param namespaceMappings - Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that does not used reserved XPath characters, such as the forward slash "/".
         * @returns An array where each item represents an entry matched by the XPath query.
         */
        query(xpath: string, namespaceMappings: {
            [key: string]: string;
        }): OfficeExtension.ClientResult<string[]>;
        /**
         * Sets the full XML content of the custom XML part.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param xml - Required. XML content to be set.
         */
        setXml(xml: string): void;
        /**
         * Updates the value of an attribute with the given name of the element identified by xpath.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI),
         * an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose
         * the tree is the following:
         *
         * ```xml
         * <Day>
         *   <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
         *     <Week>something</Week>
         *   </Month>
         * </Day>
         * ```
         *
         * The xpath to `<Week>` must be \/Day\/greg:Month\/Week, where greg is an alias that is mapped to
         * "http:\/\/calendartypes.org\/xsds\/GregorianCalendar" in the namespaceMappings parameter.
         * @beta
         *
         * @param xpath - Required. Absolute path to the single element in XPath notation.
         * @param namespaceMappings - Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that does not used reserved XPath characters, such as the forward slash "/".
         * @param name - Required. Name of the attribute.
         * @param value - Required. New value of the attribute.
         */
        updateAttribute(xpath: string, namespaceMappings: {
            [key: string]: string;
        }, name: string, value: string): void;
        /**
         * Updates the XML of the element identified by xpath.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI),
         * an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose
         * the tree is the following:
         *
         * ```xml
         * <Day>
         *   <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
         *     <Week>something</Week>
         *   </Month>
         * </Day>
         * ```
         *
         * The xpath to `<Week>` must be \/Day\/greg:Month\/Week, where greg is an alias that is mapped to
         * "http:\/\/calendartypes.org\/xsds\/GregorianCalendar" in the namespaceMappings parameter.
         * @beta
         *
         * @param xpath - Required. Absolute path to the single element in XPath notation.
         * @param xml - Required. New XML content to be stored.
         * @param namespaceMappings - Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that does not used reserved XPath characters, such as the forward slash "/".
         */
        updateElement(xpath: string, xml: string, namespaceMappings: {
            [key: string]: string;
        }): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.CustomXmlPartLoadOptions): Word.CustomXmlPart;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.CustomXmlPart;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.CustomXmlPart;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.CustomXmlPart;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.CustomXmlPart;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.CustomXmlPart object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlPartData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.CustomXmlPartData;
    }
    /**
     * Contains the collection of {@link Word.CustomXmlPart} objects.
     *
     * @remarks
     * [Api set: WordApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class CustomXmlPartCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.CustomXmlPart[];
        /**
         * Adds a new custom XML part to the document.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param xml - Required. XML content. Must be a valid XML fragment.
         */
        add(xml: string): Word.CustomXmlPart;
        /**
         * Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param namespaceUri - Required. The namespace URI.
         */
        getByNamespace(namespaceUri: string): Word.CustomXmlPartScopedCollection;
        /**
         * Gets the number of items in the collection.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a custom XML part based on its ID. Read only.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param id - ID or index of the custom XML part to be retrieved.
         */
        getItem(id: string): Word.CustomXmlPart;
        /**
         * Gets a custom XML part based on its ID. Returns a null object if the CustomXmlPart does not exist.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param id - Required. ID of the object to be retrieved.
         */
        getItemOrNullObject(id: string): Word.CustomXmlPart;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.CustomXmlPartCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CustomXmlPartCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.CustomXmlPartCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CustomXmlPartCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.CustomXmlPartCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.CustomXmlPartCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.CustomXmlPartCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlPartCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.CustomXmlPartCollectionData;
    }
    /**
     * Contains the collection of {@link Word.CustomXmlPart} objects with a specific namespace.
     *
     * @remarks
     * [Api set: WordApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class CustomXmlPartScopedCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.CustomXmlPart[];
        /**
         * Gets the number of items in the collection.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a custom XML part based on its ID. Read only.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param id - ID of the custom XML part to be retrieved.
         */
        getItem(id: string): Word.CustomXmlPart;
        /**
         * Gets a custom XML part based on its ID. Returns a null object if the CustomXmlPart does not exist in the collection.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param id - Required. ID of the object to be retrieved.
         */
        getItemOrNullObject(id: string): Word.CustomXmlPart;
        /**
         * If the collection contains exactly one item, this method returns it. Otherwise, this method produces an error.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        getOnlyItem(): Word.CustomXmlPart;
        /**
         * If the collection contains exactly one item, this method returns it. Otherwise, this method returns a null object.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        getOnlyItemOrNullObject(): Word.CustomXmlPart;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.CustomXmlPartScopedCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CustomXmlPartScopedCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.CustomXmlPartScopedCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CustomXmlPartScopedCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.CustomXmlPartScopedCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.CustomXmlPartScopedCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.CustomXmlPartScopedCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlPartScopedCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.CustomXmlPartScopedCollectionData;
    }
    /**
     * The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class Document extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly body: Word.Body;
        /**
         * Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         * Gets the custom XML parts in the document. Read-only.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly customXmlParts: Word.CustomXmlPartCollection;
        /**
         * Gets the properties of the document. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly properties: Word.DocumentProperties;
        /**
         * Gets the collection of section objects in the document. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly sections: Word.SectionCollection;
        /**
         * Gets the add-in's settings in the document. Read-only.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly settings: Word.SettingCollection;
        /**
         * Gets or sets the ChangeTracking mode.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        changeTrackingMode: Word.ChangeTrackingMode | "Off" | "TrackAll" | "TrackMineOnly";
        /**
         * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly saved: boolean;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.DocumentUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Document): void;
        /**
         * Deletes a bookmark, if it exists, from the document.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param name - Required. The bookmark name, which is case-insensitive.
         */
        deleteBookmark(name: string): void;
        /**
         * Gets a bookmark's range. Throws an error if the bookmark does not exist.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param name - Required. The bookmark name, which is case-insensitive.
         */
        getBookmarkRange(name: string): Word.Range;
        /**
         * Gets a bookmark's range. Returns a null object if the bookmark does not exist.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param name - Required. The bookmark name, which is case-insensitive.
         */
        getBookmarkRangeOrNullObject(name: string): Word.Range;
        /**
         * Gets the document's endnotes in a single body. Read-only.
                    Not implemented in Word on the web.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getEndnoteBody(): Word.Body;
        /**
         * Gets the document's footnotes in a single body. Read-only.
                    Not implemented in Word on the web.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getFootnoteBody(): Word.Body;
        /**
         * Gets the current selection of the document. Multiple selections are not supported.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getSelection(): Word.Range;
        /**
         * Saves the document. This uses the Word default file naming convention if the document has not been saved before.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        save(): void;
        /**
         * Performs a search with the specified search options on the scope of the whole document. The search results are a collection of range objects.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
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
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.DocumentLoadOptions): Word.Document;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Document;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Document;
        /**
         * Occurs when a content control is added. Run context.sync() in the handler to get the new content control's properties.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         *
         * @eventproperty
         * @beta
         */
        readonly onContentControlAdded: OfficeExtension.EventHandlers<Word.ContentControlEventArgs>;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Document;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Document;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Document object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.DocumentData;
    }
    /**
     * The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class DocumentCreated extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc. Read-only.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.3]
         */
        readonly body: Word.Body;
        /**
         * Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc. Read-only.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.3]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         * Gets the custom XML parts in the document. Read-only.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.4]
         * @beta
         */
        readonly customXmlParts: Word.CustomXmlPartCollection;
        /**
         * Gets the properties of the document. Read-only.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.3]
         */
        readonly properties: Word.DocumentProperties;
        /**
         * Gets the collection of section objects in the document. Read-only.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.3]
         */
        readonly sections: Word.SectionCollection;
        /**
         * Gets the add-in's settings in the document. Read-only.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.4]
         * @beta
         */
        readonly settings: Word.SettingCollection;
        /**
         * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.3]
         */
        readonly saved: boolean;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.DocumentCreatedUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.DocumentCreated): void;
        /**
         * Deletes a bookmark, if it exists, from the document.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.4]
         * @beta
         *
         * @param name - Required. The bookmark name, which is case-insensitive.
         */
        deleteBookmark(name: string): void;
        /**
         * Gets a bookmark's range. Throws an error if the bookmark does not exist.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.4]
         * @beta
         *
         * @param name - Required. The bookmark name, which is case-insensitive.
         */
        getBookmarkRange(name: string): Word.Range;
        /**
         * Gets a bookmark's range. Returns a null object if the bookmark does not exist.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.4]
         * @beta
         *
         * @param name - Required. The bookmark name, which is case-insensitive.
         */
        getBookmarkRangeOrNullObject(name: string): Word.Range;
        /**
         * Opens the document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        open(): void;
        /**
         * Saves the document. This uses the Word default file naming convention if the document has not been saved before.
         *
         * @remarks
         * [Api set: WordApiHiddenDocument 1.3]
         */
        save(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.DocumentCreatedLoadOptions): Word.DocumentCreated;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.DocumentCreated;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.DocumentCreated;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.DocumentCreated;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.DocumentCreated;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.DocumentCreated object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentCreatedData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.DocumentCreatedData;
    }
    /**
     * Represents document properties.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class DocumentProperties extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the collection of custom properties of the document. Read only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly customProperties: Word.CustomPropertyCollection;
        /**
         * Gets the application name of the document. Read only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly applicationName: string;
        /**
         * Gets or sets the author of the document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        author: string;
        /**
         * Gets or sets the category of the document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        category: string;
        /**
         * Gets or sets the comments of the document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        comments: string;
        /**
         * Gets or sets the company of the document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        company: string;
        /**
         * Gets the creation date of the document. Read only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly creationDate: Date;
        /**
         * Gets or sets the format of the document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        format: string;
        /**
         * Gets or sets the keywords of the document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        keywords: string;
        /**
         * Gets the last author of the document. Read only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly lastAuthor: string;
        /**
         * Gets the last print date of the document. Read only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly lastPrintDate: Date;
        /**
         * Gets the last save time of the document. Read only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly lastSaveTime: Date;
        /**
         * Gets or sets the manager of the document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        manager: string;
        /**
         * Gets the revision number of the document. Read only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly revisionNumber: string;
        /**
         * Gets security settings of the document. Read only. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly security: number;
        /**
         * Gets or sets the subject of the document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        subject: string;
        /**
         * Gets the template of the document. Read only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly template: string;
        /**
         * Gets or sets the title of the document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        title: string;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.DocumentPropertiesUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.DocumentProperties): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.DocumentPropertiesLoadOptions): Word.DocumentProperties;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.DocumentProperties;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.DocumentProperties;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.DocumentProperties;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.DocumentProperties;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.DocumentProperties object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentPropertiesData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.DocumentPropertiesData;
    }
    /**
     * Represents a field.
     *
     * @remarks
     * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
     * @beta
     */
    export class Field extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the parent body of the field. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly parentBody: Word.Body;
        /**
         * Gets the content control that contains the field. Throws an error if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         * Gets the content control that contains the field. Returns a null object if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         * Gets the table that contains the field. Throws an error if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly parentTable: Word.Table;
        /**
         * Gets the table cell that contains the field. Throws an error if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly parentTableCell: Word.TableCell;
        /**
         * Gets the table cell that contains the field. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         * Gets the table that contains the field. Returns a null object if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         * Gets the field's code instruction. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly code: string;
        /**
         * Gets the field's result data. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly result: string;
        /**
         * Deletes the field.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        delete(): void;
        /**
         * Gets the next field. Throws an error if this field is the last one.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        getNext(): Word.Field;
        /**
         * Gets the next field. Returns a null object if this field is the last one.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        getNextOrNullObject(): Word.Field;
        /**
         * Gets the whole field as a range.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        getRange(): Word.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.FieldLoadOptions): Word.Field;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Field;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Field;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Field;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Field;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Field object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FieldData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.FieldData;
    }
    /**
     * Contains a collection of {@link Word.Field} objects.
     *
     * @remarks
     * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
     * @beta
     */
    export class FieldCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Field[];
        /**
         * Gets the first field in this collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        getFirst(): Word.Field;
        /**
         * Gets the first field in this collection. Returns a null object if this collection is empty.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        getFirstOrNullObject(): Word.Field;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.FieldCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.FieldCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.FieldCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.FieldCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.FieldCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.FieldCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.FieldCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FieldCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.FieldCollectionData;
    }
    /**
     * Represents a font.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class Font extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        bold: boolean;
        /**
         * Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        color: string;
        /**
         * Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        doubleStrikeThrough: boolean;
        /**
         * Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
         * **Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        highlightColor: string;
        /**
         * Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        italic: boolean;
        /**
         * Gets or sets a value that represents the name of the font.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        name: string;
        /**
         * Gets or sets a value that represents the font size in points.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        size: number;
        /**
         * Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        strikeThrough: boolean;
        /**
         * Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        subscript: boolean;
        /**
         * Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        superscript: boolean;
        /**
         * Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        underline: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.FontUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Font): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.FontLoadOptions): Word.Font;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Font;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Font;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Font;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Font;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Font object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FontData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.FontData;
    }
    /**
     * Represents an inline picture.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class InlinePicture extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the parent paragraph that contains the inline image. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         */
        readonly paragraph: Word.Paragraph;
        /**
         * Gets the content control that contains the inline image. Throws an error if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         * Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         * Gets the table that contains the inline image. Throws an error if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         * Gets the table cell that contains the inline image. Throws an error if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCell: Word.TableCell;
        /**
         * Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         * Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         * Gets or sets a string that represents the alternative text associated with the inline image.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        altTextDescription: string;
        /**
         * Gets or sets a string that contains the title for the inline image.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        altTextTitle: string;
        /**
         * Gets or sets a number that describes the height of the inline image.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        height: number;
        /**
         * Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        hyperlink: string;
        /**
         * Gets the format of the inline image. Read-only.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly imageFormat: Word.ImageFormat | "Unsupported" | "Undefined" | "Bmp" | "Jpeg" | "Gif" | "Tiff" | "Png" | "Icon" | "Exif" | "Wmf" | "Emf" | "Pict" | "Pdf" | "Svg";
        /**
         * Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        lockAspectRatio: boolean;
        /**
         * Gets or sets a number that describes the width of the inline image.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        width: number;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.InlinePictureUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.InlinePicture): void;
        /**
         * Deletes the inline picture from the document.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         */
        delete(): void;
        /**
         * Gets the base64 encoded string representation of the inline image.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getBase64ImageSrc(): OfficeExtension.ClientResult<string>;
        /**
         * Gets the next inline image. Throws an error if this inline image is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.InlinePicture;
        /**
         * Gets the next inline image. Returns a null object if this inline image is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.InlinePicture;
        /**
         * Gets the picture, or the starting or ending point of the picture, as a range.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location must be 'Whole', 'Start', or 'End'.
         */
        getRange(rangeLocation?: Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | "Whole" | "Start" | "End"): Word.Range;
        /**
         * Inserts a break at the specified location in the main document.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param breakType - Required. The break type to add.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): void;
        /**
         * Wraps the inline picture with a rich text content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         * Inserts a document at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
        /**
         * Inserts HTML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
        /**
         * Inserts an inline picture at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value must be 'Replace', 'Before', or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.before | Word.InsertLocation.after | "Replace" | "Before" | "After"): Word.InlinePicture;
        /**
         * Inserts OOXML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Paragraph;
        /**
         * Inserts text at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
        /**
         * Selects the inline picture. This causes Word to scroll to the selection.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param selectionMode - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects the inline picture. This causes Word to scroll to the selection.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param selectionModeString - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.InlinePictureLoadOptions): Word.InlinePicture;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.InlinePicture;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.InlinePicture;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.InlinePicture;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.InlinePicture;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.InlinePicture object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.InlinePictureData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.InlinePictureData;
    }
    /**
     * Contains a collection of {@link Word.InlinePicture} objects.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class InlinePictureCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.InlinePicture[];
        /**
         * Gets the first inline image in this collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.InlinePicture;
        /**
         * Gets the first inline image in this collection. Returns a null object if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.InlinePicture;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.InlinePictureCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.InlinePictureCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.InlinePictureCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.InlinePictureCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.InlinePictureCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.InlinePictureCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.InlinePictureCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.InlinePictureCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.InlinePictureCollectionData;
    }
    /**
     * Contains a collection of {@link Word.Paragraph} objects.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class List extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets paragraphs in the list. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly paragraphs: Word.ParagraphCollection;
        /**
         * Gets the list's id.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly id: number;
        /**
         * Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly levelExistences: boolean[];
        /**
         * Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly levelTypes: Word.ListLevelType[];
        /**
         * Gets the font of the bullet, number, or picture at the specified level in the list.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param level - Required. The level in the list.
         */
        getLevelFont(level: number): Word.Font;
        /**
         * Gets the paragraphs that occur at the specified level in the list.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         */
        getLevelParagraphs(level: number): Word.ParagraphCollection;
        /**
         * Gets the base64 encoded string representation of the picture at the specified level in the list.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param level - Required. The level in the list.
         */
        getLevelPicture(level: number): OfficeExtension.ClientResult<string>;
        /**
         * Gets the bullet, number, or picture at the specified level as a string.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         */
        getLevelString(level: number): OfficeExtension.ClientResult<string>;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value must be 'Start', 'End', 'Before', or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"): Word.Paragraph;
        /**
         * Resets the font of the bullet, number, or picture at the specified level in the list.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param level - Required. The level in the list.
         * @param resetFontName - Optional. Indicates whether to reset the font name. Default is false that indicates the font name is kept unchanged.
         */
        resetLevelFont(level: number, resetFontName?: boolean): void;
        /**
         * Sets the alignment of the bullet, number, or picture at the specified level in the list.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param alignment - Required. The level alignment that must be 'Left', 'Centered', or 'Right'.
         */
        setLevelAlignment(level: number, alignment: Word.Alignment): void;
        /**
         * Sets the alignment of the bullet, number, or picture at the specified level in the list.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param alignmentString - Required. The level alignment that must be 'Left', 'Centered', or 'Right'.
         */
        setLevelAlignment(level: number, alignmentString: "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"): void;
        /**
         * Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param listBullet - Required. The bullet.
         * @param charCode - Optional. The bullet character's code value. Used only if the bullet is 'Custom'.
         * @param fontName - Optional. The bullet's font name. Used only if the bullet is 'Custom'.
         */
        setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string): void;
        /**
         * Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param listBulletString - Required. The bullet.
         * @param charCode - Optional. The bullet character's code value. Used only if the bullet is 'Custom'.
         * @param fontName - Optional. The bullet's font name. Used only if the bullet is 'Custom'.
         */
        setLevelBullet(level: number, listBulletString: "Custom" | "Solid" | "Hollow" | "Square" | "Diamonds" | "Arrow" | "Checkmark", charCode?: number, fontName?: string): void;
        /**
         * Sets the two indents of the specified level in the list.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param textIndent - Required. The text indent in points. It is the same as paragraph left indent.
         * @param bulletNumberPictureIndent - Required. The relative indent, in points, of the bullet, number, or picture. It is the same as paragraph first line indent.
         */
        setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number): void;
        /**
         * Sets the numbering format at the specified level in the list.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param listNumbering - Required. The ordinal format.
         * @param formatString - Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.
         */
        setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string | number>): void;
        /**
         * Sets the numbering format at the specified level in the list.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param listNumberingString - Required. The ordinal format.
         * @param formatString - Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.
         */
        setLevelNumbering(level: number, listNumberingString: "None" | "Arabic" | "UpperRoman" | "LowerRoman" | "UpperLetter" | "LowerLetter", formatString?: Array<string | number>): void;
        /**
         * Sets the picture at the specified level in the list.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param level - Required. The level in the list.
         * @param base64EncodedImage - Optional. The base64 encoded image to be set. If not given, the default picture is set.
         */
        setLevelPicture(level: number, base64EncodedImage?: string): void;
        /**
         * Sets the starting number at the specified level in the list. Default value is 1.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param level - Required. The level in the list.
         * @param startingNumber - Required. The number to start with.
         */
        setLevelStartingNumber(level: number, startingNumber: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.ListLoadOptions): Word.List;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.List;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.List;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.List;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.List;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.List object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.ListData;
    }
    /**
     * Contains a collection of {@link Word.List} objects.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class ListCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.List[];
        /**
         * Gets a list by its identifier. Throws an error if there isn't a list with the identifier in this collection.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param id - Required. A list identifier.
         */
        getById(id: number): Word.List;
        /**
         * Gets a list by its identifier. Returns a null object if there isn't a list with the identifier in this collection.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param id - Required. A list identifier.
         */
        getByIdOrNullObject(id: number): Word.List;
        /**
         * Gets the first list in this collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.List;
        /**
         * Gets the first list in this collection. Returns a null object if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.List;
        /**
         * Gets a list object by its ID.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param id - The list's ID.
         */
        getItem(id: number): Word.List;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.ListCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.ListCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.ListCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.ListCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.ListCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.ListCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.ListCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.ListCollectionData;
    }
    /**
     * Represents the paragraph list item format.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class ListItem extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets or sets the level of the item in the list.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        level: number;
        /**
         * Gets the list item bullet, number, or picture as a string. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly listString: string;
        /**
         * Gets the list item order number in relation to its siblings. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly siblingIndex: number;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ListItemUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.ListItem): void;
        /**
         * Gets the list item parent, or the closest ancestor if the parent does not exist. Throws an error if the list item has no ancestor.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param parentOnly - Optional. Specifies only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.
         */
        getAncestor(parentOnly?: boolean): Word.Paragraph;
        /**
         * Gets the list item parent, or the closest ancestor if the parent does not exist. Returns a null object if the list item has no ancestor.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param parentOnly - Optional. Specifies only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.
         */
        getAncestorOrNullObject(parentOnly?: boolean): Word.Paragraph;
        /**
         * Gets all descendant list items of the list item.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param directChildrenOnly - Optional. Specifies only the list item's direct children will be returned. The default is false that indicates to get all descendant items.
         */
        getDescendants(directChildrenOnly?: boolean): Word.ParagraphCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.ListItemLoadOptions): Word.ListItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.ListItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.ListItem;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.ListItem;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.ListItem;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.ListItem object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListItemData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.ListItemData;
    }
    /**
     * Represents a footnote or endnote.
     *
     * @remarks
     * [Api set: WordApiOnline 1.1]
     */
    export class NoteItem extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the body object of the note item. It's the portion of the text within the footnote or endnote.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly body: Word.Body;
        /**
         * Represents a footnote or endnote reference in the main document.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly reference: Word.Range;
        /**
         * Represents the note item type: footnote or endnote.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly type: Word.NoteItemType | "Footnote" | "Endnote";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.NoteItemUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.NoteItem): void;
        /**
         * Deletes the note item.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        delete(): void;
        /**
         * Gets the next note item of the same type. Throws an error if this note item is the last one.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getNext(): Word.NoteItem;
        /**
         * Gets the next note item of the same type. Returns a null object if this note item is the last one.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getNextOrNullObject(): Word.NoteItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.NoteItemLoadOptions): Word.NoteItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.NoteItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.NoteItem;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.NoteItem;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.NoteItem;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.NoteItem object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.NoteItemData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.NoteItemData;
    }
    /**
     * Contains a collection of {@link Word.NoteItem} objects.
     *
     * @remarks
     * [Api set: WordApiOnline 1.1]
     */
    export class NoteItemCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.NoteItem[];
        /**
         * Gets the first note item in this collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getFirst(): Word.NoteItem;
        /**
         * Gets the first note item in this collection. Returns a null object if this collection is empty.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getFirstOrNullObject(): Word.NoteItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.NoteItemCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.NoteItemCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.NoteItemCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.NoteItemCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.NoteItemCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.NoteItemCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.NoteItemCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.NoteItemCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.NoteItemCollectionData;
    }
    /**
     * Represents a single paragraph in a selection, range, content control, or document body.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class Paragraph extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the collection of content control objects in the paragraph. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         * Gets the collection of endnotes in the paragraph. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly endnotes: Word.NoteItemCollection;
        /**
         * Gets the collection of fields in the paragraph. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly fields: Word.FieldCollection;
        /**
         * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        /**
         * Gets the collection of footnotes in the paragraph. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly footnotes: Word.NoteItemCollection;
        /**
         * Gets the collection of InlinePicture objects in the paragraph. The collection does not include floating images. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        /**
         * Gets the List to which this paragraph belongs. Throws an error if the paragraph is not in a list. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly list: Word.List;
        /**
         * Gets the ListItem for the paragraph. Throws an error if the paragraph is not part of a list. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly listItem: Word.ListItem;
        /**
         * Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly listItemOrNullObject: Word.ListItem;
        /**
         * Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly listOrNullObject: Word.List;
        /**
         * Gets the parent body of the paragraph. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentBody: Word.Body;
        /**
         * Gets the content control that contains the paragraph. Throws an error if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         * Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         * Gets the table that contains the paragraph. Throws an error if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         * Gets the table cell that contains the paragraph. Throws an error if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCell: Word.TableCell;
        /**
         * Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         * Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         * Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        alignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
        /**
         * Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        firstLineIndent: number;
        /**
         * Indicates the paragraph is the last one inside its parent body. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly isLastParagraph: boolean;
        /**
         * Checks whether the paragraph is a list item. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly isListItem: boolean;
        /**
         * Gets or sets the left indent value, in points, for the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        leftIndent: number;
        /**
         * Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        lineSpacing: number;
        /**
         * Gets or sets the amount of spacing, in grid lines, after the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        lineUnitAfter: number;
        /**
         * Gets or sets the amount of spacing, in grid lines, before the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        lineUnitBefore: number;
        /**
         * Gets or sets the outline level for the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        outlineLevel: number;
        /**
         * Gets or sets the right indent value, in points, for the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        rightIndent: number;
        /**
         * Gets or sets the spacing, in points, after the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        spaceAfter: number;
        /**
         * Gets or sets the spacing, in points, before the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        spaceBefore: number;
        /**
         * Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         * Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        /**
         * Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly tableNestingLevel: number;
        /**
         * Gets the text of the paragraph. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ParagraphUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Paragraph): void;
        /**
         * Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param listId - Required. The ID of an existing list.
         * @param level - Required. The level in the list.
         */
        attachToList(listId: number, level: number): Word.List;
        /**
         * Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         * Deletes the paragraph and its content from the document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        delete(): void;
        /**
         * Moves this paragraph out of its list, if the paragraph is a list item.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        detachFromList(): void;
        /**
         * Gets comments associated with the paragraph.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        getComments(): Word.CommentCollection;
        /**
         * Gets an HTML representation of the paragraph object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Paragraph.getOoxml()` and convert the returned XML to HTML.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         * Gets the next paragraph. Throws an error if the paragraph is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.Paragraph;
        /**
         * Gets the next paragraph. Returns a null object if the paragraph is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.Paragraph;
        /**
         * Gets the Office Open XML (OOXML) representation of the paragraph object.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         * Gets the previous paragraph. Throws an error if the paragraph is the first one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getPrevious(): Word.Paragraph;
        /**
         * Gets the previous paragraph. Returns a null object if the paragraph is the first one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getPreviousOrNullObject(): Word.Paragraph;
        /**
         * Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location must be 'Whole', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | Word.RangeLocation.after | Word.RangeLocation.content | "Whole" | "Start" | "End" | "After" | "Content"): Word.Range;
        /**
         * Gets reviewed text based on ChangeTrackingVersion selection.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param changeTrackingVersion - Optional. The value must be 'Original' or 'Current'. The default is 'Current'.
         */
        getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion): OfficeExtension.ClientResult<string>;
        /**
         * Gets reviewed text based on ChangeTrackingVersion selection.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param changeTrackingVersionString - Optional. The value must be 'Original' or 'Current'. The default is 'Current'.
         */
        getReviewedText(changeTrackingVersionString?: "Original" | "Current"): OfficeExtension.ClientResult<string>;
        /**
         * Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        getTextRanges(endingMarks: string[], trimSpacing?: boolean): Word.RangeCollection;
        /**
         * Inserts a break at the specified location in the main document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add to the document.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): void;
        /**
         * Wraps the paragraph object with a rich text content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         * Inserts a document into the paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts HTML into the paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the paragraph.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts a picture into the paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.InlinePicture;
        /**
         * Inserts OOXML into the paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in the paragraph.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Paragraph;
        /**
         * Inserts a table with the specified number of rows and columns.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After", values?: string[][]): Word.Table;
        /**
         * Inserts text into the paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', or 'End'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Performs a search with the specified SearchOptions on the scope of the paragraph object. The search results are a collection of range objects.
         *
         * @remarks
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
         * Selects and navigates the Word UI to the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects and navigates the Word UI to the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionModeString - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        /**
         * Splits the paragraph into child ranges by using delimiters.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param delimiters - Required. The delimiters as an array of strings.
         * @param trimDelimiters - Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        /**
         * Starts a new list with this paragraph. Fails if the paragraph is already a list item.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        startNewList(): Word.List;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.ParagraphLoadOptions): Word.Paragraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Paragraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Paragraph;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Paragraph;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Paragraph;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Paragraph object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ParagraphData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.ParagraphData;
    }
    /**
     * Contains a collection of {@link Word.Paragraph} objects.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class ParagraphCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Paragraph[];
        /**
         * Gets the first paragraph in this collection. Throws an error if the collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Paragraph;
        /**
         * Gets the first paragraph in this collection. Returns a null object if the collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.Paragraph;
        /**
         * Gets the last paragraph in this collection. Throws an error if the collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getLast(): Word.Paragraph;
        /**
         * Gets the last paragraph in this collection. Returns a null object if the collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getLastOrNullObject(): Word.Paragraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.ParagraphCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.ParagraphCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.ParagraphCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.ParagraphCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.ParagraphCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.ParagraphCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.ParagraphCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ParagraphCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.ParagraphCollectionData;
    }
    /**
     * Represents a contiguous area in a document.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class Range extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the collection of content control objects in the range. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        /**
         * Gets the collection of endnotes in the range. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly endnotes: Word.NoteItemCollection;
        /**
         * Gets the collection of field objects in the range. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly fields: Word.FieldCollection;
        /**
         * Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        /**
         * Gets the collection of footnotes in the range. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly footnotes: Word.NoteItemCollection;
        /**
         * Gets the collection of inline picture objects in the range. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        /**
         * Gets the collection of list objects in the range. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly lists: Word.ListCollection;
        /**
         * Gets the collection of paragraph objects in the range. Read-only. **Important**: For requirement sets 1.1 and 1.2, paragraphs in tables wholly contained within this range are not returned. From requirement set 1.3, paragraphs in such tables are also returned.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly paragraphs: Word.ParagraphCollection;
        /**
         * Gets the parent body of the range. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentBody: Word.Body;
        /**
         * Gets the content control that contains the range. Throws an error if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         * Gets the content control that contains the range. Returns a null object if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         * Gets the table that contains the range. Throws an error if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         * Gets the table cell that contains the range. Throws an error if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCell: Word.TableCell;
        /**
         * Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         * Gets the table that contains the range. Returns a null object if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         * Gets the collection of table objects in the range. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly tables: Word.TableCollection;
        /**
         * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        hyperlink: string;
        /**
         * Checks whether the range length is zero. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly isEmpty: boolean;
        /**
         * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        style: string;
        /**
         * Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        /**
         * Gets the text of the range. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.RangeUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Range): void;
        /**
         * Clears the contents of the range object. The user can perform the undo operation on the cleared content.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         * Compares this range's location with another range's location.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param range - Required. The range to compare with this range.
         */
        compareLocationWith(range: Word.Range): OfficeExtension.ClientResult<Word.LocationRelation>;
        /**
         * Deletes the range and its content from the document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        delete(): void;
        /**
         * Returns a new range that extends from this range in either direction to cover another range. This range is not changed. Throws an error if the two ranges do not have a union.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param range - Required. Another range.
         */
        expandTo(range: Word.Range): Word.Range;
        /**
         * Returns a new range that extends from this range in either direction to cover another range. This range is not changed. Returns a null object if the two ranges do not have a union.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param range - Required. Another range.
         */
        expandToOrNullObject(range: Word.Range): Word.Range;
        /**
         * Gets the names all bookmarks in or overlapping the range. A bookmark is hidden if its name starts with the underscore character.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param includeHidden - Optional. Indicates whether to include hidden bookmarks. Default is false which indicates that the hidden bookmarks are excluded.
         * @param includeAdjacent - Optional. Indicates whether to include bookmarks that are adjacent to the range. Default is false which indicates that the adjacent bookmarks are excluded.
         */
        getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean): OfficeExtension.ClientResult<string[]>;
        /**
         * Gets comments associated with the range.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         * @returns
         */
        getComments(): Word.CommentCollection;
        /**
         * Gets an HTML representation of the range object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Range.getOoxml()` and convert the returned XML to HTML.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         * Gets hyperlink child ranges within the range.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getHyperlinkRanges(): Word.RangeCollection;
        /**
         * Gets the next text range by using punctuation marks and/or other ending marks. Throws an error if this text range is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.
         */
        getNextTextRange(endingMarks: string[], trimSpacing?: boolean): Word.Range;
        /**
         * Gets the next text range by using punctuation marks and/or other ending marks. Returns a null object if this text range is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.
         */
        getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean): Word.Range;
        /**
         * Gets the OOXML representation of the range object.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        /**
         * Clones the range, or gets the starting or ending point of the range as a new range.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location must be 'Whole', 'Start', 'End', 'After', or 'Content'.
         */
        getRange(rangeLocation?: Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | Word.RangeLocation.after | Word.RangeLocation.content | "Whole" | "Start" | "End" | "After" | "Content"): Word.Range;
        /**
         * Gets reviewed text based on ChangeTrackingVersion selection.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param changeTrackingVersion - Optional. The value must be 'Original' or 'Current'. The default is 'Current'.
         */
        getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion): OfficeExtension.ClientResult<string>;
        /**
         * Gets reviewed text based on ChangeTrackingVersion selection.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param changeTrackingVersionString - Optional. The value must be 'Original' or 'Current'. The default is 'Current'.
         */
        getReviewedText(changeTrackingVersionString?: "Original" | "Current"): OfficeExtension.ClientResult<string>;
        /**
         * Gets the text child ranges in the range by using punctuation marks and/or other ending marks.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param endingMarks - Required. The punctuation marks and/or other ending marks as an array of strings.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        getTextRanges(endingMarks: string[], trimSpacing?: boolean): Word.RangeCollection;
        /**
         * Inserts a bookmark on the range. If a bookmark of the same name exists somewhere, it is deleted first.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param name - Required. The bookmark name, which is case-insensitive. If the name starts with an underscore character, the bookmark is an hidden one.
         */
        insertBookmark(name: string): void;
        /**
         * Inserts a break at the specified location in the main document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): void;
        /**
         * Insert a comment on the range.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param commentText - Required. The comment text to be inserted.
         * @returns comment object
         */
        insertComment(commentText: string): Word.Comment;
        /**
         * Wraps the range object with a rich text content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         * Inserts an endnote. The endnote reference is placed after the range.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param insertText - Optional. Text to be inserted into the endnote body. The default is "".
         */
        insertEndnote(insertText?: string): Word.NoteItem;
        /**
         * Inserts a document at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
        /**
         * Inserts a footnote. The footnote reference is placed after the range.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         *
         * @param insertText - Optional. Text to be inserted into the footnote body. The default is "".
         */
        insertFootnote(insertText?: string): Word.NoteItem;
        /**
         * Inserts HTML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
        /**
         * Inserts a picture at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.InlinePicture;
        /**
         * Inserts OOXML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Paragraph;
        /**
         * Inserts a table with the specified number of rows and columns.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After", values?: string[][]): Word.Table;
        /**
         * Inserts text at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
        /**
         * Returns a new range as the intersection of this range with another range. This range is not changed. Throws an error if the two ranges are not overlapped or adjacent.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param range - Required. Another range.
         */
        intersectWith(range: Word.Range): Word.Range;
        /**
         * Returns a new range as the intersection of this range with another range. This range is not changed. Returns a null object if the two ranges are not overlapped or adjacent.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param range - Required. Another range.
         */
        intersectWithOrNullObject(range: Word.Range): Word.Range;
        /**
         * Performs a search with the specified SearchOptions on the scope of the range object. The search results are a collection of range objects.
         *
         * @remarks
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
         * Selects and navigates the Word UI to the range.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects and navigates the Word UI to the range.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionModeString - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        /**
         * Splits the range into child ranges by using delimiters.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param delimiters - Required. The delimiters as an array of strings.
         * @param multiParagraphs - Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.
         * @param trimDelimiters - Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
         * @param trimSpacing - Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.
         */
        split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.RangeLoadOptions): Word.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Range;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Range;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Range;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Range object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.RangeData;
    }
    /**
     * Contains a collection of {@link Word.Range} objects.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class RangeCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Range[];
        /**
         * Gets the first range in this collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Range;
        /**
         * Gets the first range in this collection. Returns a null object if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.RangeCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.RangeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.RangeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.RangeCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.RangeCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.RangeCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.RangeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.RangeCollectionData;
    }
    /**
     * Specifies the options to be included in a search operation.
     *
     * To learn more about how to use search options in the Word JavaScript APIs, read {@link https://docs.microsoft.com/office/dev/add-ins/word/search-option-guidance | Use search options to find text in your Word add-in}.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class SearchOptions extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        ignorePunct: boolean;
        /**
         * Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        ignoreSpace: boolean;
        /**
         * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        matchCase: boolean;
        /**
         * Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        matchPrefix: boolean;
        /**
         * Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        matchSuffix: boolean;
        /**
         * Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        matchWholeWord: boolean;
        /**
         * Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        matchWildcards: boolean;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.SearchOptionsUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.SearchOptions): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.SearchOptionsLoadOptions): Word.SearchOptions;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.SearchOptions;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.SearchOptions;
        /**
         * Create a new instance of Word.SearchOptions object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): Word.SearchOptions;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.SearchOptions object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SearchOptionsData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.SearchOptionsData;
    }
    /**
     * Represents a section in a Word document.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class Section extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the body object of the section. This does not include the header/footer and other section metadata. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly body: Word.Body;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.SectionUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Section): void;
        /**
         * Gets one of the section's footers.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param type - Required. The type of footer to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.
         */
        getFooter(type: Word.HeaderFooterType): Word.Body;
        /**
         * Gets one of the section's footers.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param typeString - Required. The type of footer to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.
         */
        getFooter(typeString: "Primary" | "FirstPage" | "EvenPages"): Word.Body;
        /**
         * Gets one of the section's headers.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param type - Required. The type of header to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.
         */
        getHeader(type: Word.HeaderFooterType): Word.Body;
        /**
         * Gets one of the section's headers.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param typeString - Required. The type of header to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.
         */
        getHeader(typeString: "Primary" | "FirstPage" | "EvenPages"): Word.Body;
        /**
         * Gets the next section. Throws an error if this section is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.Section;
        /**
         * Gets the next section. Returns a null object if this section is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.SectionLoadOptions): Word.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Section;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Section;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Section;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Section object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.SectionData;
    }
    /**
     * Contains the collection of the document's {@link Word.Section} objects.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class SectionCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Section[];
        /**
         * Gets the first section in this collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Section;
        /**
         * Gets the first section in this collection. Returns a null object if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.SectionCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.SectionCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.SectionCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.SectionCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.SectionCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.SectionCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.SectionCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.SectionCollectionData;
    }
    /**
     * Represents a setting of the add-in.
     *
     * @remarks
     * [Api set: WordApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class Setting extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the key of the setting. Read only.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly key: string;
        /**
         * Gets or sets the value of the setting.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        value: any;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.SettingUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Setting): void;
        /**
         * Deletes the setting.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        delete(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.SettingLoadOptions): Word.Setting;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Setting;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Setting;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Setting;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Setting;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Setting object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SettingData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.SettingData;
    }
    /**
     * Contains the collection of {@link Word.Setting} objects.
     *
     * @remarks
     * [Api set: WordApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class SettingCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Setting[];
        /**
         * Creates a new setting or sets an existing setting.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param key - Required. The setting's key, which is case-sensitive.
         * @param value - Required. The setting's value.
         */
        add(key: string, value: any): Word.Setting;
        /**
         * Deletes all settings in this add-in.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        deleteAll(): void;
        /**
         * Gets the count of settings.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a setting object by its key, which is case-sensitive. Throws an error if the setting does not exist.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param key - The key that identifies the setting object.
         */
        getItem(key: string): Word.Setting;
        /**
         * Gets a setting object by its key, which is case-sensitive. Returns a null object if the setting does not exist.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param key - Required. The key that identifies the setting object.
         */
        getItemOrNullObject(key: string): Word.Setting;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.SettingCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.SettingCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.SettingCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.SettingCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.SettingCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.SettingCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.SettingCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SettingCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.SettingCollectionData;
    }
    /**
     * Represents a table in a Word document.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class Table extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the collection of endnotes in the table. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly endnotes: Word.NoteItemCollection;
        /**
         * Gets the collection of field objects in the table. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly fields: Word.FieldCollection;
        /**
         * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly font: Word.Font;
        /**
         * Gets the collection of footnotes in the table. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly footnotes: Word.NoteItemCollection;
        /**
         * Gets the parent body of the table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentBody: Word.Body;
        /**
         * Gets the content control that contains the table. Throws an error if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControl: Word.ContentControl;
        /**
         * Gets the content control that contains the table. Returns a null object if there isn't a parent content control. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentContentControlOrNullObject: Word.ContentControl;
        /**
         * Gets the table that contains this table. Throws an error if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         * Gets the table cell that contains this table. Throws an error if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCell: Word.TableCell;
        /**
         * Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableCellOrNullObject: Word.TableCell;
        /**
         * Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTableOrNullObject: Word.Table;
        /**
         * Gets all of the table rows. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly rows: Word.TableRowCollection;
        /**
         * Gets the child tables nested one level deeper. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly tables: Word.TableCollection;
        /**
         * Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        alignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
        /**
         * Gets and sets the number of header rows.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        headerRowCount: number;
        /**
         * Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
        /**
         * Indicates whether all of the table rows are uniform. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly isUniform: boolean;
        /**
         * Gets the nesting level of the table. Top-level tables have level 1. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly nestingLevel: number;
        /**
         * Gets the number of rows in the table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly rowCount: number;
        /**
         * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        shadingColor: string;
        /**
         * Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        style: string;
        /**
         * Gets and sets whether the table has banded columns.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        styleBandedColumns: boolean;
        /**
         * Gets and sets whether the table has banded rows.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        styleBandedRows: boolean;
        /**
         * Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        styleBuiltIn: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        /**
         * Gets and sets whether the table has a first column with a special style.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        styleFirstColumn: boolean;
        /**
         * Gets and sets whether the table has a last column with a special style.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        styleLastColumn: boolean;
        /**
         * Gets and sets whether the table has a total (last) row with a special style.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        styleTotalRow: boolean;
        /**
         * Gets and sets the text values in the table, as a 2D Javascript array.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        values: string[][];
        /**
         * Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        /**
         * Gets and sets the width of the table in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        width: number;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.TableUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Table): void;
        /**
         * Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. It must be 'Start' or 'End', corresponding to the appropriate side of the table.
         * @param columnCount - Required. Number of columns to add.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        addColumns(insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End", columnCount: number, values?: string[][]): void;
        /**
         * Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. It must be 'Start' or 'End'.
         * @param rowCount - Required. Number of rows to add.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        addRows(insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End", rowCount: number, values?: string[][]): Word.TableRowCollection;
        /**
         * Autofits the table columns to the width of the window.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        autoFitWindow(): void;
        /**
         * Clears the contents of the table.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        clear(): void;
        /**
         * Deletes the entire table.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        delete(): void;
        /**
         * Deletes specific columns. This is applicable to uniform tables.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param columnIndex - Required. The first column to delete.
         * @param columnCount - Optional. The number of columns to delete. Default 1.
         */
        deleteColumns(columnIndex: number, columnCount?: number): void;
        /**
         * Deletes specific rows.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rowIndex - Required. The first row to delete.
         * @param rowCount - Optional. The number of rows to delete. Default 1.
         */
        deleteRows(rowIndex: number, rowCount?: number): void;
        /**
         * Distributes the column widths evenly. This is applicable to uniform tables.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        distributeColumns(): void;
        /**
         * Gets the border style for the specified border.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param borderLocation - Required. The border location.
         */
        getBorder(borderLocation: Word.BorderLocation): Word.TableBorder;
        /**
         * Gets the border style for the specified border.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param borderLocationString - Required. The border location.
         */
        getBorder(borderLocationString: "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"): Word.TableBorder;
        /**
         * Gets the table cell at a specified row and column. Throws an error if the specified table cell does not exist.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rowIndex - Required. The index of the row.
         * @param cellIndex - Required. The index of the cell in the row.
         */
        getCell(rowIndex: number, cellIndex: number): Word.TableCell;
        /**
         * Gets the table cell at a specified row and column. Returns a null object if the specified table cell does not exist.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rowIndex - Required. The index of the row.
         * @param cellIndex - Required. The index of the cell in the row.
         */
        getCellOrNullObject(rowIndex: number, cellIndex: number): Word.TableCell;
        /**
         * Gets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocation: Word.CellPaddingLocation): OfficeExtension.ClientResult<number>;
        /**
         * Gets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocationString - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocationString: "Top" | "Left" | "Bottom" | "Right"): OfficeExtension.ClientResult<number>;
        /**
         * Gets the next table. Throws an error if this table is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.Table;
        /**
         * Gets the next table. Returns a null object if this table is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.Table;
        /**
         * Gets the paragraph after the table. Throws an error if there isn't a paragraph after the table.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getParagraphAfter(): Word.Paragraph;
        /**
         * Gets the paragraph after the table. Returns a null object if there isn't a paragraph after the table.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getParagraphAfterOrNullObject(): Word.Paragraph;
        /**
         * Gets the paragraph before the table. Throws an error if there isn't a paragraph before the table.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getParagraphBefore(): Word.Paragraph;
        /**
         * Gets the paragraph before the table. Returns a null object if there isn't a paragraph before the table.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getParagraphBeforeOrNullObject(): Word.Paragraph;
        /**
         * Gets the range that contains this table, or the range at the start or end of the table.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rangeLocation - Optional. The range location must be 'Whole', 'Start', 'End', or 'After'.
         */
        getRange(rangeLocation?: Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | Word.RangeLocation.after | "Whole" | "Start" | "End" | "After"): Word.Range;
        /**
         * Inserts a content control on the table.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        insertContentControl(): Word.ContentControl;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Paragraph;
        /**
         * Inserts a table with the specified number of rows and columns.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param insertLocation - Required. The value must be 'Before' or 'After'.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After", values?: string[][]): Word.Table;
        /**
         * Merges the cells bounded inclusively by a first and last cell.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param topRow - Required. The row of the first cell
         * @param firstCell - Required. The index of the first cell in its row
         * @param bottomRow - Required. The row of the last cell
         * @param lastCell - Required. The index of the last cell in its row
         */
        mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number): Word.TableCell;
        /**
         * Performs a search with the specified SearchOptions on the scope of the table object. The search results are a collection of range objects.
         *
         * @remarks
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
         * Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param selectionMode - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param selectionModeString - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        /**
         * Sets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number): void;
        /**
         * Sets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocationString - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocationString: "Top" | "Left" | "Bottom" | "Right", cellPadding: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.TableLoadOptions): Word.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Table;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Table;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Table;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Table object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.TableData;
    }
    /**
     * Contains the collection of the document's Table objects.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class TableCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.Table[];
        /**
         * Gets the first table in this collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.Table;
        /**
         * Gets the first table in this collection. Returns a null object if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.TableCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.TableCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.TableCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.TableCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.TableCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.TableCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.TableCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.TableCollectionData;
    }
    /**
     * Represents a row in a Word document.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class TableRow extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets cells. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly cells: Word.TableCellCollection;
        /**
         * Gets the collection of endnotes in the table row. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly endnotes: Word.NoteItemCollection;
        /**
         * Gets the collection of field objects in the table row. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly fields: Word.FieldCollection;
        /**
         * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly font: Word.Font;
        /**
         * Gets the collection of footnotes in the table row. Read-only.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        readonly footnotes: Word.NoteItemCollection;
        /**
         * Gets parent table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         * Gets the number of cells in the row. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly cellCount: number;
        /**
         * Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
        /**
         * Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly isHeader: boolean;
        /**
         * Gets and sets the preferred height of the row in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        preferredHeight: number;
        /**
         * Gets the index of the row in its parent table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly rowIndex: number;
        /**
         * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        shadingColor: string;
        /**
         * Gets and sets the text values in the row, as a 2D Javascript array.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        values: string[][];
        /**
         * Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.TableRowUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.TableRow): void;
        /**
         * Clears the contents of the row.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        clear(): void;
        /**
         * Deletes the entire row.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        delete(): void;
        /**
         * Gets the border style of the cells in the row.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param borderLocation - Required. The border location.
         */
        getBorder(borderLocation: Word.BorderLocation): Word.TableBorder;
        /**
         * Gets the border style of the cells in the row.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param borderLocationString - Required. The border location.
         */
        getBorder(borderLocationString: "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"): Word.TableBorder;
        /**
         * Gets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocation: Word.CellPaddingLocation): OfficeExtension.ClientResult<number>;
        /**
         * Gets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocationString - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocationString: "Top" | "Left" | "Bottom" | "Right"): OfficeExtension.ClientResult<number>;
        /**
         * Gets the next row. Throws an error if this row is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.TableRow;
        /**
         * Gets the next row. Returns a null object if this row is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.TableRow;
        /**
         * Inserts a content control on the row.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        insertContentControl(): Word.ContentControl;
        /**
         * Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. Where the new rows should be inserted, relative to the current row. It must be 'Before' or 'After'.
         * @param rowCount - Required. Number of rows to add
         * @param values - Optional. Strings to insert in the new rows, specified as a 2D array. The number of cells in each row must not exceed the number of cells in the existing row.
         */
        insertRows(insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After", rowCount: number, values?: string[][]): Word.TableRowCollection;
        /**
         * Merges the row into one cell.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        merge(): Word.TableCell;
        /**
         * Performs a search with the specified SearchOptions on the scope of the row. The search results are a collection of range objects.
         *
         * @remarks
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
         * Selects the row and navigates the Word UI to it.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param selectionMode - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects the row and navigates the Word UI to it.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param selectionModeString - Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        /**
         * Sets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number): void;
        /**
         * Sets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocationString - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocationString: "Top" | "Left" | "Bottom" | "Right", cellPadding: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.TableRowLoadOptions): Word.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.TableRow;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.TableRow;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.TableRow;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.TableRow object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableRowData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.TableRowData;
    }
    /**
     * Contains the collection of the document's TableRow objects.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class TableRowCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.TableRow[];
        /**
         * Gets the first row in this collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.TableRow;
        /**
         * Gets the first row in this collection. Returns a null object if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.TableRowCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.TableRowCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.TableRowCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.TableRowCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.TableRowCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.TableRowCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.TableRowCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableRowCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.TableRowCollectionData;
    }
    /**
     * Represents a table cell in a Word document.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class TableCell extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the body object of the cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly body: Word.Body;
        /**
         * Gets the parent row of the cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentRow: Word.TableRow;
        /**
         * Gets the parent table of the cell. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly parentTable: Word.Table;
        /**
         * Gets the index of the cell in its row. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly cellIndex: number;
        /**
         * Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        columnWidth: number;
        /**
         * Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        horizontalAlignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
        /**
         * Gets the index of the cell's row in the table. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly rowIndex: number;
        /**
         * Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        shadingColor: string;
        /**
         * Gets and sets the text of the cell.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        value: string;
        /**
         * Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        verticalAlignment: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        /**
         * Gets the width of the cell in points. Read-only.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        readonly width: number;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.TableCellUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.TableCell): void;
        /**
         * Deletes the column containing this cell. This is applicable to uniform tables.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        deleteColumn(): void;
        /**
         * Deletes the row containing this cell.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        deleteRow(): void;
        /**
         * Gets the border style for the specified border.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param borderLocation - Required. The border location.
         */
        getBorder(borderLocation: Word.BorderLocation): Word.TableBorder;
        /**
         * Gets the border style for the specified border.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param borderLocationString - Required. The border location.
         */
        getBorder(borderLocationString: "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"): Word.TableBorder;
        /**
         * Gets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocation: Word.CellPaddingLocation): OfficeExtension.ClientResult<number>;
        /**
         * Gets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocationString - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         */
        getCellPadding(cellPaddingLocationString: "Top" | "Left" | "Bottom" | "Right"): OfficeExtension.ClientResult<number>;
        /**
         * Gets the next cell. Throws an error if this cell is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNext(): Word.TableCell;
        /**
         * Gets the next cell. Returns a null object if this cell is the last one.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getNextOrNullObject(): Word.TableCell;
        /**
         * Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. It must be 'Before' or 'After'.
         * @param columnCount - Required. Number of columns to add.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertColumns(insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After", columnCount: number, values?: string[][]): void;
        /**
         * Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param insertLocation - Required. It must be 'Before' or 'After'.
         * @param rowCount - Required. Number of rows to add.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertRows(insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After", rowCount: number, values?: string[][]): Word.TableRowCollection;
        /**
         * Sets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocation - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number): void;
        /**
         * Sets cell padding in points.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         *
         * @param cellPaddingLocationString - Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
         * @param cellPadding - Required. The cell padding.
         */
        setCellPadding(cellPaddingLocationString: "Top" | "Left" | "Bottom" | "Right", cellPadding: number): void;
        /**
         * Splits the cell into the specified number of rows and columns.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param rowCount - Required. The number of rows to split into. Must be a divisor of the number of underlying rows.
         * @param columnCount - Required. The number of columns to split into.
         */
        split(rowCount: number, columnCount: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.TableCellLoadOptions): Word.TableCell;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.TableCell;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.TableCell;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.TableCell;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.TableCell;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.TableCell object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableCellData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.TableCellData;
    }
    /**
     * Contains the collection of the document's TableCell objects.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class TableCellCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Word.TableCell[];
        /**
         * Gets the first table cell in this collection. Throws an error if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirst(): Word.TableCell;
        /**
         * Gets the first table cell in this collection. Returns a null object if this collection is empty.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        getFirstOrNullObject(): Word.TableCell;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.TableCellCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.TableCellCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.TableCellCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.TableCellCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.TableCellCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.TableCellCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.TableCellCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableCellCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.TableCellCollectionData;
    }
    /**
     * Specifies the border style.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    export class TableBorder extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets or sets the table border color.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        color: string;
        /**
         * Gets or sets the type of the table border.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        type: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
        /**
         * Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        width: number;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.TableBorderUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.TableBorder): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Word.Interfaces.TableBorderLoadOptions): Word.TableBorder;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.TableBorder;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.TableBorder;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you are using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.TableBorder;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://docs.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.TableBorder;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.TableBorder object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableBorderData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.TableBorderData;
    }
    /**
     * ChangeTracking mode.
     *
     * @remarks
     * [Api set: WordApiOnline 1.1]
     */
    enum ChangeTrackingMode {
        /**
         * ChangeTracking is turned off.
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        off = "Off",
        /**
         * ChangeTracking is turned on for everyone.
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        trackAll = "TrackAll",
        /**
         * Tracking is turned on for my changes only.
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        trackMineOnly = "TrackMineOnly",
    }
    /**
     * Specify the current version or the original version of the text.
     *
     * @remarks
     * [Api set: WordApiOnline 1.1]
     */
    enum ChangeTrackingVersion {
        /**
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        original = "Original",
        /**
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        current = "Current",
    }
    /**
     * Note item type
     *
     * @remarks
     * [Api set: WordApiOnline 1.1]
     */
    enum NoteItemType {
        /**
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        footnote = "Footnote",
        /**
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        endnote = "Endnote",
    }
    /**
     * Provides information about the type of a raised event.
     *
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum EventType {
        /**
         * ContentControlDeleted represents the event that the content control has been deleted.
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        contentControlDeleted = "ContentControlDeleted",
        /**
         * ContentControlSelectionChanged represents the event that the selection in the content control has been changed.
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        contentControlSelectionChanged = "ContentControlSelectionChanged",
        /**
         * ContentControlDataChanged represents the event that the data in the content control have been changed.
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        contentControlDataChanged = "ContentControlDataChanged",
        /**
         * ContentControlAdded represents the event a content control has been added to the document.
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        contentControlAdded = "ContentControlAdded",
    }
    /**
     * Provides information about the content control that raised an event.
     *
     * @remarks
     * [Api set: WordApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export interface ContentControlEventArgs {
        /**
         * The object that raised the event. Load this object to get its properties.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        contentControl: Word.ContentControl;
        /**
         * The event type. See Word.EventType for details.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        eventType: Word.EventType | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded";
    }
    /**
     * Specifies supported content control types and subtypes.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    enum ContentControlType {
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        unknown = "Unknown",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richTextInline = "RichTextInline",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richTextParagraphs = "RichTextParagraphs",
        /**
         * Contains a whole cell.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richTextTableCell = "RichTextTableCell",
        /**
         * Contains a whole row.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richTextTableRow = "RichTextTableRow",
        /**
         * Contains a whole table.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richTextTable = "RichTextTable",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        plainTextInline = "PlainTextInline",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        plainTextParagraph = "PlainTextParagraph",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        picture = "Picture",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        buildingBlockGallery = "BuildingBlockGallery",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        checkBox = "CheckBox",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        comboBox = "ComboBox",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dropDownList = "DropDownList",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        datePicker = "DatePicker",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        repeatingSection = "RepeatingSection",
        /**
         * Identifies a rich text content control.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richText = "RichText",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        plainText = "PlainText",
    }
    /**
     * ContentControl appearance.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     *
     * Content control appearance options are bounding box, tags, or hidden.
     */
    enum ContentControlAppearance {
        /**
         * Represents a content control shown as a shaded rectangle or bounding box (with optional title).
         * @remarks
         * [Api set: WordApi 1.1]
         */
        boundingBox = "BoundingBox",
        /**
         * Represents a content control shown as start and end markers.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        tags = "Tags",
        /**
         * Represents a content control that is not shown.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        hidden = "Hidden",
    }
    /**
     * The supported styles for underline format.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    enum UnderlineType {
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        mixed = "Mixed",
        /**
         * No underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        none = "None",
        /**
         * **Warning**: `hidden` has been deprecated.
         *
         * @deprecated `hidden` is no longer supported.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        hidden = "Hidden",
        /**
         * **Warning**: `dotLine` has been deprecated.
         *
         * @deprecated `dotLine` is no longer supported.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dotLine = "DotLine",
        /**
         * A single underline. This is the default value.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        single = "Single",
        /**
         * Only underline individual words.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        word = "Word",
        /**
         * A double underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        double = "Double",
        /**
         * A single thick underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        thick = "Thick",
        /**
         * A dotted underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dotted = "Dotted",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dottedHeavy = "DottedHeavy",
        /**
         * A single dash underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dashLine = "DashLine",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dashLineHeavy = "DashLineHeavy",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dashLineLong = "DashLineLong",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dashLineLongHeavy = "DashLineLongHeavy",
        /**
         * An alternating dot-dash underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dotDashLine = "DotDashLine",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dotDashLineHeavy = "DotDashLineHeavy",
        /**
         * An alternating dot-dot-dash underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        twoDotDashLine = "TwoDotDashLine",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        twoDotDashLineHeavy = "TwoDotDashLineHeavy",
        /**
         * A single wavy underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        wave = "Wave",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        waveHeavy = "WaveHeavy",
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        waveDouble = "WaveDouble",
    }
    /**
     * Specifies the form of a break.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    enum BreakType {
        /**
         * Page break at the insertion point.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        page = "Page",
        /**
         * **Warning**: `next` has been deprecated. Use `sectionNext` instead.
         *
         * @deprecated Use `sectionNext` instead.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        next = "Next",
        /**
         * Section break on next page.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        sectionNext = "SectionNext",
        /**
         * New section without a corresponding page break.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        sectionContinuous = "SectionContinuous",
        /**
         * Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        sectionEven = "SectionEven",
        /**
         * Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        sectionOdd = "SectionOdd",
        /**
         * Line break.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        line = "Line",
    }
    /**
     * The insertion location types.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     *
     * To be used with an API call, such as `obj.insertSomething(newStuff, location);`.
     * If the location is "Before" or "After", the new content will be outside of the modified object.
     * If the location is "Start" or "End", the new content will be included as part of the modified object.
     */
    enum InsertLocation {
        /**
         * Add content before the contents of the calling object.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        before = "Before",
        /**
         * Add content after the contents of the calling object.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        after = "After",
        /**
         * Prepend content to the contents of the calling object.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        start = "Start",
        /**
         * Append content to the contents of the calling object.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        end = "End",
        /**
         * Replace the contents of the current object.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        replace = "Replace",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.1]
     */
    enum Alignment {
        /**
         * @remarks
         * [Api set: WordApi 1.1]
         */
        mixed = "Mixed",
        /**
         * Unknown alignment.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        unknown = "Unknown",
        /**
         * Alignment to the left.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        left = "Left",
        /**
         * Alignment to the center.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        centered = "Centered",
        /**
         * Alignment to the right.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        right = "Right",
        /**
         * Fully justified alignment.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        justified = "Justified",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.1]
     */
    enum HeaderFooterType {
        /**
         * Returns the header or footer on all pages of a section, but excludes the first page or even pages if they are different.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        primary = "Primary",
        /**
         * Returns the header or footer on the first page of a section.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        firstPage = "FirstPage",
        /**
         * Returns all headers or footers on even-numbered pages of a section.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        evenPages = "EvenPages",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum BodyType {
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        unknown = "Unknown",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        mainDoc = "MainDoc",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        section = "Section",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        header = "Header",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        footer = "Footer",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        tableCell = "TableCell",
        /**
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        footnote = "Footnote",
        /**
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        endnote = "Endnote",
        /**
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        noteItem = "NoteItem",
    }
    /**
     * This enum sets where the cursor (insertion point) in the document is after a selection.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    enum SelectionMode {
        /**
         * The entire range is selected.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        select = "Select",
        /**
         * The cursor is at the beginning of the selection (just before the start of the selected range).
         * @remarks
         * [Api set: WordApi 1.1]
         */
        start = "Start",
        /**
         * The cursor is at the end of the selection (just after the end of the selected range).
         * @remarks
         * [Api set: WordApi 1.1]
         */
        end = "End",
    }
    /**
     * @remarks
     * [Api set: WordApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum ImageFormat {
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        unsupported = "Unsupported",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        undefined = "Undefined",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        bmp = "Bmp",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        jpeg = "Jpeg",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        gif = "Gif",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        tiff = "Tiff",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        png = "Png",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        icon = "Icon",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        exif = "Exif",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        wmf = "Wmf",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        emf = "Emf",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        pict = "Pict",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        pdf = "Pdf",
        /**
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         */
        svg = "Svg",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum RangeLocation {
        /**
         * The object's whole range. If the object is a paragraph content control or table content control, the EOP or Table characters after the content control are also included.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        whole = "Whole",
        /**
         * The starting point of the object. For content control, it is the point after the opening tag.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        start = "Start",
        /**
         * The ending point of the object. For paragraph, it is the point before the EOP. For content control, it is the point before the closing tag.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        end = "End",
        /**
         * For content control only. It is the point before the opening tag.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        before = "Before",
        /**
         * The point after the object. If the object is a paragraph content control or table content control, it is the point after the EOP or Table characters.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        after = "After",
        /**
         * The range between 'Start' and 'End'.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        content = "Content",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum LocationRelation {
        /**
         * Indicates that this instance and the range are in different sub-documents.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        unrelated = "Unrelated",
        /**
         * Indicates that this instance and the range represent the same range.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        equal = "Equal",
        /**
         * Indicates that this instance contains the range and that it shares the same start character. The range does not share the same end character as this instance.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        containsStart = "ContainsStart",
        /**
         * Indicates that this instance contains the range and that it shares the same end character. The range does not share the same start character as this instance.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        containsEnd = "ContainsEnd",
        /**
         * Indicates that this instance contains the range, with the exception of the start and end character of this instance.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        contains = "Contains",
        /**
         * Indicates that this instance is inside the range and that it shares the same start character. The range does not share the same end character as this instance.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        insideStart = "InsideStart",
        /**
         * Indicates that this instance is inside the range and that it shares the same end character. The range does not share the same start character as this instance.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        insideEnd = "InsideEnd",
        /**
         * Indicates that this instance is inside the range. The range does not share the same start and end characters as this instance.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        inside = "Inside",
        /**
         * Indicates that this instance occurs before, and is adjacent to, the range.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        adjacentBefore = "AdjacentBefore",
        /**
         * Indicates that this instance starts before the range and overlaps the range’s first character.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        overlapsBefore = "OverlapsBefore",
        /**
         * Indicates that this instance occurs before the range.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        before = "Before",
        /**
         * Indicates that this instance occurs after, and is adjacent to, the range.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        adjacentAfter = "AdjacentAfter",
        /**
         * Indicates that this instance starts inside the range and overlaps the range’s last character.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        overlapsAfter = "OverlapsAfter",
        /**
         * Indicates that this instance occurs after the range.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        after = "After",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum BorderLocation {
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        top = "Top",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        left = "Left",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        bottom = "Bottom",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        right = "Right",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        insideHorizontal = "InsideHorizontal",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        insideVertical = "InsideVertical",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        inside = "Inside",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        outside = "Outside",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        all = "All",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum CellPaddingLocation {
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        top = "Top",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        left = "Left",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        bottom = "Bottom",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        right = "Right",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum BorderType {
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        mixed = "Mixed",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        none = "None",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        single = "Single",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        double = "Double",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        dotted = "Dotted",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        dashed = "Dashed",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        dotDashed = "DotDashed",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        dot2Dashed = "Dot2Dashed",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        triple = "Triple",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        thinThickSmall = "ThinThickSmall",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        thickThinSmall = "ThickThinSmall",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        thinThickThinSmall = "ThinThickThinSmall",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        thinThickMed = "ThinThickMed",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        thickThinMed = "ThickThinMed",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        thinThickThinMed = "ThinThickThinMed",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        thinThickLarge = "ThinThickLarge",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        thickThinLarge = "ThickThinLarge",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        thinThickThinLarge = "ThinThickThinLarge",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        wave = "Wave",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        doubleWave = "DoubleWave",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        dashedSmall = "DashedSmall",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        dashDotStroked = "DashDotStroked",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        threeDEmboss = "ThreeDEmboss",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        threeDEngrave = "ThreeDEngrave",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum VerticalAlignment {
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        mixed = "Mixed",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        top = "Top",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        center = "Center",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        bottom = "Bottom",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum ListLevelType {
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        bullet = "Bullet",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        number = "Number",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        picture = "Picture",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum ListBullet {
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        custom = "Custom",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        solid = "Solid",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        hollow = "Hollow",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        square = "Square",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        diamonds = "Diamonds",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        arrow = "Arrow",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        checkmark = "Checkmark",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum ListNumbering {
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        none = "None",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        arabic = "Arabic",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        upperRoman = "UpperRoman",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        lowerRoman = "LowerRoman",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        upperLetter = "UpperLetter",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        lowerLetter = "LowerLetter",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum Style {
        /**
         * Mixed styles or other style not in this list.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        other = "Other",
        /**
         * Reset character and paragraph style to default.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        normal = "Normal",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        heading1 = "Heading1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        heading2 = "Heading2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        heading3 = "Heading3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        heading4 = "Heading4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        heading5 = "Heading5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        heading6 = "Heading6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        heading7 = "Heading7",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        heading8 = "Heading8",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        heading9 = "Heading9",
        /**
         * Table-of-content level 1.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        toc1 = "Toc1",
        /**
         * Table-of-content level 2.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        toc2 = "Toc2",
        /**
         * Table-of-content level 3.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        toc3 = "Toc3",
        /**
         * Table-of-content level 4.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        toc4 = "Toc4",
        /**
         * Table-of-content level 5.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        toc5 = "Toc5",
        /**
         * Table-of-content level 6.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        toc6 = "Toc6",
        /**
         * Table-of-content level 7.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        toc7 = "Toc7",
        /**
         * Table-of-content level 8.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        toc8 = "Toc8",
        /**
         * Table-of-content level 9.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        toc9 = "Toc9",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        footnoteText = "FootnoteText",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        header = "Header",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        footer = "Footer",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        caption = "Caption",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        footnoteReference = "FootnoteReference",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        endnoteReference = "EndnoteReference",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        endnoteText = "EndnoteText",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        title = "Title",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        subtitle = "Subtitle",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        hyperlink = "Hyperlink",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        strong = "Strong",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        emphasis = "Emphasis",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        noSpacing = "NoSpacing",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listParagraph = "ListParagraph",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        quote = "Quote",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        intenseQuote = "IntenseQuote",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        subtleEmphasis = "SubtleEmphasis",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        intenseEmphasis = "IntenseEmphasis",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        subtleReference = "SubtleReference",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        intenseReference = "IntenseReference",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        bookTitle = "BookTitle",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        bibliography = "Bibliography",
        /**
         * Table-of-content heading.
         * @remarks
         * [Api set: WordApi 1.3]
         */
        tocHeading = "TocHeading",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        tableGrid = "TableGrid",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        plainTable1 = "PlainTable1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        plainTable2 = "PlainTable2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        plainTable3 = "PlainTable3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        plainTable4 = "PlainTable4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        plainTable5 = "PlainTable5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        tableGridLight = "TableGridLight",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable1Light = "GridTable1Light",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable1Light_Accent1 = "GridTable1Light_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable1Light_Accent2 = "GridTable1Light_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable1Light_Accent3 = "GridTable1Light_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable1Light_Accent4 = "GridTable1Light_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable1Light_Accent5 = "GridTable1Light_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable1Light_Accent6 = "GridTable1Light_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable2 = "GridTable2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable2_Accent1 = "GridTable2_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable2_Accent2 = "GridTable2_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable2_Accent3 = "GridTable2_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable2_Accent4 = "GridTable2_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable2_Accent5 = "GridTable2_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable2_Accent6 = "GridTable2_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable3 = "GridTable3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable3_Accent1 = "GridTable3_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable3_Accent2 = "GridTable3_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable3_Accent3 = "GridTable3_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable3_Accent4 = "GridTable3_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable3_Accent5 = "GridTable3_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable3_Accent6 = "GridTable3_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable4 = "GridTable4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable4_Accent1 = "GridTable4_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable4_Accent2 = "GridTable4_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable4_Accent3 = "GridTable4_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable4_Accent4 = "GridTable4_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable4_Accent5 = "GridTable4_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable4_Accent6 = "GridTable4_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable5Dark = "GridTable5Dark",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable5Dark_Accent1 = "GridTable5Dark_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable5Dark_Accent2 = "GridTable5Dark_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable5Dark_Accent3 = "GridTable5Dark_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable5Dark_Accent4 = "GridTable5Dark_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable5Dark_Accent5 = "GridTable5Dark_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable5Dark_Accent6 = "GridTable5Dark_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable6Colorful = "GridTable6Colorful",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable6Colorful_Accent1 = "GridTable6Colorful_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable6Colorful_Accent2 = "GridTable6Colorful_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable6Colorful_Accent3 = "GridTable6Colorful_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable6Colorful_Accent4 = "GridTable6Colorful_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable6Colorful_Accent5 = "GridTable6Colorful_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable6Colorful_Accent6 = "GridTable6Colorful_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable7Colorful = "GridTable7Colorful",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable7Colorful_Accent1 = "GridTable7Colorful_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable7Colorful_Accent2 = "GridTable7Colorful_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable7Colorful_Accent3 = "GridTable7Colorful_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable7Colorful_Accent4 = "GridTable7Colorful_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable7Colorful_Accent5 = "GridTable7Colorful_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        gridTable7Colorful_Accent6 = "GridTable7Colorful_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable1Light = "ListTable1Light",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable1Light_Accent1 = "ListTable1Light_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable1Light_Accent2 = "ListTable1Light_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable1Light_Accent3 = "ListTable1Light_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable1Light_Accent4 = "ListTable1Light_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable1Light_Accent5 = "ListTable1Light_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable1Light_Accent6 = "ListTable1Light_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable2 = "ListTable2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable2_Accent1 = "ListTable2_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable2_Accent2 = "ListTable2_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable2_Accent3 = "ListTable2_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable2_Accent4 = "ListTable2_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable2_Accent5 = "ListTable2_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable2_Accent6 = "ListTable2_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable3 = "ListTable3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable3_Accent1 = "ListTable3_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable3_Accent2 = "ListTable3_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable3_Accent3 = "ListTable3_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable3_Accent4 = "ListTable3_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable3_Accent5 = "ListTable3_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable3_Accent6 = "ListTable3_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable4 = "ListTable4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable4_Accent1 = "ListTable4_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable4_Accent2 = "ListTable4_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable4_Accent3 = "ListTable4_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable4_Accent4 = "ListTable4_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable4_Accent5 = "ListTable4_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable4_Accent6 = "ListTable4_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable5Dark = "ListTable5Dark",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable5Dark_Accent1 = "ListTable5Dark_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable5Dark_Accent2 = "ListTable5Dark_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable5Dark_Accent3 = "ListTable5Dark_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable5Dark_Accent4 = "ListTable5Dark_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable5Dark_Accent5 = "ListTable5Dark_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable5Dark_Accent6 = "ListTable5Dark_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable6Colorful = "ListTable6Colorful",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable6Colorful_Accent1 = "ListTable6Colorful_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable6Colorful_Accent2 = "ListTable6Colorful_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable6Colorful_Accent3 = "ListTable6Colorful_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable6Colorful_Accent4 = "ListTable6Colorful_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable6Colorful_Accent5 = "ListTable6Colorful_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable6Colorful_Accent6 = "ListTable6Colorful_Accent6",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable7Colorful = "ListTable7Colorful",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable7Colorful_Accent1 = "ListTable7Colorful_Accent1",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable7Colorful_Accent2 = "ListTable7Colorful_Accent2",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable7Colorful_Accent3 = "ListTable7Colorful_Accent3",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable7Colorful_Accent4 = "ListTable7Colorful_Accent4",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable7Colorful_Accent5 = "ListTable7Colorful_Accent5",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        listTable7Colorful_Accent6 = "ListTable7Colorful_Accent6",
    }
    /**
     * @remarks
     * [Api set: WordApi 1.3]
     */
    enum DocumentPropertyType {
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        string = "String",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        number = "Number",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        date = "Date",
        /**
         * @remarks
         * [Api set: WordApi 1.3]
         */
        boolean = "Boolean",
    }
    enum ErrorCodes {
        accessDenied = "AccessDenied",
        generalException = "GeneralException",
        invalidArgument = "InvalidArgument",
        itemNotFound = "ItemNotFound",
        notImplemented = "NotImplemented",
        searchDialogIsOpen = "SearchDialogIsOpen",
        searchStringInvalidOrTooLong = "SearchStringInvalidOrTooLong",
    }
    export module Interfaces {
        /**
        * Provides ways to load properties of only a subset of members of a collection.
        */
        export interface CollectionLoadOptions {
            /**
            * Specify the number of items in the queried collection to be included in the result.
            */
            $top?: number;
            /**
            * Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.
            */
            $skip?: number;
        }
        /** An interface for updating data on the Body object, for use in `body.set({ ... })`. */
        export interface BodyUpdateData {
            /**
            * Gets the text format of the body. Use this to get and set font name, size, color and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             * Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             * Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        }
        /** An interface for updating data on the Comment object, for use in `comment.set({ ... })`. */
        export interface CommentUpdateData {
            /**
            * Gets or sets the comment's content range.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            contentRange?: Word.Interfaces.CommentContentRangeUpdateData;
            /**
             * Gets or sets the comment's content as plain text.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            content?: string;
            /**
             * Gets or sets the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            resolved?: boolean;
        }
        /** An interface for updating data on the CommentCollection object, for use in `commentCollection.set({ ... })`. */
        export interface CommentCollectionUpdateData {
            items?: Word.Interfaces.CommentData[];
        }
        /** An interface for updating data on the CommentContentRange object, for use in `commentContentRange.set({ ... })`. */
        export interface CommentContentRangeUpdateData {
            /**
             * Gets or sets a value that indicates whether the comment text is bold.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            bold?: boolean;
            /**
             * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            hyperlink?: string;
            /**
             * Gets or sets a value that indicates whether the comment text is italicized.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            italic?: boolean;
            /**
             * Gets or sets a value that indicates whether the comment text has a strikethrough.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            strikeThrough?: boolean;
            /**
             * Gets or sets a value that indicates the comment text's underline type. 'None' if the comment text is not underlined.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
        }
        /** An interface for updating data on the CommentReply object, for use in `commentReply.set({ ... })`. */
        export interface CommentReplyUpdateData {
            /**
            * Gets or sets the commentReply's content range.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            contentRange?: Word.Interfaces.CommentContentRangeUpdateData;
            /**
            * Gets the parent comment of this reply.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            parentComment?: Word.Interfaces.CommentUpdateData;
            /**
             * Gets or sets the comment reply's content. The string is plain text.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            content?: string;
        }
        /** An interface for updating data on the CommentReplyCollection object, for use in `commentReplyCollection.set({ ... })`. */
        export interface CommentReplyCollectionUpdateData {
            items?: Word.Interfaces.CommentReplyData[];
        }
        /** An interface for updating data on the ContentControl object, for use in `contentControl.set({ ... })`. */
        export interface ContentControlUpdateData {
            /**
            * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             * Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
            /**
             * Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotDelete?: boolean;
            /**
             * Gets or sets a value that indicates whether the user can edit the contents of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotEdit?: boolean;
            /**
             * Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: string;
            /**
             * Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             * **Note**: The set operation for this property is not supported in Word on the web.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            placeholderText?: string;
            /**
             * Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            removeWhenEdited?: boolean;
            /**
             * Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             * Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             * Gets or sets a tag to identify a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            tag?: string;
            /**
             * Gets or sets the title for a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            title?: string;
        }
        /** An interface for updating data on the ContentControlCollection object, for use in `contentControlCollection.set({ ... })`. */
        export interface ContentControlCollectionUpdateData {
            items?: Word.Interfaces.ContentControlData[];
        }
        /** An interface for updating data on the CustomProperty object, for use in `customProperty.set({ ... })`. */
        export interface CustomPropertyUpdateData {
            /**
             * Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            value?: any;
        }
        /** An interface for updating data on the CustomPropertyCollection object, for use in `customPropertyCollection.set({ ... })`. */
        export interface CustomPropertyCollectionUpdateData {
            items?: Word.Interfaces.CustomPropertyData[];
        }
        /** An interface for updating data on the CustomXmlPartCollection object, for use in `customXmlPartCollection.set({ ... })`. */
        export interface CustomXmlPartCollectionUpdateData {
            items?: Word.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the CustomXmlPartScopedCollection object, for use in `customXmlPartScopedCollection.set({ ... })`. */
        export interface CustomXmlPartScopedCollectionUpdateData {
            items?: Word.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the Document object, for use in `document.set({ ... })`. */
        export interface DocumentUpdateData {
            /**
            * Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyUpdateData;
            /**
            * Gets the properties of the document.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            properties?: Word.Interfaces.DocumentPropertiesUpdateData;
            /**
             * Gets or sets the ChangeTracking mode.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            changeTrackingMode?: Word.ChangeTrackingMode | "Off" | "TrackAll" | "TrackMineOnly";
        }
        /** An interface for updating data on the DocumentCreated object, for use in `documentCreated.set({ ... })`. */
        export interface DocumentCreatedUpdateData {
            /**
            * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
            *
            * @remarks
            * [Api set: WordApiHiddenDocument 1.3]
            */
            body?: Word.Interfaces.BodyUpdateData;
            /**
            * Gets the properties of the document.
            *
            * @remarks
            * [Api set: WordApiHiddenDocument 1.3]
            */
            properties?: Word.Interfaces.DocumentPropertiesUpdateData;
        }
        /** An interface for updating data on the DocumentProperties object, for use in `documentProperties.set({ ... })`. */
        export interface DocumentPropertiesUpdateData {
            /**
             * Gets or sets the author of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            author?: string;
            /**
             * Gets or sets the category of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            category?: string;
            /**
             * Gets or sets the comments of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            comments?: string;
            /**
             * Gets or sets the company of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            company?: string;
            /**
             * Gets or sets the format of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            format?: string;
            /**
             * Gets or sets the keywords of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            keywords?: string;
            /**
             * Gets or sets the manager of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            manager?: string;
            /**
             * Gets or sets the subject of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            subject?: string;
            /**
             * Gets or sets the title of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            title?: string;
        }
        /** An interface for updating data on the FieldCollection object, for use in `fieldCollection.set({ ... })`. */
        export interface FieldCollectionUpdateData {
            items?: Word.Interfaces.FieldData[];
        }
        /** An interface for updating data on the Font object, for use in `font.set({ ... })`. */
        export interface FontUpdateData {
            /**
             * Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            bold?: boolean;
            /**
             * Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: string;
            /**
             * Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            doubleStrikeThrough?: boolean;
            /**
             * Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
             * **Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            highlightColor?: string;
            /**
             * Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            italic?: boolean;
            /**
             * Gets or sets a value that represents the name of the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            name?: string;
            /**
             * Gets or sets a value that represents the font size in points.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            size?: number;
            /**
             * Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            strikeThrough?: boolean;
            /**
             * Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            subscript?: boolean;
            /**
             * Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            superscript?: boolean;
            /**
             * Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
        }
        /** An interface for updating data on the InlinePicture object, for use in `inlinePicture.set({ ... })`. */
        export interface InlinePictureUpdateData {
            /**
             * Gets or sets a string that represents the alternative text associated with the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: string;
            /**
             * Gets or sets a string that contains the title for the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: string;
            /**
             * Gets or sets a number that describes the height of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            height?: number;
            /**
             * Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            hyperlink?: string;
            /**
             * Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lockAspectRatio?: boolean;
            /**
             * Gets or sets a number that describes the width of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            width?: number;
        }
        /** An interface for updating data on the InlinePictureCollection object, for use in `inlinePictureCollection.set({ ... })`. */
        export interface InlinePictureCollectionUpdateData {
            items?: Word.Interfaces.InlinePictureData[];
        }
        /** An interface for updating data on the ListCollection object, for use in `listCollection.set({ ... })`. */
        export interface ListCollectionUpdateData {
            items?: Word.Interfaces.ListData[];
        }
        /** An interface for updating data on the ListItem object, for use in `listItem.set({ ... })`. */
        export interface ListItemUpdateData {
            /**
             * Gets or sets the level of the item in the list.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            level?: number;
        }
        /** An interface for updating data on the NoteItem object, for use in `noteItem.set({ ... })`. */
        export interface NoteItemUpdateData {
            /**
            * Represents the body object of the note item. It's the portion of the text within the footnote or endnote.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            body?: Word.Interfaces.BodyUpdateData;
            /**
            * Represents a footnote or endnote reference in the main document.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            reference?: Word.Interfaces.RangeUpdateData;
        }
        /** An interface for updating data on the NoteItemCollection object, for use in `noteItemCollection.set({ ... })`. */
        export interface NoteItemCollectionUpdateData {
            items?: Word.Interfaces.NoteItemData[];
        }
        /** An interface for updating data on the Paragraph object, for use in `paragraph.set({ ... })`. */
        export interface ParagraphUpdateData {
            /**
            * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
            * Gets the ListItem for the paragraph. Throws an error if the paragraph is not part of a list.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            listItem?: Word.Interfaces.ListItemUpdateData;
            /**
            * Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            listItemOrNullObject?: Word.Interfaces.ListItemUpdateData;
            /**
             * Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: number;
            /**
             * Gets or sets the left indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            leftIndent?: number;
            /**
             * Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: number;
            /**
             * Gets or sets the amount of spacing, in grid lines, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: number;
            /**
             * Gets or sets the amount of spacing, in grid lines, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: number;
            /**
             * Gets or sets the outline level for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: number;
            /**
             * Gets or sets the right indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            rightIndent?: number;
            /**
             * Gets or sets the spacing, in points, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: number;
            /**
             * Gets or sets the spacing, in points, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: number;
            /**
             * Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             * Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        }
        /** An interface for updating data on the ParagraphCollection object, for use in `paragraphCollection.set({ ... })`. */
        export interface ParagraphCollectionUpdateData {
            items?: Word.Interfaces.ParagraphData[];
        }
        /** An interface for updating data on the Range object, for use in `range.set({ ... })`. */
        export interface RangeUpdateData {
            /**
            * Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            hyperlink?: string;
            /**
             * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             * Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
        }
        /** An interface for updating data on the RangeCollection object, for use in `rangeCollection.set({ ... })`. */
        export interface RangeCollectionUpdateData {
            items?: Word.Interfaces.RangeData[];
        }
        /** An interface for updating data on the SearchOptions object, for use in `searchOptions.set({ ... })`. */
        export interface SearchOptionsUpdateData {
            /**
             * Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignorePunct?: boolean;
            /**
             * Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignoreSpace?: boolean;
            /**
             * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchCase?: boolean;
            /**
             * Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchPrefix?: boolean;
            /**
             * Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchSuffix?: boolean;
            /**
             * Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWholeWord?: boolean;
            /**
             * Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWildcards?: boolean;
        }
        /** An interface for updating data on the Section object, for use in `section.set({ ... })`. */
        export interface SectionUpdateData {
            /**
            * Gets the body object of the section. This does not include the header/footer and other section metadata.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyUpdateData;
        }
        /** An interface for updating data on the SectionCollection object, for use in `sectionCollection.set({ ... })`. */
        export interface SectionCollectionUpdateData {
            items?: Word.Interfaces.SectionData[];
        }
        /** An interface for updating data on the Setting object, for use in `setting.set({ ... })`. */
        export interface SettingUpdateData {
            /**
             * Gets or sets the value of the setting.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            value?: any;
        }
        /** An interface for updating data on the SettingCollection object, for use in `settingCollection.set({ ... })`. */
        export interface SettingCollectionUpdateData {
            items?: Word.Interfaces.SettingData[];
        }
        /** An interface for updating data on the Table object, for use in `table.set({ ... })`. */
        export interface TableUpdateData {
            /**
            * Gets the font. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             * Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Gets and sets the number of header rows.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            headerRowCount?: number;
            /**
             * Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: string;
            /**
             * Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            style?: string;
            /**
             * Gets and sets whether the table has banded columns.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBandedColumns?: boolean;
            /**
             * Gets and sets whether the table has banded rows.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBandedRows?: boolean;
            /**
             * Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             * Gets and sets whether the table has a first column with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleFirstColumn?: boolean;
            /**
             * Gets and sets whether the table has a last column with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleLastColumn?: boolean;
            /**
             * Gets and sets whether the table has a total (last) row with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleTotalRow?: boolean;
            /**
             * Gets and sets the text values in the table, as a 2D Javascript array.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            values?: string[][];
            /**
             * Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
            /**
             * Gets and sets the width of the table in points.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /** An interface for updating data on the TableCollection object, for use in `tableCollection.set({ ... })`. */
        export interface TableCollectionUpdateData {
            items?: Word.Interfaces.TableData[];
        }
        /** An interface for updating data on the TableRow object, for use in `tableRow.set({ ... })`. */
        export interface TableRowUpdateData {
            /**
            * Gets the font. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             * Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Gets and sets the preferred height of the row in points.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            preferredHeight?: number;
            /**
             * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: string;
            /**
             * Gets and sets the text values in the row, as a 2D Javascript array.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            values?: string[][];
            /**
             * Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        }
        /** An interface for updating data on the TableRowCollection object, for use in `tableRowCollection.set({ ... })`. */
        export interface TableRowCollectionUpdateData {
            items?: Word.Interfaces.TableRowData[];
        }
        /** An interface for updating data on the TableCell object, for use in `tableCell.set({ ... })`. */
        export interface TableCellUpdateData {
            /**
            * Gets the body object of the cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            body?: Word.Interfaces.BodyUpdateData;
            /**
             * Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            columnWidth?: number;
            /**
             * Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: string;
            /**
             * Gets and sets the text of the cell.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            value?: string;
            /**
             * Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        }
        /** An interface for updating data on the TableCellCollection object, for use in `tableCellCollection.set({ ... })`. */
        export interface TableCellCollectionUpdateData {
            items?: Word.Interfaces.TableCellData[];
        }
        /** An interface for updating data on the TableBorder object, for use in `tableBorder.set({ ... })`. */
        export interface TableBorderUpdateData {
            /**
             * Gets or sets the table border color.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            color?: string;
            /**
             * Gets or sets the type of the table border.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            type?: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
            /**
             * Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling `body.toJSON()`. */
        export interface BodyData {
            /**
            * Gets the collection of rich text content control objects in the body. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            /**
            * Gets the collection of field objects in the body. Read-only.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            fields?: Word.Interfaces.FieldData[];
            /**
            * Gets the text format of the body. Use this to get and set font name, size, color and other properties. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            /**
            * Gets the collection of InlinePicture objects in the body. The collection does not include floating images. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            /**
            * Gets the collection of list objects in the body. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            lists?: Word.Interfaces.ListData[];
            /**
            * Gets the collection of paragraph objects in the body. Read-only. **Important**: Paragraphs in tables are not returned for requirement sets 1.1 and 1.2. From requirement set 1.3, paragraphs in tables are also returned.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            /**
            * Gets the collection of table objects in the body. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            tables?: Word.Interfaces.TableData[];
            /**
             * Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             * Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             * Gets the text of the body. Use the insertText method to insert text. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: string;
            /**
             * Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            type?: Word.BodyType | "Unknown" | "MainDoc" | "Section" | "Header" | "Footer" | "TableCell" | "Footnote" | "Endnote" | "NoteItem";
        }
        /** An interface describing the data returned by calling `comment.toJSON()`. */
        export interface CommentData {
            /**
            * Gets or sets the comment's content range.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            contentRange?: Word.Interfaces.CommentContentRangeData;
            /**
            * Gets the collection of reply objects associated with the comment.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            replies?: Word.Interfaces.CommentReplyData[];
            /**
             * Gets the email of the comment's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorEmail?: string;
            /**
             * Gets the name of the comment's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorName?: string;
            /**
             * Gets or sets the comment's content as plain text.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            content?: string;
            /**
             * Gets the creation date of the comment.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            creationDate?: Date;
            /**
             * Gets the ID of the comment. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            id?: string;
            /**
             * Gets or sets the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            resolved?: boolean;
        }
        /** An interface describing the data returned by calling `commentCollection.toJSON()`. */
        export interface CommentCollectionData {
            items?: Word.Interfaces.CommentData[];
        }
        /** An interface describing the data returned by calling `commentContentRange.toJSON()`. */
        export interface CommentContentRangeData {
            /**
             * Gets or sets a value that indicates whether the comment text is bold.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            bold?: boolean;
            /**
             * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            hyperlink?: string;
            /**
             * Checks whether the range length is zero. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            isEmpty?: boolean;
            /**
             * Gets or sets a value that indicates whether the comment text is italicized.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            italic?: boolean;
            /**
             * Gets or sets a value that indicates whether the comment text has a strikethrough.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            strikeThrough?: boolean;
            /**
             * Gets the text of the comment range. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            text?: string;
            /**
             * Gets or sets a value that indicates the comment text's underline type. 'None' if the comment text is not underlined.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
        }
        /** An interface describing the data returned by calling `commentReply.toJSON()`. */
        export interface CommentReplyData {
            /**
            * Gets or sets the commentReply's content range.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            contentRange?: Word.Interfaces.CommentContentRangeData;
            /**
            * Gets the parent comment of this reply.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            parentComment?: Word.Interfaces.CommentData;
            /**
             * Gets the email of the comment reply's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorEmail?: string;
            /**
             * Gets the name of the comment reply's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorName?: string;
            /**
             * Gets or sets the comment reply's content. The string is plain text.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            content?: string;
            /**
             * Gets the creation date of the comment reply.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            creationDate?: Date;
            /**
             * Gets the ID of the comment reply. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            id?: string;
        }
        /** An interface describing the data returned by calling `commentReplyCollection.toJSON()`. */
        export interface CommentReplyCollectionData {
            items?: Word.Interfaces.CommentReplyData[];
        }
        /** An interface describing the data returned by calling `contentControl.toJSON()`. */
        export interface ContentControlData {
            /**
            * Gets the collection of content control objects in the content control. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            /**
            * Gets the collection of field objects in the contentcontrol. Read-only.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            fields?: Word.Interfaces.FieldData[];
            /**
            * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            /**
            * Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            /**
            * Gets the collection of list objects in the content control. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            lists?: Word.Interfaces.ListData[];
            /**
            * Get the collection of paragraph objects in the content control. Read-only. **Important**: For requirement sets 1.1 and 1.2, paragraphs in tables wholly contained within this content control are not returned. From requirement set 1.3, paragraphs in such tables are also returned.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            /**
            * Gets the collection of table objects in the content control. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            tables?: Word.Interfaces.TableData[];
            /**
             * Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
            /**
             * Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotDelete?: boolean;
            /**
             * Gets or sets a value that indicates whether the user can edit the contents of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotEdit?: boolean;
            /**
             * Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: string;
            /**
             * Gets an integer that represents the content control identifier. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            id?: number;
            /**
             * Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             * **Note**: The set operation for this property is not supported in Word on the web.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            placeholderText?: string;
            /**
             * Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            removeWhenEdited?: boolean;
            /**
             * Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             * Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             * Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            subtype?: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText";
            /**
             * Gets or sets a tag to identify a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            tag?: string;
            /**
             * Gets the text of the content control. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: string;
            /**
             * Gets or sets the title for a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            title?: string;
            /**
             * Gets the content control type. Only rich text content controls are supported currently. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            type?: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText";
        }
        /** An interface describing the data returned by calling `contentControlCollection.toJSON()`. */
        export interface ContentControlCollectionData {
            items?: Word.Interfaces.ContentControlData[];
        }
        /** An interface describing the data returned by calling `customProperty.toJSON()`. */
        export interface CustomPropertyData {
            /**
             * Gets the key of the custom property. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            key?: string;
            /**
             * Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            type?: Word.DocumentPropertyType | "String" | "Number" | "Date" | "Boolean";
            /**
             * Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            value?: any;
        }
        /** An interface describing the data returned by calling `customPropertyCollection.toJSON()`. */
        export interface CustomPropertyCollectionData {
            items?: Word.Interfaces.CustomPropertyData[];
        }
        /** An interface describing the data returned by calling `customXmlPart.toJSON()`. */
        export interface CustomXmlPartData {
            /**
             * Gets the ID of the custom XML part. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: string;
            /**
             * Gets the namespace URI of the custom XML part. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            namespaceUri?: string;
        }
        /** An interface describing the data returned by calling `customXmlPartCollection.toJSON()`. */
        export interface CustomXmlPartCollectionData {
            items?: Word.Interfaces.CustomXmlPartData[];
        }
        /** An interface describing the data returned by calling `customXmlPartScopedCollection.toJSON()`. */
        export interface CustomXmlPartScopedCollectionData {
            items?: Word.Interfaces.CustomXmlPartData[];
        }
        /** An interface describing the data returned by calling `document.toJSON()`. */
        export interface DocumentData {
            /**
            * Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyData;
            /**
            * Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            /**
            * Gets the custom XML parts in the document. Read-only.
            *
            * @remarks
            * [Api set: WordApi BETA (PREVIEW ONLY)]
            * @beta
            */
            customXmlParts?: Word.Interfaces.CustomXmlPartData[];
            /**
            * Gets the properties of the document. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            properties?: Word.Interfaces.DocumentPropertiesData;
            /**
            * Gets the collection of section objects in the document. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            sections?: Word.Interfaces.SectionData[];
            /**
            * Gets the add-in's settings in the document. Read-only.
            *
            * @remarks
            * [Api set: WordApi BETA (PREVIEW ONLY)]
            * @beta
            */
            settings?: Word.Interfaces.SettingData[];
            /**
             * Gets or sets the ChangeTracking mode.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            changeTrackingMode?: Word.ChangeTrackingMode | "Off" | "TrackAll" | "TrackMineOnly";
            /**
             * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            saved?: boolean;
        }
        /** An interface describing the data returned by calling `documentCreated.toJSON()`. */
        export interface DocumentCreatedData {
            /**
            * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc. Read-only.
            *
            * @remarks
            * [Api set: WordApiHiddenDocument 1.3]
            */
            body?: Word.Interfaces.BodyData;
            /**
            * Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc. Read-only.
            *
            * @remarks
            * [Api set: WordApiHiddenDocument 1.3]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            /**
            * Gets the custom XML parts in the document. Read-only.
            *
            * @remarks
            * [Api set: WordApiHiddenDocument 1.4]
            * @beta
            */
            customXmlParts?: Word.Interfaces.CustomXmlPartData[];
            /**
            * Gets the properties of the document. Read-only.
            *
            * @remarks
            * [Api set: WordApiHiddenDocument 1.3]
            */
            properties?: Word.Interfaces.DocumentPropertiesData;
            /**
            * Gets the collection of section objects in the document. Read-only.
            *
            * @remarks
            * [Api set: WordApiHiddenDocument 1.3]
            */
            sections?: Word.Interfaces.SectionData[];
            /**
            * Gets the add-in's settings in the document. Read-only.
            *
            * @remarks
            * [Api set: WordApiHiddenDocument 1.4]
            * @beta
            */
            settings?: Word.Interfaces.SettingData[];
            /**
             * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
             *
             * @remarks
             * [Api set: WordApiHiddenDocument 1.3]
             */
            saved?: boolean;
        }
        /** An interface describing the data returned by calling `documentProperties.toJSON()`. */
        export interface DocumentPropertiesData {
            /**
            * Gets the collection of custom properties of the document. Read only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            customProperties?: Word.Interfaces.CustomPropertyData[];
            /**
             * Gets the application name of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            applicationName?: string;
            /**
             * Gets or sets the author of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            author?: string;
            /**
             * Gets or sets the category of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            category?: string;
            /**
             * Gets or sets the comments of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            comments?: string;
            /**
             * Gets or sets the company of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            company?: string;
            /**
             * Gets the creation date of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            creationDate?: Date;
            /**
             * Gets or sets the format of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            format?: string;
            /**
             * Gets or sets the keywords of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            keywords?: string;
            /**
             * Gets the last author of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            lastAuthor?: string;
            /**
             * Gets the last print date of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            lastPrintDate?: Date;
            /**
             * Gets the last save time of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            lastSaveTime?: Date;
            /**
             * Gets or sets the manager of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            manager?: string;
            /**
             * Gets the revision number of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            revisionNumber?: string;
            /**
             * Gets security settings of the document. Read only. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            security?: number;
            /**
             * Gets or sets the subject of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            subject?: string;
            /**
             * Gets the template of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            template?: string;
            /**
             * Gets or sets the title of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            title?: string;
        }
        /** An interface describing the data returned by calling `field.toJSON()`. */
        export interface FieldData {
            /**
             * Gets the field's code instruction. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
             * @beta
             */
            code?: string;
            /**
             * Gets the field's result data. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
             * @beta
             */
            result?: string;
        }
        /** An interface describing the data returned by calling `fieldCollection.toJSON()`. */
        export interface FieldCollectionData {
            items?: Word.Interfaces.FieldData[];
        }
        /** An interface describing the data returned by calling `font.toJSON()`. */
        export interface FontData {
            /**
             * Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            bold?: boolean;
            /**
             * Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: string;
            /**
             * Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            doubleStrikeThrough?: boolean;
            /**
             * Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
             * **Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            highlightColor?: string;
            /**
             * Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            italic?: boolean;
            /**
             * Gets or sets a value that represents the name of the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            name?: string;
            /**
             * Gets or sets a value that represents the font size in points.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            size?: number;
            /**
             * Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            strikeThrough?: boolean;
            /**
             * Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            subscript?: boolean;
            /**
             * Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            superscript?: boolean;
            /**
             * Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
        }
        /** An interface describing the data returned by calling `inlinePicture.toJSON()`. */
        export interface InlinePictureData {
            /**
             * Gets or sets a string that represents the alternative text associated with the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: string;
            /**
             * Gets or sets a string that contains the title for the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: string;
            /**
             * Gets or sets a number that describes the height of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            height?: number;
            /**
             * Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            hyperlink?: string;
            /**
             * Gets the format of the inline image. Read-only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            imageFormat?: Word.ImageFormat | "Unsupported" | "Undefined" | "Bmp" | "Jpeg" | "Gif" | "Tiff" | "Png" | "Icon" | "Exif" | "Wmf" | "Emf" | "Pict" | "Pdf" | "Svg";
            /**
             * Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lockAspectRatio?: boolean;
            /**
             * Gets or sets a number that describes the width of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling `inlinePictureCollection.toJSON()`. */
        export interface InlinePictureCollectionData {
            items?: Word.Interfaces.InlinePictureData[];
        }
        /** An interface describing the data returned by calling `list.toJSON()`. */
        export interface ListData {
            /**
            * Gets paragraphs in the list. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            /**
             * Gets the list's id.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            id?: number;
            /**
             * Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            levelExistences?: boolean[];
            /**
             * Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            levelTypes?: Word.ListLevelType[];
        }
        /** An interface describing the data returned by calling `listCollection.toJSON()`. */
        export interface ListCollectionData {
            items?: Word.Interfaces.ListData[];
        }
        /** An interface describing the data returned by calling `listItem.toJSON()`. */
        export interface ListItemData {
            /**
             * Gets or sets the level of the item in the list.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            level?: number;
            /**
             * Gets the list item bullet, number, or picture as a string. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            listString?: string;
            /**
             * Gets the list item order number in relation to its siblings. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            siblingIndex?: number;
        }
        /** An interface describing the data returned by calling `noteItem.toJSON()`. */
        export interface NoteItemData {
            /**
            * Represents the body object of the note item. It's the portion of the text within the footnote or endnote.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            body?: Word.Interfaces.BodyData;
            /**
            * Represents a footnote or endnote reference in the main document.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            reference?: Word.Interfaces.RangeData;
            /**
             * Represents the note item type: footnote or endnote.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            type?: Word.NoteItemType | "Footnote" | "Endnote";
        }
        /** An interface describing the data returned by calling `noteItemCollection.toJSON()`. */
        export interface NoteItemCollectionData {
            items?: Word.Interfaces.NoteItemData[];
        }
        /** An interface describing the data returned by calling `paragraph.toJSON()`. */
        export interface ParagraphData {
            /**
            * Gets the collection of fields in the paragraph. Read-only.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            fields?: Word.Interfaces.FieldData[];
            /**
            * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            /**
            * Gets the collection of InlinePicture objects in the paragraph. The collection does not include floating images. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            /**
            * Gets the ListItem for the paragraph. Throws an error if the paragraph is not part of a list. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            listItem?: Word.Interfaces.ListItemData;
            /**
            * Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            listItemOrNullObject?: Word.Interfaces.ListItemData;
            /**
             * Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: number;
            /**
             * Indicates the paragraph is the last one inside its parent body. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isLastParagraph?: boolean;
            /**
             * Checks whether the paragraph is a list item. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isListItem?: boolean;
            /**
             * Gets or sets the left indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            leftIndent?: number;
            /**
             * Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: number;
            /**
             * Gets or sets the amount of spacing, in grid lines, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: number;
            /**
             * Gets or sets the amount of spacing, in grid lines, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: number;
            /**
             * Gets or sets the outline level for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: number;
            /**
             * Gets or sets the right indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            rightIndent?: number;
            /**
             * Gets or sets the spacing, in points, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: number;
            /**
             * Gets or sets the spacing, in points, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: number;
            /**
             * Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             * Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             * Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            tableNestingLevel?: number;
            /**
             * Gets the text of the paragraph. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling `paragraphCollection.toJSON()`. */
        export interface ParagraphCollectionData {
            items?: Word.Interfaces.ParagraphData[];
        }
        /** An interface describing the data returned by calling `range.toJSON()`. */
        export interface RangeData {
            /**
            * Gets the collection of field objects in the range. Read-only.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            fields?: Word.Interfaces.FieldData[];
            /**
            * Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            /**
            * Gets the collection of inline picture objects in the range. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.2]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            /**
             * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            hyperlink?: string;
            /**
             * Checks whether the range length is zero. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isEmpty?: boolean;
            /**
             * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            /**
             * Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             * Gets the text of the range. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling `rangeCollection.toJSON()`. */
        export interface RangeCollectionData {
            items?: Word.Interfaces.RangeData[];
        }
        /** An interface describing the data returned by calling `searchOptions.toJSON()`. */
        export interface SearchOptionsData {
            /**
             * Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignorePunct?: boolean;
            /**
             * Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignoreSpace?: boolean;
            /**
             * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchCase?: boolean;
            /**
             * Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchPrefix?: boolean;
            /**
             * Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchSuffix?: boolean;
            /**
             * Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWholeWord?: boolean;
            /**
             * Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWildcards?: boolean;
        }
        /** An interface describing the data returned by calling `section.toJSON()`. */
        export interface SectionData {
            /**
            * Gets the body object of the section. This does not include the header/footer and other section metadata. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyData;
        }
        /** An interface describing the data returned by calling `sectionCollection.toJSON()`. */
        export interface SectionCollectionData {
            items?: Word.Interfaces.SectionData[];
        }
        /** An interface describing the data returned by calling `setting.toJSON()`. */
        export interface SettingData {
            /**
             * Gets the key of the setting. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            key?: string;
            /**
             * Gets or sets the value of the setting.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            value?: any;
        }
        /** An interface describing the data returned by calling `settingCollection.toJSON()`. */
        export interface SettingCollectionData {
            items?: Word.Interfaces.SettingData[];
        }
        /** An interface describing the data returned by calling `table.toJSON()`. */
        export interface TableData {
            /**
            * Gets the collection of field objects in the table. Read-only.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            fields?: Word.Interfaces.FieldData[];
            /**
            * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontData;
            /**
            * Gets all of the table rows. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            rows?: Word.Interfaces.TableRowData[];
            /**
            * Gets the child tables nested one level deeper. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            tables?: Word.Interfaces.TableData[];
            /**
             * Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Gets and sets the number of header rows.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            headerRowCount?: number;
            /**
             * Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Indicates whether all of the table rows are uniform. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isUniform?: boolean;
            /**
             * Gets the nesting level of the table. Top-level tables have level 1. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            nestingLevel?: number;
            /**
             * Gets the number of rows in the table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            rowCount?: number;
            /**
             * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: string;
            /**
             * Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            style?: string;
            /**
             * Gets and sets whether the table has banded columns.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBandedColumns?: boolean;
            /**
             * Gets and sets whether the table has banded rows.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBandedRows?: boolean;
            /**
             * Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: Word.Style | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
            /**
             * Gets and sets whether the table has a first column with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleFirstColumn?: boolean;
            /**
             * Gets and sets whether the table has a last column with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleLastColumn?: boolean;
            /**
             * Gets and sets whether the table has a total (last) row with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleTotalRow?: boolean;
            /**
             * Gets and sets the text values in the table, as a 2D Javascript array.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            values?: string[][];
            /**
             * Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
            /**
             * Gets and sets the width of the table in points.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling `tableCollection.toJSON()`. */
        export interface TableCollectionData {
            items?: Word.Interfaces.TableData[];
        }
        /** An interface describing the data returned by calling `tableRow.toJSON()`. */
        export interface TableRowData {
            /**
            * Gets cells. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            cells?: Word.Interfaces.TableCellData[];
            /**
            * Gets the collection of field objects in the table row. Read-only.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            fields?: Word.Interfaces.FieldData[];
            /**
            * Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontData;
            /**
             * Gets the number of cells in the row. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            cellCount?: number;
            /**
             * Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isHeader?: boolean;
            /**
             * Gets and sets the preferred height of the row in points.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            preferredHeight?: number;
            /**
             * Gets the index of the row in its parent table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            rowIndex?: number;
            /**
             * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: string;
            /**
             * Gets and sets the text values in the row, as a 2D Javascript array.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            values?: string[][];
            /**
             * Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
        }
        /** An interface describing the data returned by calling `tableRowCollection.toJSON()`. */
        export interface TableRowCollectionData {
            items?: Word.Interfaces.TableRowData[];
        }
        /** An interface describing the data returned by calling `tableCell.toJSON()`. */
        export interface TableCellData {
            /**
            * Gets the body object of the cell. Read-only.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            body?: Word.Interfaces.BodyData;
            /**
             * Gets the index of the cell in its row. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            cellIndex?: number;
            /**
             * Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            columnWidth?: number;
            /**
             * Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Gets the index of the cell's row in the table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            rowIndex?: number;
            /**
             * Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: string;
            /**
             * Gets and sets the text of the cell.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            value?: string;
            /**
             * Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
            /**
             * Gets the width of the cell in points. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling `tableCellCollection.toJSON()`. */
        export interface TableCellCollectionData {
            items?: Word.Interfaces.TableCellData[];
        }
        /** An interface describing the data returned by calling `tableBorder.toJSON()`. */
        export interface TableBorderData {
            /**
             * Gets or sets the table border color.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            color?: string;
            /**
             * Gets or sets the type of the table border.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            type?: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
            /**
             * Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            width?: number;
        }
        /**
         * Represents the body of a document or a section.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface BodyLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the text format of the body. Use this to get and set font name, size, color and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws an error if there isn't a parent body.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * Gets the parent body of the body. For example, a table cell body's parent body could be a header. Returns a null object if there isn't a parent body.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentBodyOrNullObject?: Word.Interfaces.BodyLoadOptions;
            /**
            * Gets the content control that contains the body. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the content control that contains the body. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the parent section of the body. Throws an error if there isn't a parent section.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentSection?: Word.Interfaces.SectionLoadOptions;
            /**
            * Gets the parent section of the body. Returns a null object if there isn't a parent section.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentSectionOrNullObject?: Word.Interfaces.SectionLoadOptions;
            /**
             * Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            /**
             * Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: boolean;
            /**
             * Gets the text of the body. Use the insertText method to insert text. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
            /**
             * Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            type?: boolean;
        }
        /**
         * Represents a comment in the document.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        export interface CommentLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets or sets the comment's content range.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            contentRange?: Word.Interfaces.CommentContentRangeLoadOptions;
            /**
             * Gets the email of the comment's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorEmail?: boolean;
            /**
             * Gets the name of the comment's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorName?: boolean;
            /**
             * Gets or sets the comment's content as plain text.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            content?: boolean;
            /**
             * Gets the creation date of the comment.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            creationDate?: boolean;
            /**
             * Gets the ID of the comment. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            id?: boolean;
            /**
             * Gets or sets the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            resolved?: boolean;
        }
        /**
         * Contains a collection of {@link Word.Comment} objects.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        export interface CommentCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets or sets the comment's content range.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            contentRange?: Word.Interfaces.CommentContentRangeLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets the email of the comment's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorEmail?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the name of the comment's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorName?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the comment's content as plain text.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            content?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the creation date of the comment.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            creationDate?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the ID of the comment. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            resolved?: boolean;
        }
        /**
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        export interface CommentContentRangeLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets or sets a value that indicates whether the comment text is bold.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            bold?: boolean;
            /**
             * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            hyperlink?: boolean;
            /**
             * Checks whether the range length is zero. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            isEmpty?: boolean;
            /**
             * Gets or sets a value that indicates whether the comment text is italicized.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            italic?: boolean;
            /**
             * Gets or sets a value that indicates whether the comment text has a strikethrough.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            strikeThrough?: boolean;
            /**
             * Gets the text of the comment range. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            text?: boolean;
            /**
             * Gets or sets a value that indicates the comment text's underline type. 'None' if the comment text is not underlined.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            underline?: boolean;
        }
        /**
         * Represents a comment reply in the document.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        export interface CommentReplyLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets or sets the commentReply's content range.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            contentRange?: Word.Interfaces.CommentContentRangeLoadOptions;
            /**
            * Gets the parent comment of this reply.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            parentComment?: Word.Interfaces.CommentLoadOptions;
            /**
             * Gets the email of the comment reply's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorEmail?: boolean;
            /**
             * Gets the name of the comment reply's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorName?: boolean;
            /**
             * Gets or sets the comment reply's content. The string is plain text.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            content?: boolean;
            /**
             * Gets the creation date of the comment reply.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            creationDate?: boolean;
            /**
             * Gets the ID of the comment reply. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            id?: boolean;
        }
        /**
         * Contains a collection of {@link Word.CommentReply} objects. Represents all comment replies in one comment thread.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        export interface CommentReplyCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets or sets the commentReply's content range.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            contentRange?: Word.Interfaces.CommentContentRangeLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the parent comment of this reply.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            parentComment?: Word.Interfaces.CommentLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets the email of the comment reply's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorEmail?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the name of the comment reply's author.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            authorName?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the comment reply's content. The string is plain text.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            content?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the creation date of the comment reply.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            creationDate?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the ID of the comment reply. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            id?: boolean;
        }
        /**
         * Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface ContentControlLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * Gets the parent body of the content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * Gets the content control that contains the content control. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the content control that contains the content control. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the table that contains the content control. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * Gets the table cell that contains the content control. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table that contains the content control. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            appearance?: boolean;
            /**
             * Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotDelete?: boolean;
            /**
             * Gets or sets a value that indicates whether the user can edit the contents of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotEdit?: boolean;
            /**
             * Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: boolean;
            /**
             * Gets an integer that represents the content control identifier. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            id?: boolean;
            /**
             * Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             * **Note**: The set operation for this property is not supported in Word on the web.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            placeholderText?: boolean;
            /**
             * Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            removeWhenEdited?: boolean;
            /**
             * Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            /**
             * Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: boolean;
            /**
             * Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            subtype?: boolean;
            /**
             * Gets or sets a tag to identify a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            tag?: boolean;
            /**
             * Gets the text of the content control. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
            /**
             * Gets or sets the title for a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            title?: boolean;
            /**
             * Gets the content control type. Only rich text content controls are supported currently. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            type?: boolean;
        }
        /**
         * Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface ContentControlCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the parent body of the content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the content control. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the content control. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains the content control. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains the content control. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains the content control. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            appearance?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotDelete?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets a value that indicates whether the user can edit the contents of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotEdit?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: boolean;
            /**
             * For EACH ITEM in the collection: Gets an integer that represents the content control identifier. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             * **Note**: The set operation for this property is not supported in Word on the web.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            placeholderText?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            removeWhenEdited?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            subtype?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets a tag to identify a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            tag?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the text of the content control. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the title for a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            title?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the content control type. Only rich text content controls are supported currently. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            type?: boolean;
        }
        /**
         * Represents a custom property.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface CustomPropertyLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets the key of the custom property. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            key?: boolean;
            /**
             * Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            type?: boolean;
            /**
             * Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            value?: boolean;
        }
        /**
         * Contains the collection of {@link Word.CustomProperty} objects.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface CustomPropertyCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the key of the custom property. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            key?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            type?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            value?: boolean;
        }
        /**
         * Represents a custom XML part.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface CustomXmlPartLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets the ID of the custom XML part. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: boolean;
            /**
             * Gets the namespace URI of the custom XML part. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            namespaceUri?: boolean;
        }
        /**
         * Contains the collection of {@link Word.CustomXmlPart} objects.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface CustomXmlPartCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the ID of the custom XML part. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the namespace URI of the custom XML part. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            namespaceUri?: boolean;
        }
        /**
         * Contains the collection of {@link Word.CustomXmlPart} objects with a specific namespace.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface CustomXmlPartScopedCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the ID of the custom XML part. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the namespace URI of the custom XML part. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            namespaceUri?: boolean;
        }
        /**
         * The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface DocumentLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyLoadOptions;
            /**
            * Gets the properties of the document.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            properties?: Word.Interfaces.DocumentPropertiesLoadOptions;
            /**
             * Gets or sets the ChangeTracking mode.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            changeTrackingMode?: boolean;
            /**
             * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            saved?: boolean;
        }
        /**
         * The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface DocumentCreatedLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
            *
            * @remarks
            * [Api set: WordApiHiddenDocument 1.3]
            */
            body?: Word.Interfaces.BodyLoadOptions;
            /**
            * Gets the properties of the document.
            *
            * @remarks
            * [Api set: WordApiHiddenDocument 1.3]
            */
            properties?: Word.Interfaces.DocumentPropertiesLoadOptions;
            /**
             * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
             *
             * @remarks
             * [Api set: WordApiHiddenDocument 1.3]
             */
            saved?: boolean;
        }
        /**
         * Represents document properties.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface DocumentPropertiesLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets the application name of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            applicationName?: boolean;
            /**
             * Gets or sets the author of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            author?: boolean;
            /**
             * Gets or sets the category of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            category?: boolean;
            /**
             * Gets or sets the comments of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            comments?: boolean;
            /**
             * Gets or sets the company of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            company?: boolean;
            /**
             * Gets the creation date of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            creationDate?: boolean;
            /**
             * Gets or sets the format of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            format?: boolean;
            /**
             * Gets or sets the keywords of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            keywords?: boolean;
            /**
             * Gets the last author of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            lastAuthor?: boolean;
            /**
             * Gets the last print date of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            lastPrintDate?: boolean;
            /**
             * Gets the last save time of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            lastSaveTime?: boolean;
            /**
             * Gets or sets the manager of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            manager?: boolean;
            /**
             * Gets the revision number of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            revisionNumber?: boolean;
            /**
             * Gets security settings of the document. Read only. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            security?: boolean;
            /**
             * Gets or sets the subject of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            subject?: boolean;
            /**
             * Gets the template of the document. Read only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            template?: boolean;
            /**
             * Gets or sets the title of the document.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            title?: boolean;
        }
        /**
         * Represents a field.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface FieldLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the parent body of the field.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * Gets the content control that contains the field. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the content control that contains the field. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the table that contains the field. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * Gets the table cell that contains the field. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table cell that contains the field. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table that contains the field. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * Gets the field's code instruction. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
             * @beta
             */
            code?: boolean;
            /**
             * Gets the field's result data. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
             * @beta
             */
            result?: boolean;
        }
        /**
         * Contains a collection of {@link Word.Field} objects.
         *
         * @remarks
         * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface FieldCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the parent body of the field.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the field. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the field. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains the field. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains the field. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains the field. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains the field. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
            * @beta
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets the field's code instruction. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
             * @beta
             */
            code?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the field's result data. Read-only.
             *
             * @remarks
             * [Api set: WordApiOnline BETA (PREVIEW ONLY)]
             * @beta
             */
            result?: boolean;
        }
        /**
         * Represents a font.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface FontLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            bold?: boolean;
            /**
             * Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: boolean;
            /**
             * Gets or sets a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            doubleStrikeThrough?: boolean;
            /**
             * Gets or sets the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color.
             * **Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            highlightColor?: boolean;
            /**
             * Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            italic?: boolean;
            /**
             * Gets or sets a value that represents the name of the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            name?: boolean;
            /**
             * Gets or sets a value that represents the font size in points.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            size?: boolean;
            /**
             * Gets or sets a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            strikeThrough?: boolean;
            /**
             * Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            subscript?: boolean;
            /**
             * Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            superscript?: boolean;
            /**
             * Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            underline?: boolean;
        }
        /**
         * Represents an inline picture.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface InlinePictureLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the parent paragraph that contains the inline image.
            *
            * @remarks
            * [Api set: WordApi 1.2]
            */
            paragraph?: Word.Interfaces.ParagraphLoadOptions;
            /**
            * Gets the content control that contains the inline image. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the table that contains the inline image. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * Gets the table cell that contains the inline image. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table that contains the inline image. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * Gets or sets a string that represents the alternative text associated with the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: boolean;
            /**
             * Gets or sets a string that contains the title for the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: boolean;
            /**
             * Gets or sets a number that describes the height of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            height?: boolean;
            /**
             * Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            hyperlink?: boolean;
            /**
             * Gets the format of the inline image. Read-only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            imageFormat?: boolean;
            /**
             * Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lockAspectRatio?: boolean;
            /**
             * Gets or sets a number that describes the width of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            width?: boolean;
        }
        /**
         * Contains a collection of {@link Word.InlinePicture} objects.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface InlinePictureCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the parent paragraph that contains the inline image.
            *
            * @remarks
            * [Api set: WordApi 1.2]
            */
            paragraph?: Word.Interfaces.ParagraphLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the inline image. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains the inline image. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains the inline image. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains the inline image. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets or sets a string that represents the alternative text associated with the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets a string that contains the title for the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets a number that describes the height of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            height?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            hyperlink?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the format of the inline image. Read-only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            imageFormat?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lockAspectRatio?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets a number that describes the width of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            width?: boolean;
        }
        /**
         * Contains a collection of {@link Word.Paragraph} objects.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface ListLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets the list's id.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            id?: boolean;
            /**
             * Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            levelExistences?: boolean;
            /**
             * Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            levelTypes?: boolean;
        }
        /**
         * Contains a collection of {@link Word.List} objects.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface ListCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the list's id.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            levelExistences?: boolean;
            /**
             * For EACH ITEM in the collection: Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            levelTypes?: boolean;
        }
        /**
         * Represents the paragraph list item format.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface ListItemLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets or sets the level of the item in the list.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            level?: boolean;
            /**
             * Gets the list item bullet, number, or picture as a string. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            listString?: boolean;
            /**
             * Gets the list item order number in relation to its siblings. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            siblingIndex?: boolean;
        }
        /**
         * Represents a footnote or endnote.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        export interface NoteItemLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the body object of the note item. It's the portion of the text within the footnote or endnote.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            body?: Word.Interfaces.BodyLoadOptions;
            /**
            * Represents a footnote or endnote reference in the main document.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            reference?: Word.Interfaces.RangeLoadOptions;
            /**
             * Represents the note item type: footnote or endnote.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            type?: boolean;
        }
        /**
         * Contains a collection of {@link Word.NoteItem} objects.
         *
         * @remarks
         * [Api set: WordApiOnline 1.1]
         */
        export interface NoteItemCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Represents the body object of the note item. It's the portion of the text within the footnote or endnote.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            body?: Word.Interfaces.BodyLoadOptions;
            /**
            * For EACH ITEM in the collection: Represents a footnote or endnote reference in the main document.
            *
            * @remarks
            * [Api set: WordApiOnline 1.1]
            */
            reference?: Word.Interfaces.RangeLoadOptions;
            /**
             * For EACH ITEM in the collection: Represents the note item type: footnote or endnote.
             *
             * @remarks
             * [Api set: WordApiOnline 1.1]
             */
            type?: boolean;
        }
        /**
         * Represents a single paragraph in a selection, range, content control, or document body.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface ParagraphLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * Gets the List to which this paragraph belongs. Throws an error if the paragraph is not in a list.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            list?: Word.Interfaces.ListLoadOptions;
            /**
            * Gets the ListItem for the paragraph. Throws an error if the paragraph is not part of a list.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            listItem?: Word.Interfaces.ListItemLoadOptions;
            /**
            * Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            listItemOrNullObject?: Word.Interfaces.ListItemLoadOptions;
            /**
            * Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            listOrNullObject?: Word.Interfaces.ListLoadOptions;
            /**
            * Gets the parent body of the paragraph.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * Gets the content control that contains the paragraph. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the table that contains the paragraph. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * Gets the table cell that contains the paragraph. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table that contains the paragraph. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            alignment?: boolean;
            /**
             * Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: boolean;
            /**
             * Indicates the paragraph is the last one inside its parent body. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isLastParagraph?: boolean;
            /**
             * Checks whether the paragraph is a list item. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isListItem?: boolean;
            /**
             * Gets or sets the left indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            leftIndent?: boolean;
            /**
             * Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: boolean;
            /**
             * Gets or sets the amount of spacing, in grid lines, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: boolean;
            /**
             * Gets or sets the amount of spacing, in grid lines, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: boolean;
            /**
             * Gets or sets the outline level for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: boolean;
            /**
             * Gets or sets the right indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            rightIndent?: boolean;
            /**
             * Gets or sets the spacing, in points, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: boolean;
            /**
             * Gets or sets the spacing, in points, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: boolean;
            /**
             * Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            /**
             * Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: boolean;
            /**
             * Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            tableNestingLevel?: boolean;
            /**
             * Gets the text of the paragraph. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
        }
        /**
         * Contains a collection of {@link Word.Paragraph} objects.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface ParagraphCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the List to which this paragraph belongs. Throws an error if the paragraph is not in a list.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            list?: Word.Interfaces.ListLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the ListItem for the paragraph. Throws an error if the paragraph is not part of a list.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            listItem?: Word.Interfaces.ListItemLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            listItemOrNullObject?: Word.Interfaces.ListItemLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            listOrNullObject?: Word.Interfaces.ListLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the parent body of the paragraph.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the paragraph. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains the paragraph. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains the paragraph. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains the paragraph. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            alignment?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: boolean;
            /**
             * For EACH ITEM in the collection: Indicates the paragraph is the last one inside its parent body. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isLastParagraph?: boolean;
            /**
             * For EACH ITEM in the collection: Checks whether the paragraph is a list item. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isListItem?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the left indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            leftIndent?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the amount of spacing, in grid lines, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the amount of spacing, in grid lines, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the outline level for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the right indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            rightIndent?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the spacing, in points, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the spacing, in points, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            tableNestingLevel?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the text of the paragraph. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
        }
        /**
         * Represents a contiguous area in a document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface RangeLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * Gets the parent body of the range.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * Gets the content control that contains the range. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the content control that contains the range. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the table that contains the range. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * Gets the table cell that contains the range. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table that contains the range. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            hyperlink?: boolean;
            /**
             * Checks whether the range length is zero. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isEmpty?: boolean;
            /**
             * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            /**
             * Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: boolean;
            /**
             * Gets the text of the range. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
        }
        /**
         * Contains a collection of {@link Word.Range} objects.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface RangeCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the parent body of the range.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the range. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the range. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains the range. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains the range. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains the range. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            hyperlink?: boolean;
            /**
             * For EACH ITEM in the collection: Checks whether the range length is zero. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isEmpty?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the text of the range. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
        }
        /**
         * Specifies the options to be included in a search operation.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface SearchOptionsLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets or sets a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignorePunct?: boolean;
            /**
             * Gets or sets a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignoreSpace?: boolean;
            /**
             * Gets or sets a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchCase?: boolean;
            /**
             * Gets or sets a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchPrefix?: boolean;
            /**
             * Gets or sets a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchSuffix?: boolean;
            /**
             * Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWholeWord?: boolean;
            /**
             * Gets or sets a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWildcards?: boolean;
        }
        /**
         * Represents a section in a Word document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface SectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the body object of the section. This does not include the header/footer and other section metadata.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyLoadOptions;
        }
        /**
         * Contains the collection of the document's {@link Word.Section} objects.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface SectionCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the body object of the section. This does not include the header/footer and other section metadata.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyLoadOptions;
        }
        /**
         * Represents a setting of the add-in.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface SettingLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets the key of the setting. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            key?: boolean;
            /**
             * Gets or sets the value of the setting.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            value?: boolean;
        }
        /**
         * Contains the collection of {@link Word.Setting} objects.
         *
         * @remarks
         * [Api set: WordApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface SettingCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the key of the setting. Read only.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            key?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the value of the setting.
             *
             * @remarks
             * [Api set: WordApi BETA (PREVIEW ONLY)]
             * @beta
             */
            value?: boolean;
        }
        /**
         * Represents a table in a Word document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface TableLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the font. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * Gets the parent body of the table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * Gets the content control that contains the table. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the content control that contains the table. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * Gets the table that contains this table. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * Gets the table cell that contains this table. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * Gets the table that contains this table. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            alignment?: boolean;
            /**
             * Gets and sets the number of header rows.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            headerRowCount?: boolean;
            /**
             * Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: boolean;
            /**
             * Indicates whether all of the table rows are uniform. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isUniform?: boolean;
            /**
             * Gets the nesting level of the table. Top-level tables have level 1. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            nestingLevel?: boolean;
            /**
             * Gets the number of rows in the table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            rowCount?: boolean;
            /**
             * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: boolean;
            /**
             * Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            style?: boolean;
            /**
             * Gets and sets whether the table has banded columns.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBandedColumns?: boolean;
            /**
             * Gets and sets whether the table has banded rows.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBandedRows?: boolean;
            /**
             * Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: boolean;
            /**
             * Gets and sets whether the table has a first column with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleFirstColumn?: boolean;
            /**
             * Gets and sets whether the table has a last column with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleLastColumn?: boolean;
            /**
             * Gets and sets whether the table has a total (last) row with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleTotalRow?: boolean;
            /**
             * Gets and sets the text values in the table, as a 2D Javascript array.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            values?: boolean;
            /**
             * Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: boolean;
            /**
             * Gets and sets the width of the table in points.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            width?: boolean;
        }
        /**
         * Contains the collection of the document's Table objects.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface TableCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the font. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the parent body of the table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentBody?: Word.Interfaces.BodyLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the table. Throws an error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the content control that contains the table. Returns a null object if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains this table. Throws an error if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains this table. Throws an error if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCell?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the table that contains this table. Returns a null object if it is not contained in a table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            alignment?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the number of header rows.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            headerRowCount?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: boolean;
            /**
             * For EACH ITEM in the collection: Indicates whether all of the table rows are uniform. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isUniform?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the nesting level of the table. Top-level tables have level 1. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            nestingLevel?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the number of rows in the table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            rowCount?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            style?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets whether the table has banded columns.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBandedColumns?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets whether the table has banded rows.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBandedRows?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleBuiltIn?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets whether the table has a first column with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleFirstColumn?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets whether the table has a last column with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleLastColumn?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets whether the table has a total (last) row with a special style.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            styleTotalRow?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the text values in the table, as a 2D Javascript array.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            values?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the width of the table in points.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            width?: boolean;
        }
        /**
         * Represents a row in a Word document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface TableRowLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the font. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * Gets parent table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
             * Gets the number of cells in the row. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            cellCount?: boolean;
            /**
             * Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: boolean;
            /**
             * Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isHeader?: boolean;
            /**
             * Gets and sets the preferred height of the row in points.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            preferredHeight?: boolean;
            /**
             * Gets the index of the row in its parent table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            rowIndex?: boolean;
            /**
             * Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: boolean;
            /**
             * Gets and sets the text values in the row, as a 2D Javascript array.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            values?: boolean;
            /**
             * Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: boolean;
        }
        /**
         * Contains the collection of the document's TableRow objects.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface TableRowCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the font. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            font?: Word.Interfaces.FontLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets parent table.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets the number of cells in the row. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            cellCount?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: boolean;
            /**
             * For EACH ITEM in the collection: Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            isHeader?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the preferred height of the row in points.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            preferredHeight?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the index of the row in its parent table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            rowIndex?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the text values in the row, as a 2D Javascript array.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            values?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: boolean;
        }
        /**
         * Represents a table cell in a Word document.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface TableCellLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the body object of the cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            body?: Word.Interfaces.BodyLoadOptions;
            /**
            * Gets the parent row of the cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentRow?: Word.Interfaces.TableRowLoadOptions;
            /**
            * Gets the parent table of the cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
             * Gets the index of the cell in its row. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            cellIndex?: boolean;
            /**
             * Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            columnWidth?: boolean;
            /**
             * Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: boolean;
            /**
             * Gets the index of the cell's row in the table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            rowIndex?: boolean;
            /**
             * Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: boolean;
            /**
             * Gets and sets the text of the cell.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            value?: boolean;
            /**
             * Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: boolean;
            /**
             * Gets the width of the cell in points. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            width?: boolean;
        }
        /**
         * Contains the collection of the document's TableCell objects.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface TableCellCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the body object of the cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            body?: Word.Interfaces.BodyLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the parent row of the cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentRow?: Word.Interfaces.TableRowLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the parent table of the cell.
            *
            * @remarks
            * [Api set: WordApi 1.3]
            */
            parentTable?: Word.Interfaces.TableLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets the index of the cell in its row. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            cellIndex?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            columnWidth?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            horizontalAlignment?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the index of the cell's row in the table. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            rowIndex?: boolean;
            /**
             * For EACH ITEM in the collection: Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            shadingColor?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the text of the cell.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            value?: boolean;
            /**
             * For EACH ITEM in the collection: Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            verticalAlignment?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the width of the cell in points. Read-only.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            width?: boolean;
        }
        /**
         * Specifies the border style.
         *
         * @remarks
         * [Api set: WordApi 1.3]
         */
        export interface TableBorderLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets or sets the table border color.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            color?: boolean;
            /**
             * Gets or sets the type of the table border.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            type?: boolean;
            /**
             * Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.
             *
             * @remarks
             * [Api set: WordApi 1.3]
             */
            width?: boolean;
        }
    }
}
export declare namespace Word {
    /**
     * The RequestContext object facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the request context is required to get access to the Word object model from the add-in.
     */
    export class RequestContext extends OfficeExtension.ClientRequestContext {
        constructor(url?: string);
        readonly document: Document;
        /** [Api set: WordApi 1.3] **/
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
     */
    export function run<T>(batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;
}


////////////////////////////////////////////////////////////////
//////////////////////// End Word APIs /////////////////////////
////////////////////////////////////////////////////////////////