import { OfficeExtension } from "../../api-extractor-inputs-office/office"
import { Office as Outlook} from "../../api-extractor-inputs-outlook/outlook"
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
         * Gets the collection of paragraph objects in the body. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly paragraphs: Word.ParagraphCollection;
        
        
        /**
         *
         * Gets the content control that contains the body. Throws an error if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        
        
        
        
        /**
         *
         * Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.1]
         */
        style: string;
        
        /**
         *
         * Gets the text of the body. Use the insertText method to insert text. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Word.Body): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.BodyUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Body): void;
        /**
         *
         * Clears the contents of the body object. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        /**
         *
         * Gets an HTML representation of the body object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word for the web, etc.). If you need exact fidelity, or consistency across platforms, use `Body.getOoxml()` and convert the returned XML to HTML.
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
         * Inserts a break at the specified location in the main document.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add to the body.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation): void;
        /**
         *
         * Inserts a break at the specified location in the main document.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakTypeString - Required. The break type to add to the body.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertBreak(breakTypeString: "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): void;
        /**
         *
         * Wraps the body object with a Rich Text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document into the body at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts a document into the body at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts HTML at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the document.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts HTML at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the document.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertHtml(html: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a picture into the body at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the body.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation): Word.InlinePicture;
        /**
         *
         * Inserts a picture into the body at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the body.
         * @param insertLocationString - Required. The value can be 'Start' or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts OOXML at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertOoxml(ooxml: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start' or 'End'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocationString - Required. The value can be 'Start' or 'End'.
         */
        insertParagraph(paragraphText: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
        
        
        /**
         *
         * Inserts text into the body at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts text into the body at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertText(text: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
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
         * @param selectionModeString - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.Body` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Body` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Body` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.BodyLoadOptions): Word.Body;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Body;
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
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Body object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BodyData`) that contains shallow copies of any loaded child properties from the original object.
        */
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
         * Get the collection of paragraph objects in the content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly paragraphs: Word.ParagraphCollection;
        
        /**
         *
         * Gets the content control that contains the content control. Throws an error if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        
        
        
        
        
        
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
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Word.ContentControl): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ContentControlUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.ContentControl): void;
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
         * Gets an HTML representation of the content control object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word for the web, etc.). If you need exact fidelity, or consistency across platforms, use `ContentControl.getOoxml()` and convert the returned XML to HTML.
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
         * Inserts a break at the specified location in the main document. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. Type of break.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'.
         */
        insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation): void;
        /**
         *
         * Inserts a break at the specified location in the main document. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakTypeString - Required. Type of break.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'.
         */
        insertBreak(breakTypeString: "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): void;
        /**
         *
         * Inserts a document into the content control at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts a document into the content control at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertFileFromBase64(base64File: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts HTML into the content control at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts HTML into the content control at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in to the content control.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertHtml(html: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts an inline picture into the content control at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation): Word.InlinePicture;
        /**
         *
         * Inserts an inline picture into the content control at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted in the content control.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.InlinePicture;
        /**
         *
         * Inserts OOXML into the content control at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts OOXML into the content control at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in to the content control.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertOoxml(ooxml: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocationString - Required. The value can be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
         */
        insertParagraph(paragraphText: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
        
        
        /**
         *
         * Inserts text into the content control at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. The text to be inserted in to the content control.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertText(text: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts text into the content control at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. The text to be inserted in to the content control.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.
         */
        insertText(text: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
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
         * @param selectionModeString - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.ContentControl` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.ContentControl` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.ContentControl` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.ContentControlLoadOptions): Word.ContentControl;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.ContentControl;
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
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.ContentControl object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlData`) that contains shallow copies of any loaded child properties from the original object.
        */
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
         * Gets a content control by its identifier. Throws an error if there isn't a content control with the identifier in this collection.
         *
         * [Api set: WordApi 1.1]
         *
         * @param id - Required. A content control identifier.
         */
        getById(id: number): Word.ContentControl;
        
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
         * `load(option?: string | string[]): Word.ContentControlCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.ContentControlCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.ContentControlCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.ContentControlCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.ContentControlCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.ContentControlCollection;
        load(option?: OfficeExtension.LoadOption): Word.ContentControlCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ContentControlCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ContentControlCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.ContentControlCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.ContentControlCollectionData;
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
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Word.Document): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.DocumentUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Document): void;
        /**
         *
         * Gets the current selection of the document. Multiple selections are not supported.
         *
         * [Api set: WordApi 1.1]
         */
        getSelection(): Word.Range;
        /**
         *
         * Saves the document. This uses the Word default file naming convention if the document has not been saved before.
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
         * `load(option?: string | string[]): Word.Document` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Document` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Document` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.DocumentLoadOptions): Word.Document;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Document;
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
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Document object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.DocumentData;
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
            **Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
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
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Word.Font): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.FontUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Font): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.Font` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Font` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Font` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.FontLoadOptions): Word.Font;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Font;
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
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Font object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FontData`) that contains shallow copies of any loaded child properties from the original object.
        */
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
         * Gets the content control that contains the inline image. Throws an error if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        
        
        
        
        
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
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Word.InlinePicture): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.InlinePictureUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.InlinePicture): void;
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
         * Inserts a break at the specified location in the main document.
         *
         * [Api set: WordApi 1.2]
         *
         * @param breakType - Required. The break type to add.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation): void;
        /**
         *
         * Inserts a break at the specified location in the main document.
         *
         * [Api set: WordApi 1.2]
         *
         * @param breakTypeString - Required. The break type to add.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakTypeString: "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): void;
        /**
         *
         * Wraps the inline picture with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts a document at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocationString - Required. The value can be 'Before' or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts HTML at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts HTML at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocationString - Required. The value can be 'Before' or 'After'.
         */
        insertHtml(html: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts an inline picture at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Before', or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation): Word.InlinePicture;
        /**
         *
         * Inserts an inline picture at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocationString - Required. The value can be 'Replace', 'Before', or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts OOXML at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocationString - Required. The value can be 'Before' or 'After'.
         */
        insertOoxml(ooxml: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocationString - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
        /**
         *
         * Inserts text at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts text at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocationString - Required. The value can be 'Before' or 'After'.
         */
        insertText(text: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
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
         * @param selectionModeString - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.InlinePicture` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.InlinePicture` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.InlinePicture` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.InlinePictureLoadOptions): Word.InlinePicture;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.InlinePicture;
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
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.InlinePicture object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.InlinePictureData`) that contains shallow copies of any loaded child properties from the original object.
        */
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
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.InlinePictureCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.InlinePictureCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.InlinePictureCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.InlinePictureCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.InlinePictureCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.InlinePictureCollection;
        load(option?: OfficeExtension.LoadOption): Word.InlinePictureCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.InlinePictureCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.InlinePictureCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.InlinePictureCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.InlinePictureCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.InlinePictureCollectionData;
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
         * Gets the content control that contains the paragraph. Throws an error if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        
        
        
        
        
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
         * Gets the text of the paragraph. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Word.Paragraph): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ParagraphUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Paragraph): void;
        
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
         * Gets an HTML representation of the paragraph object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word for the web, etc.). If you need exact fidelity, or consistency across platforms, use `Paragraph.getOoxml()` and convert the returned XML to HTML.
         *
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        
        
        /**
         *
         * Gets the Office Open XML (OOXML) representation of the paragraph object.
         *
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        
        
        
        
        
        /**
         *
         * Inserts a break at the specified location in the main document.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add to the document.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation): void;
        /**
         *
         * Inserts a break at the specified location in the main document.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakTypeString - Required. The break type to add to the document.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakTypeString: "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): void;
        /**
         *
         * Wraps the paragraph object with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document into the paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts a document into the paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertFileFromBase64(base64File: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts HTML into the paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the paragraph.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts HTML into the paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted in the paragraph.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertHtml(html: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a picture into the paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation): Word.InlinePicture;
        /**
         *
         * Inserts a picture into the paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.InlinePicture;
        /**
         *
         * Inserts OOXML into the paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in the paragraph.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts OOXML into the paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted in the paragraph.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertOoxml(ooxml: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocationString - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
        
        
        /**
         *
         * Inserts text into the paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts text into the paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', or 'End'.
         */
        insertText(text: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
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
         * @param selectionModeString - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.Paragraph` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Paragraph` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Paragraph` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.ParagraphLoadOptions): Word.Paragraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Paragraph;
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
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Paragraph object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ParagraphData`) that contains shallow copies of any loaded child properties from the original object.
        */
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
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.ParagraphCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.ParagraphCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.ParagraphCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.ParagraphCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.ParagraphCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.ParagraphCollection;
        load(option?: OfficeExtension.LoadOption): Word.ParagraphCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.ParagraphCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.ParagraphCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.ParagraphCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ParagraphCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
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
         * Gets the collection of paragraph objects in the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly paragraphs: Word.ParagraphCollection;
        
        /**
         *
         * Gets the content control that contains the range. Throws an error if there isn't a parent content control. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        
        
        
        
        
        
        
        
        /**
         *
         * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * [Api set: WordApi 1.1]
         */
        style: string;
        
        /**
         *
         * Gets the text of the range. Read-only.
         *
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Word.Range): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.RangeUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Range): void;
        /**
         *
         * Clears the contents of the range object. The user can perform the undo operation on the cleared content.
         *
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        
        /**
         *
         * Deletes the range and its content from the document.
         *
         * [Api set: WordApi 1.1]
         */
        delete(): void;
        
        
        /**
         *
         * Gets an HTML representation of the range object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method does not return the exact same HTML for the same document on different platforms (Windows, Mac, Word for the web, etc.). If you need exact fidelity, or consistency across platforms, use `Range.getOoxml()` and convert the returned XML to HTML.
         *
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        
        
        
        /**
         *
         * Gets the OOXML representation of the range object.
         *
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        
        
        
        /**
         *
         * Inserts a break at the specified location in the main document.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Required. The break type to add.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation): void;
        /**
         *
         * Inserts a break at the specified location in the main document.
         *
         * [Api set: WordApi 1.1]
         *
         * @param breakTypeString - Required. The break type to add.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertBreak(breakTypeString: "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: "Before" | "After" | "Start" | "End" | "Replace"): void;
        /**
         *
         * Wraps the range object with a rich text content control.
         *
         * [Api set: WordApi 1.1]
         */
        insertContentControl(): Word.ContentControl;
        /**
         *
         * Inserts a document at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts a document at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param base64File - Required. The base64 encoded content of a .docx file.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertFileFromBase64(base64File: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts HTML at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts HTML at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param html - Required. The HTML to be inserted.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertHtml(html: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a picture at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation): Word.InlinePicture;
        /**
         *
         * Inserts a picture at the specified location.
         *
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - Required. The base64 encoded image to be inserted.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.InlinePicture;
        /**
         *
         * Inserts OOXML at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts OOXML at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - Required. The OOXML to be inserted.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertOoxml(ooxml: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        /**
         *
         * Inserts a paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocation - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
        /**
         *
         * Inserts a paragraph at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - Required. The paragraph text to be inserted.
         * @param insertLocationString - Required. The value can be 'Before' or 'After'.
         */
        insertParagraph(paragraphText: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Paragraph;
        
        
        /**
         *
         * Inserts text at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocation - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertText(text: string, insertLocation: Word.InsertLocation): Word.Range;
        /**
         *
         * Inserts text at the specified location.
         *
         * [Api set: WordApi 1.1]
         *
         * @param text - Required. Text to be inserted.
         * @param insertLocationString - Required. The value can be 'Replace', 'Start', 'End', 'Before', or 'After'.
         */
        insertText(text: string, insertLocationString: "Before" | "After" | "Start" | "End" | "Replace"): Word.Range;
        
        
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
         * @param selectionModeString - Optional. The selection mode can be 'Select', 'Start', or 'End'. 'Select' is the default.
         */
        select(selectionModeString?: "Select" | "Start" | "End"): void;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.Range` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Range` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Range` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.RangeLoadOptions): Word.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Range;
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
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Range object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeData`) that contains shallow copies of any loaded child properties from the original object.
        */
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
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.RangeCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.RangeCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.RangeCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.RangeCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.RangeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.RangeCollection;
        load(option?: OfficeExtension.LoadOption): Word.RangeCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.RangeCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.RangeCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.RangeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
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
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Word.SearchOptions): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.SearchOptionsUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.SearchOptions): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.SearchOptions` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.SearchOptions` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.SearchOptions` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.SearchOptionsLoadOptions): Word.SearchOptions;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.SearchOptions;
        load(option?: {
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
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Word.Section): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.SectionUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Section): void;
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
         * @param typeString - Required. The type of footer to return. This value can be: 'Primary', 'FirstPage', or 'EvenPages'.
         */
        getFooter(typeString: "Primary" | "FirstPage" | "EvenPages"): Word.Body;
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
         * @param typeString - Required. The type of header to return. This value can be: 'Primary', 'FirstPage', or 'EvenPages'.
         */
        getHeader(typeString: "Primary" | "FirstPage" | "EvenPages"): Word.Body;
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.Section` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.Section` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.Section` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.SectionLoadOptions): Word.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.Section;
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
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Word.Section object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionData`) that contains shallow copies of any loaded child properties from the original object.
        */
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
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Word.SectionCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Word.SectionCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Word.SectionCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Word.Interfaces.SectionCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.SectionCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Word.SectionCollection;
        load(option?: OfficeExtension.LoadOption): Word.SectionCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): Word.SectionCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): Word.SectionCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Word.SectionCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.SectionCollectionData;
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
        /**
         *
         * AnnotationAdded represents the event an annotation has been added to the document.
         *
         */
        annotationAdded = "AnnotationAdded",
        /**
         *
         * AnnotationAdded represents the event an annotation has been updated in the document.
         *
         */
        annotationChanged = "AnnotationChanged",
        /**
         *
         * AnnotationAdded represents the event an annotation has been deleted from the document.
         *
         */
        annotationDeleted = "AnnotationDeleted",
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
         * Returns the header or footer on all pages of a section, but excludes the first page or odd pages if they are different.
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
            
        }
        /** An interface for updating data on the CustomPropertyCollection object, for use in "customPropertyCollection.set({ ... })". */
        export interface CustomPropertyCollectionUpdateData {
            items?: Word.Interfaces.CustomPropertyData[];
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
            **Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
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
             * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
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
        /** An interface for updating data on the Table object, for use in "table.set({ ... })". */
        export interface TableUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the TableCollection object, for use in "tableCollection.set({ ... })". */
        export interface TableCollectionUpdateData {
            items?: Word.Interfaces.TableData[];
        }
        /** An interface for updating data on the TableRow object, for use in "tableRow.set({ ... })". */
        export interface TableRowUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the TableRowCollection object, for use in "tableRowCollection.set({ ... })". */
        export interface TableRowCollectionUpdateData {
            items?: Word.Interfaces.TableRowData[];
        }
        /** An interface for updating data on the TableCell object, for use in "tableCell.set({ ... })". */
        export interface TableCellUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the TableCellCollection object, for use in "tableCellCollection.set({ ... })". */
        export interface TableCellCollectionUpdateData {
            items?: Word.Interfaces.TableCellData[];
        }
        /** An interface for updating data on the TableBorder object, for use in "tableBorder.set({ ... })". */
        export interface TableBorderUpdateData {
            
            
            
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
            * Gets the collection of paragraph objects in the body. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            
            /**
             *
             * Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
            /**
             *
             * Gets the text of the body. Use the insertText method to insert text. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: string;
            
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
            * Get the collection of paragraph objects in the content control. Read-only.
            *
            * [Api set: WordApi 1.1]
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            
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
            
            
            
        }
        /** An interface describing the data returned by calling "customPropertyCollection.toJSON()". */
        export interface CustomPropertyCollectionData {
            items?: Word.Interfaces.CustomPropertyData[];
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
             * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
             *
             * [Api set: WordApiHiddenDocument 1.3]
             */
            saved?: boolean;
        }
        /** An interface describing the data returned by calling "documentProperties.toJSON()". */
        export interface DocumentPropertiesData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
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
            **Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
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
            
            
            
            
        }
        /** An interface describing the data returned by calling "listCollection.toJSON()". */
        export interface ListCollectionData {
            items?: Word.Interfaces.ListData[];
        }
        /** An interface describing the data returned by calling "listItem.toJSON()". */
        export interface ListItemData {
            
            
            
        }
        /** An interface describing the data returned by calling "paragraph.toJSON()". */
        export interface ParagraphData {
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
             * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
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
        /** An interface describing the data returned by calling "table.toJSON()". */
        export interface TableData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling "tableCollection.toJSON()". */
        export interface TableCollectionData {
            items?: Word.Interfaces.TableData[];
        }
        /** An interface describing the data returned by calling "tableRow.toJSON()". */
        export interface TableRowData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling "tableRowCollection.toJSON()". */
        export interface TableRowCollectionData {
            items?: Word.Interfaces.TableRowData[];
        }
        /** An interface describing the data returned by calling "tableCell.toJSON()". */
        export interface TableCellData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling "tableCellCollection.toJSON()". */
        export interface TableCellCollectionData {
            items?: Word.Interfaces.TableCellData[];
        }
        /** An interface describing the data returned by calling "tableBorder.toJSON()". */
        export interface TableBorderData {
            
            
            
        }
        /**
         *
         * Represents the body of a document or a section.
         *
         * [Api set: WordApi 1.1]
         */
        export interface BodyLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the text format of the body. Use this to get and set font name, size, color and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            
            
            /**
            *
            * Gets the content control that contains the body. Throws an error if there isn't a parent content control.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            /**
             *
             * Gets or sets the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            /**
             *
             * Gets the text of the body. Use the insertText method to insert text. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
            
        }
        /**
         *
         * Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
         *
         * [Api set: WordApi 1.1]
         */
        export interface ContentControlLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            
            /**
            *
            * Gets the content control that contains the content control. Throws an error if there isn't a parent content control.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            /**
             *
             * Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
             *
             * [Api set: WordApi 1.1]
             */
            appearance?: boolean;
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
            color?: boolean;
            /**
             *
             * Gets an integer that represents the content control identifier. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             *
             * [Api set: WordApi 1.1]
             */
            placeholderText?: boolean;
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
            style?: boolean;
            
            
            /**
             *
             * Gets or sets a tag to identify a content control.
             *
             * [Api set: WordApi 1.1]
             */
            tag?: boolean;
            /**
             *
             * Gets the text of the content control. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
            /**
             *
             * Gets or sets the title for a content control.
             *
             * [Api set: WordApi 1.1]
             */
            title?: boolean;
            /**
             *
             * Gets the content control type. Only rich text content controls are supported currently. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            type?: boolean;
        }
        /**
         *
         * Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.
         *
         * [Api set: WordApi 1.1]
         */
        export interface ContentControlCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            
            /**
            *
            * For EACH ITEM in the collection: Gets the content control that contains the content control. Throws an error if there isn't a parent content control.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
             *
             * [Api set: WordApi 1.1]
             */
            appearance?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
             *
             * [Api set: WordApi 1.1]
             */
            cannotDelete?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets a value that indicates whether the user can edit the contents of the content control.
             *
             * [Api set: WordApi 1.1]
             */
            cannotEdit?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
             *
             * [Api set: WordApi 1.1]
             */
            color?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets an integer that represents the content control identifier. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             *
             * [Api set: WordApi 1.1]
             */
            placeholderText?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
             *
             * [Api set: WordApi 1.1]
             */
            removeWhenEdited?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            
            /**
             *
             * For EACH ITEM in the collection: Gets or sets a tag to identify a content control.
             *
             * [Api set: WordApi 1.1]
             */
            tag?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the text of the content control. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the title for a content control.
             *
             * [Api set: WordApi 1.1]
             */
            title?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the content control type. Only rich text content controls are supported currently. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            type?: boolean;
        }
        
        
        /**
         *
         * The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.
         *
         * [Api set: WordApi 1.1]
         */
        export interface DocumentLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc..
            *
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyLoadOptions;
            
            /**
             *
             * Gets or sets a value that indicates that, when opening a new document, whether it is allowed to close this document even if this document is untitled. True to close, false otherwise.
             *
             * [Api set: WordApi]
             */
            allowCloseOnUntitled?: boolean;
            /**
             *
             * Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            saved?: boolean;
        }
        
        
        /**
         *
         * Represents a font.
         *
         * [Api set: WordApi 1.1]
         */
        export interface FontLoadOptions {
            $all?: boolean;
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
            color?: boolean;
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
            **Note**: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
             *
             * [Api set: WordApi 1.1]
             */
            highlightColor?: boolean;
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
            name?: boolean;
            /**
             *
             * Gets or sets a value that represents the font size in points.
             *
             * [Api set: WordApi 1.1]
             */
            size?: boolean;
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
            underline?: boolean;
        }
        /**
         *
         * Represents an inline picture.
         *
         * [Api set: WordApi 1.1]
         */
        export interface InlinePictureLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the parent paragraph that contains the inline image.
            *
            * [Api set: WordApi 1.2]
            */
            paragraph?: Word.Interfaces.ParagraphLoadOptions;
            /**
            *
            * Gets the content control that contains the inline image. Throws an error if there isn't a parent content control.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            /**
             *
             * Gets or sets a string that represents the alternative text associated with the inline image.
             *
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: boolean;
            /**
             *
             * Gets or sets a string that contains the title for the inline image.
             *
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: boolean;
            /**
             *
             * Gets or sets a number that describes the height of the inline image.
             *
             * [Api set: WordApi 1.1]
             */
            height?: boolean;
            /**
             *
             * Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * [Api set: WordApi 1.1]
             */
            hyperlink?: boolean;
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
            width?: boolean;
        }
        /**
         *
         * Contains a collection of {@link Word.InlinePicture} objects.
         *
         * [Api set: WordApi 1.1]
         */
        export interface InlinePictureCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the parent paragraph that contains the inline image.
            *
            * [Api set: WordApi 1.2]
            */
            paragraph?: Word.Interfaces.ParagraphLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the content control that contains the inline image. Throws an error if there isn't a parent content control.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            /**
             *
             * For EACH ITEM in the collection: Gets or sets a string that represents the alternative text associated with the inline image.
             *
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets a string that contains the title for the inline image.
             *
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets a number that describes the height of the inline image.
             *
             * [Api set: WordApi 1.1]
             */
            height?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * [Api set: WordApi 1.1]
             */
            hyperlink?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.
             *
             * [Api set: WordApi 1.1]
             */
            lockAspectRatio?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets a number that describes the width of the inline image.
             *
             * [Api set: WordApi 1.1]
             */
            width?: boolean;
        }
        
        
        
        /**
         *
         * Represents a single paragraph in a selection, range, content control, or document body.
         *
         * [Api set: WordApi 1.1]
         */
        export interface ParagraphLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            
            
            
            
            
            /**
            *
            * Gets the content control that contains the paragraph. Throws an error if there isn't a parent content control.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            /**
             *
             * Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * [Api set: WordApi 1.1]
             */
            alignment?: boolean;
            /**
             *
             * Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: boolean;
            
            
            /**
             *
             * Gets or sets the left indent value, in points, for the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            leftIndent?: boolean;
            /**
             *
             * Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: boolean;
            /**
             *
             * Gets or sets the amount of spacing, in grid lines, after the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: boolean;
            /**
             *
             * Gets or sets the amount of spacing, in grid lines, before the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: boolean;
            /**
             *
             * Gets or sets the outline level for the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: boolean;
            /**
             *
             * Gets or sets the right indent value, in points, for the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            rightIndent?: boolean;
            /**
             *
             * Gets or sets the spacing, in points, after the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: boolean;
            /**
             *
             * Gets or sets the spacing, in points, before the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: boolean;
            /**
             *
             * Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            
            /**
             *
             * Gets the text of the paragraph. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
        }
        /**
         *
         * Contains a collection of {@link Word.Paragraph} objects.
         *
         * [Api set: WordApi 1.1]
         */
        export interface ParagraphCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            
            
            
            
            
            /**
            *
            * For EACH ITEM in the collection: Gets the content control that contains the paragraph. Throws an error if there isn't a parent content control.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
             *
             * [Api set: WordApi 1.1]
             */
            alignment?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: boolean;
            
            
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the left indent value, in points, for the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            leftIndent?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the amount of spacing, in grid lines, after the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the amount of spacing, in grid lines, before the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the outline level for the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the right indent value, in points, for the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            rightIndent?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the spacing, in points, after the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the spacing, in points, before the paragraph.
             *
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            
            /**
             *
             * For EACH ITEM in the collection: Gets the text of the paragraph. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
        }
        /**
         *
         * Represents a contiguous area in a document.
         *
         * [Api set: WordApi 1.1]
         */
        export interface RangeLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            
            /**
            *
            * Gets the content control that contains the range. Throws an error if there isn't a parent content control.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            
            
            /**
             *
             * Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            /**
             *
             * Gets the text of the range. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
        }
        /**
         *
         * Contains a collection of {@link Word.Range} objects.
         *
         * [Api set: WordApi 1.1]
         */
        export interface RangeCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            *
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            
            /**
            *
            * For EACH ITEM in the collection: Gets the content control that contains the range. Throws an error if there isn't a parent content control.
            *
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            
            
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            /**
             *
             * For EACH ITEM in the collection: Gets the text of the range. Read-only.
             *
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
        }
        /**
         *
         * Specifies the options to be included in a search operation.
         *
         * [Api set: WordApi 1.1]
         */
        export interface SearchOptionsLoadOptions {
            $all?: boolean;
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
        /**
         *
         * Represents a section in a Word document.
         *
         * [Api set: WordApi 1.1]
         */
        export interface SectionLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the body object of the section. This does not include the header/footer and other section metadata.
            *
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyLoadOptions;
        }
        /**
         *
         * Contains the collection of the document's {@link Word.Section} objects.
         *
         * [Api set: WordApi 1.1]
         */
        export interface SectionCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the body object of the section. This does not include the header/footer and other section metadata.
            *
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyLoadOptions;
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